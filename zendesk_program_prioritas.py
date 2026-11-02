#!/usr/bin/env python3
"""
Tarik data tiket Zendesk -> laporan "Cakupan Program Prioritas".

Output (sama dengan export Explore):
    Ticket created - Date | Ticket ID | Program Prioritas | Cakupan Program | Kategori Kendala | Tickets

Cara pakai
----------
1) Install dependensi (sekali saja):
       pip install requests pandas openpyxl

2) Isi kredensial lewat environment variable (atau isi langsung di bagian KONFIGURASI di bawah):
       ZENDESK_SUBDOMAIN   -> contoh: "namaakun" untuk https://namaakun.zendesk.com
       ZENDESK_EMAIL       -> email agent/admin
       ZENDESK_API_TOKEN   -> Admin Center > Apps and integrations > Zendesk API > API tokens
   Kalau tidak diisi, script akan menanyakannya saat dijalankan.

3) Jalankan:
       python zendesk_program_prioritas.py --start 2026-10-01 --end 2026-10-28
   atau tanpa argumen, nanti ditanya tanggalnya:
       python zendesk_program_prioritas.py

   Opsi lain:
       --output nama_file.xlsx   nama file hasil (default otomatis sesuai periode)
       --csv                     simpan juga versi CSV
       --include-empty           ikutkan tiket yang Cakupan Program-nya kosong
       --list-fields             tampilkan daftar ticket field (untuk cek nama/ID field)

Catatan
-------
- Periode dihitung berdasarkan tanggal tiket dibuat di zona waktu Asia/Jakarta,
  dan tanggal akhir ikut dihitung (inklusif).
- Field yang multi-select (mis. Cakupan Program) dipecah jadi beberapa baris,
  sama seperti di Explore.
- Nilai field diterjemahkan dari tag ke nama yang tampil (mis. "TKA::Informasi umum::Cara akses").
"""

import argparse
import getpass
import itertools
import os
import sys
import time
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import requests

# ============================== KONFIGURASI ==============================
ZENDESK_SUBDOMAIN = os.getenv("ZENDESK_SUBDOMAIN", "")
ZENDESK_EMAIL = os.getenv("ZENDESK_EMAIL", "")
ZENDESK_API_TOKEN = os.getenv("ZENDESK_API_TOKEN", "")

TIMEZONE = "Asia/Jakarta"

# Nama kolom di output -> judul ticket field di Zendesk.
# Kalau judul field di Zendesk berbeda, ganti di sini, atau isi ID field-nya
# (angka) supaya pasti cocok. Cek pakai: python zendesk_program_prioritas.py --list-fields
FIELDS = {
    "Program Prioritas": "Program Prioritas",
    "Cakupan Program": "Cakupan Program",
    "Kategori Kendala": "Kategori Kendala",
}

# Baris hanya diambil kalau kolom ini terisi (seperti file Explore aslinya).
REQUIRED_COLUMN = "Cakupan Program"

SHEET_NAME = "Cakupan Program Prioritas"
# ========================================================================


class Zendesk:
    def __init__(self, subdomain, email, token):
        self.base = f"https://{subdomain}.zendesk.com"
        self.s = requests.Session()
        self.s.auth = (f"{email}/token", token)
        self.s.headers["Accept"] = "application/json"

    def get(self, url, params=None):
        if url.startswith("/"):
            url = self.base + url
        for attempt in range(6):
            r = self.s.get(url, params=params, timeout=120)
            if r.status_code == 429:  # rate limit
                wait = int(r.headers.get("Retry-After", 60))
                print(f"  Kena rate limit, tunggu {wait} detik...")
                time.sleep(wait)
                continue
            if r.status_code >= 500:
                time.sleep(5 * (attempt + 1))
                continue
            if r.status_code == 401:
                sys.exit("Gagal login (401). Cek email dan API token.")
            if r.status_code == 403:
                sys.exit("Akses ditolak (403). Pastikan akun punya akses ke tiket/API.")
            r.raise_for_status()
            return r.json()
        r.raise_for_status()

    def ticket_fields(self):
        fields, url = [], "/api/v2/ticket_fields.json"
        while url:
            data = self.get(url)
            fields += data["ticket_fields"]
            url = data.get("next_page")
        return fields

    def tickets_created_between(self, start_utc, end_utc):
        """Search Export API (cursor) - tidak dibatasi 1.000 hasil seperti Search biasa."""
        q = (
            f"type:ticket created>={start_utc:%Y-%m-%dT%H:%M:%SZ} "
            f"created<{end_utc:%Y-%m-%dT%H:%M:%SZ}"
        )
        params = {"query": q, "filter[type]": "ticket", "page[size]": 1000}
        url = "/api/v2/search/export.json"
        while True:
            data = self.get(url, params)
            yield from data.get("results", [])
            if not data.get("meta", {}).get("has_more"):
                break
            url = data["links"]["next"]
            params = None  # cursor sudah ada di URL next


def ask(prompt, default=""):
    val = input(f"{prompt}{f' [{default}]' if default else ''}: ").strip()
    return val or default


def parse_date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        sys.exit(f"Format tanggal salah: {s!r}. Gunakan YYYY-MM-DD, contoh 2026-10-01.")


def resolve_fields(all_fields):
    """Cocokkan FIELDS ke ticket field Zendesk; kembalikan {kolom: (field_id, {tag: nama})}."""
    by_id = {f["id"]: f for f in all_fields}
    by_title = {}
    for f in all_fields:
        by_title.setdefault(f["title"].strip().lower(), f)
        if f.get("raw_title"):
            by_title.setdefault(f["raw_title"].strip().lower(), f)

    resolved = {}
    for col, key in FIELDS.items():
        f = by_id.get(int(key)) if str(key).isdigit() else by_title.get(str(key).strip().lower())
        if not f:
            sys.exit(
                f"Ticket field untuk kolom '{col}' tidak ditemukan (dicari: {key!r}).\n"
                "Jalankan dengan --list-fields untuk melihat nama/ID field, lalu ubah FIELDS di script."
            )
        options = {
            o["value"]: o["name"]
            for o in (f.get("custom_field_options") or f.get("system_field_options") or [])
        }
        resolved[col] = (f["id"], options)
        print(f"  {col:<18} -> field #{f['id']} '{f['title']}' ({f['type']}, {len(options)} opsi)")
    return resolved


def field_values(ticket, field_id, options):
    """Ambil nilai field sebagai list nama tampilan (multi-select bisa >1)."""
    raw = None
    for cf in ticket.get("custom_fields") or []:
        if cf["id"] == field_id:
            raw = cf["value"]
            break
    if raw in (None, "", []):
        return [""]
    vals = raw if isinstance(raw, list) else [raw]
    return [options.get(v, str(v)) for v in vals]


def main():
    ap = argparse.ArgumentParser(description="Tarik data Program Prioritas dari Zendesk.")
    ap.add_argument("--start", help="Tanggal mulai (YYYY-MM-DD), inklusif")
    ap.add_argument("--end", help="Tanggal akhir (YYYY-MM-DD), inklusif")
    ap.add_argument("--output", help="Nama file .xlsx hasil")
    ap.add_argument("--csv", action="store_true", help="Simpan juga sebagai CSV")
    ap.add_argument("--include-empty", action="store_true",
                    help=f"Ikutkan tiket dengan '{REQUIRED_COLUMN}' kosong")
    ap.add_argument("--list-fields", action="store_true", help="Tampilkan daftar ticket field lalu keluar")
    args = ap.parse_args()

    subdomain = ZENDESK_SUBDOMAIN or ask("Zendesk subdomain (bagian sebelum .zendesk.com)")
    email = ZENDESK_EMAIL or ask("Email Zendesk")
    token = ZENDESK_API_TOKEN or getpass.getpass("API token Zendesk: ")
    zd = Zendesk(subdomain.replace(".zendesk.com", "").replace("https://", "").strip("/"), email, token)

    if args.list_fields:
        for f in zd.ticket_fields():
            print(f"{f['id']:>15}  {f['type']:<12}  {f['title']}")
        return

    # --- Periode ---
    today = date.today()
    start_s = args.start or ask("Tanggal mulai (YYYY-MM-DD)", today.replace(day=1).isoformat())
    end_s = args.end or ask("Tanggal akhir (YYYY-MM-DD)", (today - timedelta(days=1)).isoformat())
    start_d, end_d = parse_date(start_s), parse_date(end_s)
    if end_d < start_d:
        sys.exit("Tanggal akhir tidak boleh sebelum tanggal mulai.")

    tz = ZoneInfo(TIMEZONE)
    start_local = datetime.combine(start_d, datetime.min.time(), tz)
    end_local = datetime.combine(end_d + timedelta(days=1), datetime.min.time(), tz)  # eksklusif
    start_utc = start_local.astimezone(ZoneInfo("UTC"))
    end_utc = end_local.astimezone(ZoneInfo("UTC"))
    print(f"\nPeriode: {start_d} s.d. {end_d} ({TIMEZONE})")

    # --- Field mapping ---
    print("\nMengambil definisi ticket field...")
    fields = resolve_fields(zd.ticket_fields())

    # --- Tiket ---
    print("\nMenarik tiket...")
    rows, n_tickets = [], 0
    for t in zd.tickets_created_between(start_utc, end_utc):
        created = datetime.fromisoformat(t["created_at"].replace("Z", "+00:00")).astimezone(tz)
        if not (start_local <= created < end_local):
            continue
        n_tickets += 1
        if n_tickets % 1000 == 0:
            print(f"  {n_tickets:,} tiket diproses...")

        per_col = {col: field_values(t, fid, opts) for col, (fid, opts) in fields.items()}
        if not args.include_empty and per_col[REQUIRED_COLUMN] == [""]:
            continue
        cols = list(per_col)
        for combo in itertools.product(*(per_col[c] for c in cols)):
            row = {"Ticket created - Date": created.date().isoformat(), "Ticket ID": str(t["id"])}
            row.update(dict(zip(cols, combo)))
            row["Tickets"] = 1
            rows.append(row)

    columns = ["Ticket created - Date", "Ticket ID", *FIELDS.keys(), "Tickets"]
    df = pd.DataFrame(rows, columns=columns)
    if not df.empty:
        df = (df.drop_duplicates()
                .assign(_id=lambda d: d["Ticket ID"].astype(int))
                .sort_values(["Ticket created - Date", "_id", "Cakupan Program"])
                .drop(columns="_id"))

    print(f"\nTotal tiket dalam periode : {n_tickets:,}")
    print(f"Tiket dengan {REQUIRED_COLUMN:<12}: {df['Ticket ID'].nunique() if not df.empty else 0:,}")
    print(f"Jumlah baris output        : {len(df):,}")

    out = args.output or f"Program_Prioritas_{start_d:%Y%m%d}_{end_d:%Y%m%d}.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name=SHEET_NAME[:31], index=False)
        ws = xw.sheets[SHEET_NAME[:31]]
        for col_letter, width in zip("ABCDEF", (22, 12, 55, 60, 80, 9)):
            ws.column_dimensions[col_letter].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    print(f"\nTersimpan: {os.path.abspath(out)}")

    if args.csv:
        csv_out = os.path.splitext(out)[0] + ".csv"
        df.to_csv(csv_out, index=False, encoding="utf-8-sig")
        print(f"Tersimpan: {os.path.abspath(csv_out)}")


if __name__ == "__main__":
    main()
