#!/usr/bin/env python3
# =============================================================================
#  LAPORAN PROGRAM PRIORITAS - tarik data dari Zendesk
#  Replikasi report Explore: "Laporan Program Prioritas - Automation" (#270701801)
# =============================================================================

import os

# ================================ PENGATURAN =================================
# Di GitHub, nilai rahasia diambil dari GitHub Secrets (environment variable).
# Kalau dijalankan di laptop, boleh juga diisi langsung di tanda kutip kedua.
ZENDESK_SUBDOMAIN = os.getenv("ZENDESK_SUBDOMAIN", "wartekhelp")
ZENDESK_EMAIL     = os.getenv("ZENDESK_EMAIL", "putra.perdana@teknologi.belajar.id")
ZENDESK_API_TOKEN = os.getenv("ZENDESK_API_TOKEN", "9QyvKRi97ZoYkBwqZugmzLG735Hv65qmzZccbbMc")

# URL Web App Apps Script (berakhiran /exec) dan kata rahasianya
# (harus sama dengan SECRET_KEY di Apps Script)
APPS_SCRIPT_URL = os.getenv("APPS_SCRIPT_URL", "https://script.google.com/a/macros/teknologi.belajar.id/s/AKfycbw_9S71IQGKOp4e-5pqOSjlH-ZmrLZrJcJQCty1X5WdNCKcJcWSQYMm2G9L8yQ3eeOdsA/exec")
APPS_SCRIPT_KEY = os.getenv("APPS_SCRIPT_KEY", "program_prioritas")

# Simpan juga salinan Excel? (True / False)
SIMPAN_JUGA_KE_EXCEL = False
FOLDER_OUTPUT = "."
# =============================================================================

"""
Cara pakai
----------
A. Lewat GitHub Actions (disarankan): lihat README.md di repo. Atur config di tab
   Config Google Sheet, lalu klik menu "Zendesk Sync > Jalankan tarik data".

B. Lewat Terminal:
   1. python3 -m pip install --user -r requirements.txt
   2. Isi ZENDESK_API_TOKEN, APPS_SCRIPT_URL, APPS_SCRIPT_KEY di atas.
   3. python3 laporan_program_prioritas.py

Periode, Write Mode (REPLACE/APPEND), dan Target Sheet Name diambil dari tab Config.
Data masuk ke sheet tujuan (kolom A:F) dengan header
    Ticket created - Date | Ticket ID | Program Prioritas | Cakupan Program | Kategori Kendala | Tickets
dan Last Sync Status / Total Rows Synced / Last Sync Timestamp di tab Config ikut diperbarui.

Logika yang direplikasi dari Explore
------------------------------------
- Program Prioritas      = field "Program Prioritas" (56574930913433), yang kosong dibuang
- Cakupan Program        = field "Program Prioritas - Cakupan Program" (59325718816921),
                           hanya nilai yang ada di CAKUPAN_PROGRAM_DIPILIH
- Kategori Kendala       = calculated attribute "Cat-all_(R)", rumusnya diterjemahkan
                           di fungsi cat_all() di bawah; NULL dan "Tiket Spam" dibuang
- Tanggal                = tanggal tiket dibuat, zona waktu Asia/Jakarta

Catatan: API Search Zendesk tidak memuat tiket yang sudah diarsip (tiket closed
lebih dari ~120 hari). Untuk periode beberapa bulan terakhir ini tidak berpengaruh.
"""

import os
import re
import sys
import time
from collections import Counter
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import requests

TZ = ZoneInfo("Asia/Jakarta")

# ---- Field ID Zendesk -------------------------------------------------------
F_PROGRAM_PRIORITAS = 56574930913433   # Program Prioritas
F_CAKUPAN_PROGRAM   = 59325718816921   # Program Prioritas - Cakupan Program
F_SPAM              = 41865821908121   # Apakah Tiket Spam atau Bukan?
F_RUMDIK_CAT        = 52601223718809   # Rumdik-cat_all
F_ULT_LAYANAN       = 41797113781785   # ULT - Layanan Kemendikdasmen
F_ULT_GTK           = 41679556093081   # ULT - Layanan Guru dan Tenaga Kependidikan
F_ULT_TEKDATA       = 41717086778521   # ULT - Layanan Teknologi dan Data Pendidikan
F_ULT_PKK           = 41677142192793   # ULT - Layanan Pembelajaran, Kurikulum dan Kebahasaan
F_ULT_KPKP          = 41572042274329   # ULT - Layanan Kerjasama, Pembiayaan, Kepegawaian dan Pengawasan
F_ULT_RUMDIK        = 45788213384729   # ULT - Rumah Pendidikan
F_ULT_KATEGORI      = 11227582343321   # ULT - Kategori Masalah
F_GURU_CAT          = 23993947939993   # Guru-cat_all
F_GURU_CAT_V2       = 41838938060569   # Guru-cat_all_v2
F_RKMM_CAT_V2       = 38314519366169   # RKMM-cat_all-v2
F_RKMM_CAT          = 22876967333657   # RKMM-cat_all
F_BID_CAT           = 24675392323097   # BID-cat_all
F_RAPOR_CAT         = 21396791683865   # Rapor-cat_all
F_ARKAS_CAT         = 18999366586905   # SPM-ARKAS-cat_all
F_TANYABOSP_CAT     = 10328881445145   # SPM-TanyaBOSP-cat_all
F_SIPLAH_TUJUAN     = 6498065509017    # SPM-SIPLah-Pertanyaan ditujukan kepada
F_SIPLAH_CAT        = 19031235176345   # SPM-SIPLah-cat_all
F_SIPLAH_MITRA_CAT  = 19031241379609   # SPM-SIPLah_mitra-cat_all
F_SATUDATA_CAT      = 41373460981785   # Satu_Data-cat-all
F_RUANGMURID_CAT    = 45078520814617   # Ruang_Murid - cat_all
F_PIC_ESKALASI      = 34671204324633   # PIC Narahubung - Tiket Eskalasi PIC Narahubung?
F_PIC_ISU           = 34674009891097   # PIC Narahubung - Isu Kategorisasi (text)
F_SPM_PROGRAM       = 34021530742553   # SPM - Program
F_HIEDU_SSC_PROGRAM = 14718324990361   # HiEdu - SSC - Program
F_HIEDU_PROGRAM     = 4414041923993    # HiEdu - Program
F_TATAPMUKA_JENIS   = 28476041338009   # ULT Tatap Muka - Jenis Layanan
F_KSPS_CAT          = 20508998704537   # KSPS-cat_all
F_KSPS_KS_CAT       = 26612614538137   # KSPS-Fitur_Seleksi Kepala Sekolah-cat_all
F_KSPS_PS_CAT       = 26612683229465   # KSPS-Fitur_Seleksi Pengawas Sekolah-cat_all
F_LMS_PROGRAM       = 22203271054617   # LMS-Program
F_LMS_CAT           = 20584361976217   # LMS-cat_all

# LMS-Program -> field kategori (urutan sama dengan rumus "LMS-Category_(R)")
LMS_PROGRAM_KE_FIELD = [
    ("Aksi Nyata Sejawat", 24669195563289),
    ("Bimtek Peningkatan Kompetensi Guru Informatika SMP", 22203997466009),
    ("Bimtek Penyusunan Perangkat Ajar", 22204009167001),
    ("Cerita Praktik", 24669116778905),
    ("Diklat Teknis Mendongeng", 26611331264793),
    ("Diklat Teknis Perkembangan Otak Anak", 26611371370777),
    ("Informatika SMA", 25180711470489),
    ("Pendidikan Guru Penggerak", 24669219049113),
    ("Peningkatan Kompetensi Kepala Sekolah", 23815754664601),
    ("Penyusunan Modul Ajar dan Modul Projek", 22203819148825),
    ("PPG Prajabatan", 23128739948313),
    ("Transisi PAUD-SD", 22203888543129),
    ("Program Pengembangan Keprofesian Guru PJOK", 31694649655577),
    ("Diklat Refreshment PIP", 33079980336665),
    ("Pelatihan Komite Pembelajaran (PKP)", 34785083893913),
    ("Diklat Pengawas Sekolah", 34182724692505),
    ("Bimbingan Teknis Seri Mengenal Anak Kesulitan Belajar", 35706776288025),
    ("Rekognisi Pembelajaran Lampau Pendidikan Guru Penggerak Kepala SMK PK 2024", 36162078093593),
    ("LMS Diklat Upskilling Reskilling", 38222217488281),
    ("LMS Pendidikan Berjenjang Pendidikan Inklusif", 39650004050585),
    ("LMS Bimtek dan Pelatihan Koding dan Kecerdasan Artifisial", 48100894243865),
    ("LMS MOOC Balai Bidang Otomotif dan Elektronik", 43681221506201),
    ("LMS Pembelajaran Mendalam", 48442527549593),
    ("LMS Bina Talenta Indonesia", 50765655925657),
    ("LMS Jelajah Mandiri Rumah Pendidikan", 51144520675481),
    ("LMS Tenaga Kependidikan (TENDIK)", 51144459848857),
    ("LMS Pelatihan Bakal Calon Kepala Sekolah (CKS)", 51144485296537),
    ("LMS Matematika Gembira", 51144474062233),
    ("LMS Kesehatan Reproduksi Tingkat SLB", 51144533634329),
    ("LMS Pengembangan Kompetensi Bimbingan Konseling (BK)", 51144482788761),
]

# ULT - Layanan Kemendikdasmen -> field kategori
ULT_LAYANAN_KE_FIELD = {
    "Layanan Guru dan Tenaga Kependidikan": F_ULT_GTK,
    "Layanan Teknologi dan Data Pendidikan": F_ULT_TEKDATA,
    "Layanan Pembelajaran, Kurikulum dan Kebahasaan": F_ULT_PKK,
    "Layanan Kerjasama, Pembiayaan, Kepegawaian dan Pengawasan": F_ULT_KPKP,
    "Layanan Rumah Pendidikan": F_ULT_RUMDIK,
}

# Filter Cakupan Program (sama dengan yang dicentang di report Explore)
CAKUPAN_PROGRAM_DIPILIH = {
    "Asesmen Nasional",
    "Bantuan Rumah Guru",
    "Digitalisasi Pembelajaran",
    "Gerakan 7 Kebiasaan Anak Indonesia Hebat",
    "Gerakan Numerasi Nasional",
    "Koding dan Kecerdasan Artifisial",
    "Pembangunan Bahasa dan Sastra",
    "Pembelajaran Mendalam",
    "Pemenuhan Kualifikasi Akademik S-1/D-IV",
    "Pendampingan Program Kerja Pokja Bunda PAUD",
    "Pendidikan Profesi Guru",
    "Penguatan Literasi Kebahasaan dan Kesastraan",
    "Peningkatan Kapasitas UPT dalam melakukan Advokasi dan Pendampingan Wajar 1 Tahun Prasekolah",
    "Peningkatan Kompetensi Guru BK",
    "Penyusunan Bahan Publikasi Wajib Belajar 1 Tahun Prasekolah",
    "Penyusunan Desain Besar Wajar 1 Tahun Prasekolah",
    "Penyusunan Regulasi Penguatan Bunda PAUD",
    "Program Kepemimpinan Sekolah (Pelatihan BCKS)",
    "Rapat Koordinasi Desain Besar Wajib Belajar 1 Tahun Prasekolah bersama Lintas K/L serta FGD Percepatan Wajib Belajar 1 Tahun Prasekolah bersama LPTK",
    "Revitalisasi Satuan Pendidikan",
    "Sekolah Rakyat",
    "Sosialisasi Wajib Belajar 1 Tahun Prasekolah kepada 514 Dinas Pendidikan dan Pokja Bunda PAUD",
    "Tes Kemampuan Akademik",
    "Tunjangan Guru Aparatur Sipil Negara Daerah (ASND)",
    "Tunjangan Guru, Bantuan Insentif, dan Bantuan Subsidi Upah (BSU) bagi Guru Non-ASN",
}

KATEGORI_DIBUANG = {"Tiket Spam"}   # + nilai kosong (NULL)

BARIS_PER_KIRIM = 2000   # jumlah baris per kiriman ke Google Sheet

KOLOM = ["Ticket created - Date", "Ticket ID", "Program Prioritas",
         "Cakupan Program", "Kategori Kendala", "Tickets"]


# =============================================================================
#  Koneksi Zendesk
# =============================================================================
class Zendesk:
    def __init__(self, subdomain, email, token):
        self.base = f"https://{subdomain}.zendesk.com"
        self.s = requests.Session()
        self.s.auth = (f"{email}/token", token)

    def get(self, url, params=None):
        if url.startswith("/"):
            url = self.base + url
        for attempt in range(8):
            try:
                r = self.s.get(url, params=params, timeout=120)
            except requests.RequestException as e:
                print(f"  Koneksi gagal ({e.__class__.__name__}), coba lagi...")
                time.sleep(5 * (attempt + 1))
                continue
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 60))
                print(f"  Kena rate limit Zendesk, tunggu {wait} detik...")
                time.sleep(wait)
                continue
            if r.status_code >= 500:
                time.sleep(5 * (attempt + 1))
                continue
            if r.status_code == 401:
                sys.exit("Login gagal (401). Cek ZENDESK_EMAIL dan ZENDESK_API_TOKEN.")
            if r.status_code == 403:
                sys.exit("Akses ditolak (403). Cek apakah Token access sudah aktif di Admin Center.")
            r.raise_for_status()
            return r.json()
        sys.exit(f"Gagal mengambil data dari {url} setelah beberapa kali percobaan.")

    def paginate(self, url, key):
        params = {"page[size]": 100}
        while url:
            d = self.get(url, params)
            yield from d.get(key, [])
            params = None
            url = d["links"]["next"] if d.get("meta", {}).get("has_more") else None

    def search_export(self, query):
        params = {"query": query, "filter[type]": "ticket", "page[size]": 1000}
        url = "/api/v2/search/export.json"
        while url:
            d = self.get(url, params)
            yield from d.get("results", [])
            params = None
            url = d["links"]["next"] if d.get("meta", {}).get("has_more") else None


# =============================================================================
#  Akses nilai field per tiket (tag dropdown -> nama yang tampil di Explore)
# =============================================================================
class Tiket:
    def __init__(self, raw, opsi, brand_names, form_names):
        self.raw = raw
        self.brand = brand_names.get(raw.get("brand_id"))
        self.form = form_names.get(raw.get("ticket_form_id"))
        self._cf = {c["id"]: c["value"] for c in raw.get("custom_fields") or []}
        self._opsi = opsi

    def f(self, field_id):
        """Nilai field seperti di Explore; None kalau kosong."""
        v = self._cf.get(field_id)
        if v in (None, "", []):
            return None
        if isinstance(v, list):
            v = v[0]
        if isinstance(v, str):
            return self._opsi.get(field_id, {}).get(v, v)
        return str(v)


# =============================================================================
#  Calculated attribute Explore -> Python
# =============================================================================
def brand_overall(t):
    """Brand Overall = Brand Overall 2, kalau kosong pakai Program Hiedu (PMM)."""
    kampus_merdeka = t.brand == "Mahasiswa & Perguruan Tinggi Kampus Merdeka"
    bo1 = t.f(F_HIEDU_PROGRAM) if kampus_merdeka else t.brand            # Brand Overall 1
    bo2 = t.f(F_HIEDU_SSC_PROGRAM) if bo1 is None else bo1               # Brand Overall 2
    if bo2 is not None:
        return bo2
    if kampus_merdeka and t.f(F_HIEDU_PROGRAM) is None:                  # Program Hiedu (PMM)
        return "Pertukaran Mahasiswa Merdeka"
    return t.f(F_HIEDU_PROGRAM)


def layanan_tatap_muka(t):
    v = t.f(F_TATAPMUKA_JENIS)
    if v is not None and "::" in v:
        m = re.search(r".*::([^:]+)$", v)
        return m.group(1) if m else None
    return v


def category_tatap_muka(t):
    ltm = layanan_tatap_muka(t)
    if t.brand == "Layanan Tatap Muka":
        peta = {"Platform Merdeka Mengajar": F_GURU_CAT, "Akun Belajar.id": F_BID_CAT,
                "ARKAS": F_ARKAS_CAT, "SIPLah": F_SIPLAH_CAT, "TanyaBOSP": F_TANYABOSP_CAT}
        if ltm in peta:
            return t.f(peta[ltm])
    return ltm


def ksps_category(t):
    bagian = [t.f(F_KSPS_CAT), t.f(F_KSPS_PS_CAT), t.f(F_KSPS_KS_CAT)]
    hasil = "".join(b for b in bagian if b is not None)
    return hasil or None


def lms_category(t):
    program = t.f(F_LMS_PROGRAM)
    if program is None:
        return t.f(F_LMS_CAT)
    for nama, fid in LMS_PROGRAM_KE_FIELD:
        if program == nama:
            return t.f(fid)
    return None


def cat_all(t):
    """Terjemahan rumus calculated attribute Explore 'Cat-all_(R)' (Category Check (R))."""
    f, form = t.f, t.form
    bo = brand_overall(t)
    ult = f(F_ULT_LAYANAN)
    rumdik = f(F_RUMDIK_CAT)
    guru, guru_v2 = f(F_GURU_CAT), f(F_GURU_CAT_V2)
    rkmm, rkmm_v2 = f(F_RKMM_CAT), f(F_RKMM_CAT_V2)

    if f(F_SPAM) == "Ya, Tiket Spam":
        return "Tiket Spam"
    if bo in ("Rumah Pendidikan", "ULT Kemendikdasmen") and rumdik is not None:
        return rumdik

    if form in ("Formulir Layanan Kemendikdasmen", "Formulir Sosial Media Kemendikdasmen RI"):
        if ult in ULT_LAYANAN_KE_FIELD:
            return f(ULT_LAYANAN_KE_FIELD[ult])
        if ult is None:
            return f(F_ULT_KATEGORI)

    if form == "Formulir Ruang GTK":
        return guru_v2 if guru_v2 is not None else guru

    if form == "Formulir Ruang Kolaborasi Merdeka Mengajar":
        if rkmm_v2 is None:
            return rkmm
        if rkmm is None:
            return rkmm_v2
    if form == "Ruang Kolaborasi Merdeka Belajar" and rkmm is not None and rkmm_v2 is not None:
        return rkmm_v2

    if form == "Formulir Sistem Pengangkatan KSPS":
        return ksps_category(t)
    if form == "Formulir LMS Guru":
        return lms_category(t)
    # Rumus asli: ELIF [Rumdik - Program] = "Rumdik - Ruang GTK" -> tidak pernah terpenuhi
    # karena "Rumdik - Program" tidak pernah menghasilkan nilai itu, jadi dilewati.

    if form == "Formulir Akun belajar.id":
        return f(F_BID_CAT)
    if form == "Formulir Rapor Pendidikan":
        return f(F_RAPOR_CAT)
    if form == "Formulir ARKAS":
        return f(F_ARKAS_CAT)
    if form == "Formulir SDS" and bo == "ARKAS":
        return f(F_ARKAS_CAT)
    if form == "Formulir SDS" and bo in ("BOSP", "TanyaBOSP"):
        return f(F_TANYABOSP_CAT)
    if form in ("Formulir SDS", "Formulir SIPLah") and bo == "SIPLah":
        tujuan = f(F_SIPLAH_TUJUAN)
        if tujuan == "Tim SIPLah Kemendikbud":
            return f(F_SIPLAH_CAT)
        if tujuan == "Mitra SIPLah":
            return f(F_SIPLAH_MITRA_CAT)

    if form == "Formulir Agent SPM" and ult in ULT_LAYANAN_KE_FIELD:
        return f(ULT_LAYANAN_KE_FIELD[ult])
    if form == "Formulir Layanan Tatap Muka":
        return category_tatap_muka(t)
    if form == "Formulir Portal Satu Data":
        return f(F_SATUDATA_CAT)
    if form == "Formulir Ruang Murid":
        return f(F_RUANGMURID_CAT)
    if f(F_PIC_ESKALASI) == "Ya, Eskalasi PIC Narahubung":
        return f(F_PIC_ISU)

    if form in ("Formulir Rumah Pendidikan", "Formulir Layanan Kemendikdasmen",
                "Formulir Agent SPM", "Formulir Sosial Media Kemendikdasmen RI",
                "Formulir Ruang GTK") and rumdik is not None:
        return rumdik
    if bo == "ULT Kemendikdasmen" and ult is not None:
        return ult
    if bo == "Rumah Pendidikan" and rumdik is not None:
        return rumdik
    if bo == "Akun belajar.id" and f(F_BID_CAT) is not None:
        return f(F_BID_CAT)
    if bo == "Ruang GTK" and guru_v2 is not None:
        return guru_v2
    if bo == "Ruang GTK" and guru is not None:
        return guru
    if bo == "Layanan ULT Kemendikdasmen" and ult is not None:
        return ult
    if bo == "Layanan Tatap Muka" and category_tatap_muka(t) is not None:
        return category_tatap_muka(t)
    if f(F_SPM_PROGRAM) == "ARKAS" and f(F_ARKAS_CAT) is not None:
        return f(F_ARKAS_CAT)
    return None



# =============================================================================
#  Google Sheet (lewat Apps Script Web App)
# =============================================================================
class Sheet:
    def __init__(self, url, key):
        self.url, self.key = url.strip(), key
        self.s = requests.Session()

    def _cek(self, r):
        r.raise_for_status()
        try:
            d = r.json()
        except ValueError:
            judul = re.search(r"<title>(.*?)</title>", r.text, re.S | re.I)
            judul = judul.group(1).strip() if judul else r.text[:120].strip()
            alamat = r.url.split("?")[0]
            if "/exec" not in self.url:
                sebab = "APPS_SCRIPT_URL bukan URL Web App (harus .../macros/s/.../exec)."
            elif "accounts.google.com" in r.url or "ServiceLogin" in r.url or "Sign in" in judul:
                sebab = ("Web App minta login Google -> ubah 'Who has access' jadi 'Anyone' "
                         "(Deploy > Manage deployments > Edit).")
            else:
                sebab = "Cek deployment Web App (Execute as: Me, Who has access: Anyone)."
            raise RuntimeError(
                f"Balasan Apps Script bukan JSON. {sebab}\n"
                f"  HTTP {r.status_code} dari {alamat}\n  Judul halaman: {judul}")
        if not d.get("ok"):
            raise RuntimeError(f"Apps Script: {d.get('error')}")
        return d

    def get(self, action):
        return self._cek(self.s.get(self.url, params={"action": action, "key": self.key}, timeout=120))

    def post(self, payload):
        payload = dict(payload, key=self.key)
        for attempt in range(4):
            try:
                return self._cek(self.s.post(self.url, json=payload, timeout=300))
            except requests.RequestException:
                if attempt == 3:
                    raise
                time.sleep(5 * (attempt + 1))

    def status(self, teks, rows=None):
        try:
            self.post({"action": "status", "status": teks, "rows": rows})
        except Exception as e:  # status gagal jangan sampai menghentikan proses
            print(f"  (Gagal update status di Sheet: {e})")


# =============================================================================
#  Main
# =============================================================================
def parse_tanggal(s):
    return datetime.strptime(s.strip(), "%Y-%m-%d").date()


def proses_tiket(raw, opsi, brand_names, form_names, start_local, end_local, stat):
    """Kembalikan satu baris output, atau None kalau tiket tidak masuk report."""
    created = datetime.fromisoformat(raw["created_at"].replace("Z", "+00:00")).astimezone(TZ)
    if not (start_local <= created < end_local):
        return None
    stat["tiket_dalam_periode"] += 1
    t = Tiket(raw, opsi, brand_names, form_names)

    program = t.f(F_PROGRAM_PRIORITAS)
    if program is None:
        stat["dibuang: Program Prioritas kosong"] += 1
        return None
    cakupan = t.f(F_CAKUPAN_PROGRAM)
    if cakupan not in CAKUPAN_PROGRAM_DIPILIH:
        stat["dibuang: Cakupan Program kosong/tidak dipilih"] += 1
        if cakupan is not None:
            stat[f"  (Cakupan tidak ada di daftar) {cakupan}"] += 1
        return None
    kategori = cat_all(t)
    if kategori is None:
        stat["dibuang: Kategori Kendala kosong"] += 1
        return None
    if kategori in KATEGORI_DIBUANG:
        stat["dibuang: Tiket Spam"] += 1
        return None

    return {
        "Ticket created - Date": created.date().isoformat(),
        "Ticket ID": raw["id"],
        "Program Prioritas": program,
        "Cakupan Program": cakupan,
        "Kategori Kendala": kategori,
        "Tickets": 1,
    }


def tarik_data(start_d, end_d):
    """Tarik tiket dari Zendesk untuk periode (WIB, inklusif). Kembalikan DataFrame."""
    start_local = datetime.combine(start_d, datetime.min.time(), TZ)
    end_local = datetime.combine(end_d + timedelta(days=1), datetime.min.time(), TZ)
    utc = ZoneInfo("UTC")

    zd = Zendesk(ZENDESK_SUBDOMAIN, ZENDESK_EMAIL, ZENDESK_API_TOKEN)

    print("Mengambil daftar field, brand, dan form dari Zendesk...")
    opsi = {}
    for fld in zd.paginate("/api/v2/ticket_fields.json", "ticket_fields"):
        opts = fld.get("custom_field_options") or []
        if opts:
            opsi[fld["id"]] = {o["value"]: o["name"] for o in opts}
    brand_names = {b["id"]: b["name"] for b in zd.paginate("/api/v2/brands.json", "brands")}
    form_names = {f["id"]: f["name"] for f in zd.get("/api/v2/ticket_forms.json").get("ticket_forms", [])}

    print("Menarik tiket (yang Program Prioritas-nya terisi)...")
    query = (f"type:ticket "
             f"created>={start_local.astimezone(utc):%Y-%m-%dT%H:%M:%SZ} "
             f"created<{end_local.astimezone(utc):%Y-%m-%dT%H:%M:%SZ} "
             f"custom_field_{F_PROGRAM_PRIORITAS}:*")
    stat = Counter()
    rows, n = [], 0
    for raw in zd.search_export(query):
        n += 1
        if n % 1000 == 0:
            print(f"  {n:,} tiket diproses...")
        row = proses_tiket(raw, opsi, brand_names, form_names, start_local, end_local, stat)
        if row:
            rows.append(row)

    df = pd.DataFrame(rows, columns=KOLOM)
    if not df.empty:
        df = df.drop_duplicates("Ticket ID").sort_values(["Ticket created - Date", "Ticket ID"])

    print("\nRingkasan:")
    print(f"  Tiket dengan Program Prioritas : {stat['tiket_dalam_periode']:,}")
    for k, v in sorted(stat.items()):
        if k != "tiket_dalam_periode":
            print(f"  {k}: {v:,}")
    print(f"  Baris masuk laporan            : {len(df):,}")
    return df


def kirim_ke_sheet(sheet, df, target, mode):
    """Kirim data per potongan. REPLACE: potongan pertama menghapus data lama dulu."""
    rows = [[r[0], int(r[1]), r[2], r[3], r[4], int(r[5])] for r in df.itertuples(index=False)]
    total = len(rows)
    if total == 0:
        if mode == "REPLACE":
            sheet.post({"action": "write", "target_sheet": target, "reset": True, "rows": []})
        return 0
    terkirim = 0
    for i in range(0, total, BARIS_PER_KIRIM):
        potong = rows[i:i + BARIS_PER_KIRIM]
        sheet.post({"action": "write", "target_sheet": target,
                    "reset": (mode == "REPLACE" and i == 0), "rows": potong})
        terkirim += len(potong)
        print(f"  {terkirim:,} / {total:,} baris terkirim")
    return terkirim


def simpan_excel(df, start_d, end_d):
    out = os.path.join(FOLDER_OUTPUT, f"Program_Prioritas_{start_d:%Y%m%d}_{end_d:%Y%m%d}.xlsx")
    os.makedirs(FOLDER_OUTPUT, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="Cakupan Program Prioritas", index=False)
    print(f"Salinan Excel: {os.path.abspath(out)}")


def main():
    if not ZENDESK_API_TOKEN or ZENDESK_API_TOKEN == "TEMPEL_API_TOKEN_DI_SINI":
        sys.exit("ZENDESK_API_TOKEN belum diisi (GitHub Secret atau bagian atas script).")
    if not APPS_SCRIPT_URL or "TEMPEL" in APPS_SCRIPT_URL:
        sys.exit("APPS_SCRIPT_URL belum diisi (GitHub Secret atau bagian atas script).")

    sheet = Sheet(APPS_SCRIPT_URL, APPS_SCRIPT_KEY)

    print("Membaca config dari Google Sheet...")
    try:
        cfg = sheet.get("config")
    except Exception as e:
        sys.exit(f"Gagal membaca config: {e}")

    start_d, end_d = parse_tanggal(cfg["start_date"]), parse_tanggal(cfg["end_date"])
    mode, target = cfg["mode"], cfg["target_sheet"]
    info_target = (f"sudah ada, {cfg['target_data_rows']:,} baris data" if cfg["target_exists"]
                   else "belum ada, akan dibuat")
    print(f"""
  Spreadsheet : {cfg.get('spreadsheet', '')}
  Periode     : {start_d:%d %b %Y} s.d. {end_d:%d %b %Y} (WIB)
  Write Mode  : {mode} {'(data lama di sheet tujuan akan DIHAPUS)' if mode == 'REPLACE' else '(ditambahkan di bawah data terakhir)'}
  Sheet tujuan: {target} ({info_target})
""")
    interaktif = sys.stdin.isatty() and not os.getenv("GITHUB_ACTIONS")
    if interaktif and input("Lanjut? [Y/n]: ").strip().lower() in ("n", "no", "tidak", "t"):
        print("Dibatalkan.")
        return

    periode = f"{start_d:%Y-%m-%d} s/d {end_d:%Y-%m-%d}"
    sheet.status(f"⏳ Sedang berjalan ({periode})")
    try:
        df = tarik_data(start_d, end_d)
        print(f"\nMengirim ke sheet '{target}' ({mode})...")
        n = kirim_ke_sheet(sheet, df, target, mode)
    except BaseException as e:
        pesan = "dibatalkan" if isinstance(e, KeyboardInterrupt) else str(e) or e.__class__.__name__
        sheet.status(f"❌ Gagal ({periode}): {pesan}"[:300])
        if isinstance(e, SystemExit):
            raise
        sys.exit(f"\nGagal: {pesan}")

    sheet.status(f"✅ Sukses ({periode})", rows=n)
    if SIMPAN_JUGA_KE_EXCEL:
        simpan_excel(df, start_d, end_d)
    print(f"\nSelesai. {n:,} baris masuk ke sheet '{target}'.")


if __name__ == "__main__":
    main()
