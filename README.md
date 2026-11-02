# Laporan Program Prioritas (Zendesk → Google Sheet)

Menarik data tiket Zendesk sesuai report Explore **"Laporan Program Prioritas - Automation"**
lalu menuliskannya ke Google Sheet. Dijalankan oleh GitHub Actions, dipicu dari menu Google Sheet.

## Isi repo
| File | Fungsi |
|---|---|
| `laporan_program_prioritas.py` | Tarik data Zendesk, terapkan logika Explore, kirim ke Sheet |
| `requirements.txt` | Library Python yang dibutuhkan |
| `.github/workflows/sync.yml` | Workflow GitHub Actions yang menjalankan script |

## GitHub Secrets yang wajib diisi
Settings → Secrets and variables → Actions → **New repository secret**

| Nama | Isi |
|---|---|
| `ZENDESK_API_TOKEN` | API token Zendesk |
| `APPS_SCRIPT_URL` | URL Web App Apps Script (berakhiran `/exec`) |
| `APPS_SCRIPT_KEY` | Kata rahasia, sama dengan `SECRET_KEY` di Apps Script |

## Cara pakai
1. Atur **Start Date**, **End Date**, **Write Mode** (REPLACE/APPEND), **Target Sheet Name** di tab Config.
2. Klik menu **Zendesk Sync → ▶ Jalankan tarik data** di Google Sheet.
3. Tunggu 1–3 menit. **Last Sync Status** berubah jadi ✅ (atau ❌ dengan penyebabnya).

Bisa juga dijalankan langsung dari GitHub: tab **Actions → Tarik Laporan Program Prioritas → Run workflow**.
Log lengkap tiap run ada di tab Actions.
