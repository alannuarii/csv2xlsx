# csv2xlsx (CSV &bull; Excel Bidirectional Converter)

Aplikasi web monolit mandiri (*self-hosted*) yang cepat, aman, dan elegan untuk mengonversi data secara dua arah antara file terdelimitasi (`.csv`, `.tsv`, `.txt`) dan file spreadsheet Microsoft Excel (`.xlsx`) dengan deteksi otomatis format, penataan gaya profesional, dukungan multi-sheet, dan tabel pratinjau interaktif.

![csv2xlsx banner](app/static/favicon.svg)

---

## Fitur Utama

- **Konversi Dua Arah (CSV ⇄ XLSX)**:
  - **CSV ke Excel (.xlsx)**: Deteksi otomatis pemisah (koma `,`, titik koma `;`, tab `\t`, pipa `|`) dan *encoding* (`UTF-8`, `UTF-8-SIG`, `Latin-1`, `Windows-1252`), styling header kontras, auto-filter, auto-fit lebar kolom, dan smart typing.
  - **Excel (.xlsx) ke CSV**: Ekspor lembar kerja Excel ke file terdelimitasi dengan dukungan pemilihan sheet (*multi-sheet selector*), pilihan delimiter fleksibel (default: Koma `,`), dan pilihan encoding (default: `UTF-8 dengan BOM` untuk kompatibilitas sempurna dengan Microsoft Excel).
- **Pilihan Default Terbaik (*Best Defaults*)**: Sistem otomatis memilih konfigurasi terbaik jika pengguna tidak menentukan pilihan manual (pemilihan sheet aktif, delimiter standar `,`, dan encoding `utf-8-sig`).
- **Pratinjau Data Interaktif (*Data Preview*)**: Menampilkan kartu statistik (total baris, total kolom, delimiter, encoding / info sheet) dan pratinjau tabel 20 baris pertama sebelum file diunduh.
- **Deteksi Jenis File Otomatis**: Antarmuka secara cerdas mengenali file yang di-drop dan otomatis menyesuaikan mode konversi yang sesuai.
- **Smart Cell Typing**: Mengonversi teks angka dan boolean ke tipe data Excel asli, sekaligus menjaga angka dengan awalan nol (misal nomor HP `0812...` atau kode identitas `00123`) agar tetap terbaca sebagai teks tanpa merusaknya.
- **100% Privasi Server Sendiri (*Self-Hosted*)**: Seluruh pemrosesan dilakukan di server lokal/internal Anda tanpa mengirimkan data ke pihak ketiga. Berkas sementara otomatis dibersihkan.
- **Ringan & Tanpa Build Step (*Zero-Build Monolith*)**: Antarmuka SPA disajikan langsung oleh FastAPI menggunakan Tailwind CSS CDN dan Lucide Icons tanpa dependensi Node.js/npm.

---

## Arsitektur & Struktur Direktori

```text
csv2xlsx/
├── app/
│   ├── __init__.py
│   ├── main.py              # Routing FastAPI, validasi request, static file mount
│   ├── converter.py         # Engine konversi CSV ⇄ XLSX, openpyxl reader/writer & styling
│   └── static/
│       ├── favicon.svg      # Favicon format spreadsheet
│       └── index.html       # Antarmuka SPA lengkap dua arah (Tailwind, Lucide, Vanilla JS)
├── tests/
│   ├── __init__.py
│   ├── test_converter.py    # Unit tests untuk logic converter CSV & XLSX
│   └── test_api.py          # Integration tests untuk endpoint FastAPI
├── .dockerignore
├── .gitignore
├── Dockerfile               # Base python:3.11-slim, non-root user
├── docker-compose.yml       # Definisi service port 8000:8000
├── Jenkinsfile              # Pipeline CI/CD build & automated test
├── requirements.txt         # fastapi, uvicorn, openpyxl, charset-normalizer, dsb.
├── PRD.md                   # Product Requirements Document
└── README.md                # Dokumentasi instalasi dan penggunaan
```

---

## Panduan Menjalankan Aplikasi

### 1. Menjalankan dengan Docker (Disarankan)

Pastikan Docker telah terpasang di sistem Anda.

#### Menggunakan Docker Compose:
```bash
docker compose up -d --build
```
Aplikasi akan tersedia di: **`http://localhost:3021`**

Untuk menghentikan:
```bash
docker compose down
```

#### Menggunakan Docker CLI Standalone:
```bash
docker build -t csv2xlsx:latest .
docker run -d --restart always -p 3021:8000 --name csv2xlsx csv2xlsx:latest
```

---

### 2. Menjalankan Langsung di Lingkungan Python Lokal

#### Prasyarat:
- Python 3.10 atau versi yang lebih baru

#### Langkah-langkah:
```bash
# 1. Buat dan aktifkan virtual environment
python3 -m venv .venv
source .venv/bin/activate  # Linux/macOS
# atau: .venv\Scripts\activate  # Windows

# 2. Pasang dependensi
pip install --upgrade pip
pip install -r requirements.txt

# 3. Jalankan server aplikasi
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

Buka browser dan akses **`http://localhost:8000`**.

---

## Menjalankan Pengujian Otomatis (Automated Tests)

Jalankan suite pengujian unit dan integrasi dengan `pytest`:

```bash
# Pastikan virtual environment aktif
pytest tests/ -v
```

---

## Spesifikasi API REST

| Method | Endpoint | Deskripsi |
|---|---|---|
| `GET` | `/api/health` | Pemeriksaan kesehatan aplikasi (*health check*) |
| `POST` | `/api/preview` | Mengunggah file CSV dan mendapatkan metadata serta sampel pratinjau 20 baris |
| `POST` | `/api/convert` | Mengonversi file CSV ke spreadsheet `.xlsx` dengan styling profesional |
| `POST` | `/api/xlsx/preview` | Mengunggah file Excel `.xlsx`, membaca daftar sheet, dan pratinjau 20 baris |
| `POST` | `/api/xlsx/convert` | Mengonversi sheet Excel `.xlsx` menjadi file terdelimitasi `.csv` |

### Contoh Pemanggilan via cURL

#### 1. Konversi CSV ke Excel (.xlsx):
```bash
curl -X POST http://localhost:8000/api/convert \
  -F "file=@transaksi.csv" \
  -F "sheet_name=Laporan_Penjualan" \
  -F "apply_autofilter=true" \
  -F "apply_autowidth=true" \
  -o hasil_konversi.xlsx
```

#### 2. Pratinjau File Excel (.xlsx):
```bash
curl -X POST http://localhost:8000/api/xlsx/preview \
  -F "file=@laporan.xlsx"
```

#### 3. Konversi Excel (.xlsx) ke CSV:
```bash
curl -X POST http://localhost:8000/api/xlsx/convert \
  -F "file=@laporan.xlsx" \
  -F "sheet_name=Sheet1" \
  -F "delimiter=," \
  -F "encoding=utf-8-sig" \
  -o hasil_konversi.csv
```

---

## Lisensi
Didistribusikan di bawah lisensi MIT. Silakan gunakan dan kembangkan sesuai kebutuhan Anda.
