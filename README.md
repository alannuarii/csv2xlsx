# csv2xlsx (CSV to XLSX Converter)

Aplikasi web monolit mandiri (*self-hosted*) yang cepat, aman, dan elegan untuk mengonversi data terdelimitasi (`.csv`, `.tsv`, `.txt`) menjadi file spreadsheet Microsoft Excel (`.xlsx`) siap pakai dengan deteksi otomatis format, penataan gaya profesional, dan tabel pratinjau interaktif.

![csv2xlsx banner](app/static/favicon.svg)

---

## Fitur Utama

- **Konversi Tanpa Konfigurasi (*Zero Configuration*)**: Deteksi otomatis *encoding* (`UTF-8`, `UTF-8-SIG`, `Latin-1`, `Windows-1252`) dan karakter pemisah/delimiter (koma `,`, titik koma `;`, tab `\t`, pipa `|`).
- **Format Excel Profesional**: Baris header otomatis tebal (*bold*), warna latar kontras, border sel tipis, dan tombol filter (*auto-filter*) aktif.
- **Penataan Lebar Kolom Proporsional (*Auto-Fit Column Width*)**: Menghitung panjang karakter tiap kolom secara dinamis sehingga data tidak terpotong.
- **Smart Cell Typing**: Mengonversi teks angka dan boolean ke tipe data Excel asli, sekaligus menjaga angka dengan awalan nol (misal nomor HP `0812...` atau kode identitas `00123`) agar tetap terbaca sebagai teks.
- **Pratinjau Data Interaktif (*Data Preview*)**: Menampilkan kartu statistik (total baris, total kolom, delimiter, encoding) dan pratinjau tabel 20 baris pertama sebelum file diunduh.
- **100% Privasi Server Sendiri (*Self-Hosted*)**: Seluruh pemrosesan dilakukan di server lokal/internal Anda tanpa mengirimkan data ke pihak ketiga. File sementara otomatis dibersihkan.
- **Ringan & Tanpa Build Step (*Zero-Build Monolith*)**: Antarmuka SPA disajikan langsung oleh FastAPI menggunakan Tailwind CSS CDN dan Lucide Icons tanpa dependensi Node.js/npm.

---

## Arsitektur & Struktur Direktori

```text
csv2xlsx/
├── app/
│   ├── __init__.py
│   ├── main.py              # Routing FastAPI, validasi request, static file mount
│   ├── converter.py         # Delimiter & encoding detector, openpyxl writer & styling
│   └── static/
│       ├── favicon.svg      # Favicon format spreadsheet
│       └── index.html       # Antarmuka SPA lengkap (Tailwind, Lucide, Vanilla JS)
├── tests/
│   ├── __init__.py
│   ├── test_converter.py    # Unit tests untuk logic converter
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
| `POST` | `/api/preview` | Mengunggah file CSV dan mendapatkan metadata serta sampel 20 baris |
| `POST` | `/api/convert` | Mengonversi file CSV dan mengunduh aliran biner file `.xlsx` |

### Contoh Pemanggilan via cURL

#### Health Check:
```bash
curl -X GET http://localhost:8000/api/health
```

#### Konversi File CSV ke XLSX:
```bash
curl -X POST http://localhost:8000/api/convert \
  -F "file=@transaksi.csv" \
  -F "sheet_name=Laporan_Penjualan" \
  -F "apply_autofilter=true" \
  -F "apply_autowidth=true" \
  -o hasil_konversi.xlsx
```

---

## Lisensi
Didistribusikan di bawah lisensi MIT. Silakan gunakan dan kembangkan sesuai kebutuhan Anda.
