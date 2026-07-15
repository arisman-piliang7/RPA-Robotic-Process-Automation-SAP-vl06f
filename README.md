# RPA – SAP VL06F Automatic Report Downloader

Aplikasi RPA untuk mengunduh laporan **VL06F** (Outbound Deliveries – Goods Issues)
dari SAP secara otomatis, dengan dukungan penjadwalan (*scheduling*).

Dibuat untuk PT Pertamina Patra Niaga.

---

## Fitur

| Fitur | Keterangan |
|---|---|
| Otomatis login SAP | Menggunakan SAP GUI Scripting API (COM) |
| Jalankan transaksi VL06F | Isi parameter pemilihan, eksekusi laporan |
| Ekspor laporan | Simpan ke file `.xlsx` atau `.txt` |
| Penjadwalan | Cron schedule via APScheduler (misal: setiap hari kerja pukul 07:00) |
| Konfigurasi fleksibel | Semua parameter ada di `config.ini` |
| Logging harian | File log dirotasi otomatis setiap hari |

---

## Prasyarat

| Prasyarat | Versi minimum |
|---|---|
| Windows | 10 / Server 2016 |
| SAP GUI for Windows | 7.50+ |
| Python | 3.11+ |
| SAP GUI Scripting | Diaktifkan di SAP GUI Options |

### Aktifkan SAP GUI Scripting

1. Buka SAP GUI → **Options** (Alt+F7).
2. Pilih **Accessibility & Scripting → Scripting**.
3. Centang **Enable scripting**.
4. Klik OK dan restart SAP GUI.

---

## Instalasi

```bash
# 1. Clone repositori
git clone https://github.com/arisman-piliang7/RPA-Robotic-Process-Automation-SAP-vl06f.git
cd RPA-Robotic-Process-Automation-SAP-vl06f

# 2. Buat virtual environment (opsional tapi disarankan)
python -m venv .venv
.venv\Scripts\activate   # Windows

# 3. Install dependensi
pip install -r requirements.txt

# 4. Salin contoh konfigurasi dan isi dengan nilai yang sebenarnya
copy config.ini.example config.ini
```

> **Penting:** Jangan pernah commit `config.ini` ke repository karena berisi password SAP.
> File ini sudah terdaftar di `.gitignore`.

---

## Konfigurasi (`config.ini`)

Buka `config.ini` dan sesuaikan setiap bagian:

### `[SAP]`

```ini
system_description = SAP ECC Production   ; nama sistem di SAP Logon
client             = 100                   ; nomor client/mandant
username           = YOUR_USERNAME
password           = YOUR_PASSWORD
language           = EN
screen_wait_timeout = 30
```

### `[VL06F]`

```ini
shipping_point   = SP01          ; kode shipping point (pisahkan koma jika lebih dari satu)
date_range_mode  = today         ; today | yesterday | last_n_days | custom
last_n_days      = 7             ; digunakan jika date_range_mode = last_n_days
date_from        = 01.01.2024    ; digunakan jika date_range_mode = custom
date_to          = 31.01.2024
```

### `[OUTPUT]`

```ini
output_dir  = output/{YYYY}/{MM}
file_name   = VL06F_Report_{YYYY}{MM}{DD}_{HH}{MIN}{SS}.xlsx
file_format = xlsx               ; xlsx | txt
```

Placeholder yang didukung: `{YYYY}`, `{MM}`, `{DD}`, `{HH}`, `{MIN}`, `{SS}`.

### `[SCHEDULE]`

```ini
enabled          = false         ; ubah ke true untuk mode penjadwalan
cron_day_of_week = mon-fri
cron_hour        = 7
cron_minute      = 0
timezone         = Asia/Jakarta
```

---

## Penggunaan

### Jalankan sekali (manual)

```bash
python main.py
```

### Jalankan dengan jadwal otomatis

```bash
python main.py --schedule
```

Proses akan berjalan terus-menerus dan mengeksekusi unduhan sesuai jadwal yang
dikonfigurasi di `[SCHEDULE]`.  Tekan **Ctrl+C** untuk menghentikan.

### Pilih file konfigurasi lain

```bash
python main.py --config C:\rpa\produksi\config.ini
```

### Bantuan

```bash
python main.py --help
```

---

## Menjalankan sebagai Windows Service / Task Scheduler

### Windows Task Scheduler

1. Buka **Task Scheduler** → **Create Basic Task**.
2. Atur trigger sesuai kebutuhan (misal: daily jam 07:00).
3. Action: **Start a program**
   - Program: `C:\rpa\.venv\Scripts\python.exe`
   - Arguments: `C:\rpa\main.py --config C:\rpa\config.ini`
   - Start in: `C:\rpa`

> Jika menggunakan `--schedule`, cukup jalankan sekali tanpa Task Scheduler.

---

## Struktur Proyek

```
├── main.py                  # Entry point utama
├── scheduler.py             # Penjadwalan dengan APScheduler
├── config.ini.example       # Template konfigurasi (aman untuk commit)
├── requirements.txt         # Dependensi Python
├── sap/
│   ├── __init__.py
│   ├── connector.py         # Manajemen koneksi SAP GUI
│   └── vl06f.py             # Otomasi transaksi VL06F
├── utils/
│   ├── __init__.py
│   ├── logger.py            # Setup logging
│   └── file_utils.py        # Helper path & folder output
├── output/                  # Hasil unduhan (dibuat otomatis, di-gitignore)
└── logs/                    # File log (dibuat otomatis, di-gitignore)
```

---

## Log

Log tersimpan di folder `logs/` dengan nama `rpa_vl06f.log`.
File dirotasi setiap tengah malam, backup 7 hari terakhir disimpan.

---

## Lisensi

Internal use only – PT Pertamina Patra Niaga.
