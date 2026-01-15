# Monitor Utilisasi Perangkat Jaringan

![Python Version](https://img.shields.io/badge/python-3.7+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Status](https://img.shields.io/badge/status-stable-brightgreen.svg)

Sebuah tool monitoring dan pelaporan jaringan komprehensif yang dirancang untuk manajemen infrastruktur jaringan enterprise. Tool ini mengotomatisasi pengumpulan data utilisasi, metrik performa sistem, dan inventori hardware dari peralatan jaringan, menghasilkan laporan Excel profesional dengan analitik yang detail.

## 🚀 Fitur

### Kemampuan Inti
- **Multi-Device Monitoring**: Monitoring bersamaan untuk beberapa perangkat jaringan
- **Analisis Utilisasi FPC**: Analisis detail utilisasi port dan traffic
- **Inventori Hardware**: Pelacakan komponen hardware yang komprehensif
- **Monitoring Performa Sistem**: Metrik CPU, memori, storage, dan temperatur
- **Manajemen Alarm**: Pengumpulan alarm real-time dan monitoring status
- **Pelaporan Profesional**: Laporan Excel otomatis dengan beberapa worksheet

### Fitur Lanjutan
- **Deteksi Modul Intelligent**: Identifikasi otomatis modul FPC dari hardware chassis
- **Inferensi Tipe SFP**: Deteksi pintar tipe optical transceiver
- **Analisis Pola Traffic**: Analitik utilisasi traffic yang canggih
- **Timestamping Zone-aware**: Deteksi otomatis zona waktu Indonesia (WIB/WITA/WIT)
- **Sequential Processing**: Pemrosesan satu-perangkat-per-waktu yang reliable untuk tingkat keberhasilan maksimum
- **Debug Logging**: Logging komprehensif dengan output debug yang terorganisir

### Worksheet Laporan
1. **Utilisasi FPC** - Data utilisasi FPC utama
2. **Utilisasi Port** - Utilisasi detail level port
3. **Status Alarm** - Kondisi alarm saat ini
4. **Inventori Hardware** - Daftar lengkap komponen hardware  
5. **Performa Sistem** - Metrik kesehatan sistem
6. **Ringkasan Dashboard** - Ringkasan eksekutif dan metrik kunci

## 📋 Persyaratan

### Kebutuhan Sistem
- **Sistem Operasi**: Windows 10/11, Linux, macOS
- **Versi Python**: 3.7 atau lebih tinggi
- **Memori**: Minimum 4GB RAM yang direkomendasikan
- **Storage**: 500MB ruang kosong untuk laporan dan log

### Kebutuhan Jaringan
- Konektivitas jaringan ke perangkat target
- Akses TACACS+ server untuk autentikasi
- Akses SSH ke peralatan jaringan (biasanya port 22)

### Dependensi Python
```bash
pip install paramiko openpyxl
```

## 🛠️ Instalasi

### 1. Clone Repository
```bash
git clone https://github.com/wokding/monitoing.git
cd monitoing
```

### 2. Install Dependensi Python
```bash
pip install paramiko openpyxl
```

### 3. Konfigurasi Kredensial
**PENTING**: Jangan pernah commit file kredensial aktual ke Git!

```bash
# Copy template kredensial
cp _telkom_access.xml.example _telkom_access.xml

# Edit dengan kredensial Anda
# Windows: notepad _telkom_access.xml
# Linux/Mac: nano _telkom_access.xml
```

### 4. Konfigurasi Daftar Perangkat
Edit file `list_cnop.txt` dengan daftar perangkat yang ingin dimonitor.

### 5. Verifikasi Setup
```bash
python fpc_utilisasi.py
```

## ⚙️ Konfigurasi

### 1. File Konfigurasi Akses (PENTING!)
**SECURITY NOTICE**: File `_telkom_access.xml` berisi kredensial sensitif dan **TIDAK BOLEH** di-commit ke Git!

#### Setup Kredensial:
```bash
# 1. Copy template kredensial
cp _telkom_access.xml.example _telkom_access.xml

# 2. Edit dengan kredensial Anda
notepad _telkom_access.xml  # Windows
# atau
nano _telkom_access.xml     # Linux/Mac

# 3. Pastikan file permissions aman (Linux/Mac)
chmod 600 _telkom_access.xml
```

#### Format Konfigurasi:
```xml
<telkom-access>
    <tacacs-list>
        <tacacs-server>10.x.x.x</tacacs-server>
        <tacacs-server>10.x.x.x</tacacs-server>
    </tacacs-list>
    <tacacs-access>
        <tacacs-user>your-tacacs-username</tacacs-user>
        <tacacs-pass>your-tacacs-password</tacacs-pass>
    </tacacs-access>
    <router-access>
        <router-user>your-router-username</router-user>
        <router-pass>your-router-password</router-pass>
    </router-access>
</telkom-access>
```

**⚠️ CATATAN KEAMANAN**:
- File `_telkom_access.xml` sudah ada di `.gitignore`
- Jangan pernah commit kredensial aktual
- Gunakan `.xml.example` sebagai template
- Pastikan file permissions yang tepat

### 2. Konfigurasi Daftar Perangkat
File `list_cnop.txt` sudah tersedia dengan daftar perangkat default. Edit sesuai kebutuhan:

```
R3.KYA.PE-MOBILE.1
R3.KYA.PE-MOBILE.2
R4.NSK.PE-MOBILE.1
R4.NSK.PE-MOBILE.2
R5.GYG.ASBR-TSEL.1
R5.GYG.RR-TSEL.1
R5.KBL.ASBR-TSEL.1
R5.KBL.RR-TSEL.1
```

**Catatan**: Tambahkan atau hapus perangkat sesuai dengan yang ingin Anda monitor.

### 3. Pengaturan Autentikasi

| Parameter | Deskripsi | Contoh |
|-----------|-----------|--------|
| `tacacs-server` | Alamat IP server TACACS+ | 10.62.170.56 |
| `tacacs-user` | Username untuk autentikasi TACACS+ | network-admin |
| `tacacs-pass` | Password untuk autentikasi TACACS+ | SecurePass123 |
| `router-user` | Username level perangkat | device-admin |
| `router-pass` | Password level perangkat | DevicePass456 |

## 🚀 Penggunaan

### Cara Menjalankan Script

#### Metode 1: Menggunakan Command Prompt
1. Buka Command Prompt (cmd) atau PowerShell
2. Navigasi ke folder script:
   ```powershell
   cd "C:\Users\User\OneDrive\Desktop\TELKOM_Script-v25"
   ```
3. Jalankan script:
   ```powershell
   python fpc_utilisasi.py
   ```

#### Metode 2: Langsung dari File Explorer
1. Buka folder `C:\Users\User\OneDrive\Desktop\TELKOM_Script-v25`
2. Klik kanan pada area kosong, pilih "Open PowerShell window here"
3. Ketik: `python fpc_utilisasi.py`

#### Metode 3: Double-click (jika Python sudah di-associate)
1. Double-click file `fpc_utilisasi.py` langsung dari File Explorer
2. Script akan berjalan di command prompt yang terbuka otomatis

### Eksekusi Dasar
```powershell
cd "C:\Users\User\OneDrive\Desktop\TELKOM_Script-v25"
python fpc_utilisasi.py
```

### Opsi Command Line
Script mendukung berbagai mode eksekusi dan parameter (konfigurasi dalam script):

- **Sequential Processing**: Mode default untuk reliabilitas maksimum
- **Debug Mode**: Logging yang ditingkatkan untuk troubleshooting
- **Custom Timeout**: Timeout koneksi yang dapat disesuaikan
- **Retry Logic**: Percobaan ulang yang dapat dikonfigurasi untuk koneksi yang gagal

### Alur Eksekusi
1. **Inisialisasi**: Muat konfigurasi dan validasi dependensi
2. **Autentikasi**: Koneksi ke server TACACS+
3. **Device Discovery**: Baca daftar perangkat dari konfigurasi
4. **Pengumpulan Data**: Monitoring perangkat secara berurutan
5. **Pemrosesan Data**: Parse dan analisis data yang dikumpulkan
6. **Pembuatan Laporan**: Buat laporan Excel dengan beberapa worksheet
7. **Cleanup**: Organisir file debug dan log

## 📊 File Output

### Lokasi Script dan File
**Lokasi Script**:
```
C:\Users\User\OneDrive\Desktop\TELKOM_Script-v25\
├── fpc_utilisasi.py          # Script utama yang dijalankan
├── _telkom_access.xml        # Konfigurasi akses dan kredensial
├── list_cnop.txt            # Daftar perangkat target
└── README.md                # Dokumentasi
```

### Lokasi Output
Setelah script dijalankan, file hasil akan dibuat di Desktop:
```
C:\Users\User\Desktop\
├── FPC Utilisasi Reports\     # Folder utama laporan
│   ├── Daily Reports\
│   │   └── FPC_Utilization_2025-09-16_143022.xlsx
│   ├── Monthly Reports\
│   │   └── FPC_Monthly_2025-09.xlsx
│   └── All Debug\
│       ├── Logs\
│       │   ├── connection_errors.log
│       │   ├── chassis_missing_modules.log
│       │   └── processing_details.log
│       └── XML Outputs\
│           ├── R3.KYA.PE-MOBILE.1_chassis_hardware.xml
│           └── R3.KYA.PE-MOBILE.1_interface_statistics.xml
```

**Catatan**: Script akan otomatis membuat folder-folder ini jika belum ada.

### Isi Laporan

#### Sheet Utilisasi Utama
- Nama Node dan Divisi Regional
- Nama Interface dan Deskripsi  
- Tipe Modul dan Kapasitas Port
- Traffic Saat Ini (GB) dan Utilisasi (%)
- Alert Traffic dan Indikator Status

#### Sheet Utilisasi Port
- Statistik detail level port
- Analisis Last Flapped
- Deteksi keberadaan SFP
- Status konfigurasi
- Alert flapping

#### Sheet Status Alarm
- Kondisi alarm real-time
- Klasifikasi dan tingkat keparahan alarm
- Timestamp dan deskripsi
- Pelacakan status

#### Sheet Inventori Hardware
- Daftar lengkap komponen
- Nomor part dan serial number
- Deskripsi model dan versi
- Status dan kesehatan komponen

#### Sheet Performa Sistem
- Utilisasi CPU dan load average
- Penggunaan dan ketersediaan memori
- Kapasitas dan utilisasi storage
- Monitoring temperatur
- Informasi versi software

#### Sheet Ringkasan Dashboard
- Metrik ringkasan eksekutif
- Key performance indicator
- Analisis tren
- Ringkasan alert

## 🔧 Konfigurasi Lanjutan

### Pengaturan Timeout
```python
BANNER_TIMEOUT = 180        # Timeout koneksi awal (detik)
INITIAL_TEST_RETRIES = 5    # Percobaan ulang koneksi
PER_NODE_RETRIES = 5        # Percobaan ulang per perangkat
MAX_WORKERS = 1             # Pemrosesan berurutan (direkomendasikan)
```

### Performance Tuning
- **Sequential Processing**: Direkomendasikan untuk reliabilitas maksimum
- **Connection Pooling**: Penggunaan ulang koneksi jika memungkinkan
- **Memory Management**: Pembersihan otomatis struktur data besar
- **Progress Tracking**: Update status pemrosesan real-time

### Konfigurasi Debug
```python
# Aktifkan comprehensive debug logging
DEBUG_MODE = True
DEBUG_FOLDER = "All Debug"
LOG_LEVEL = "DEBUG"
```

## 🐛 Troubleshooting

### Masalah Umum

#### Kegagalan Koneksi
```
ERROR: Failed to connect to device X.X.X.X
```
**Solusi**:
- Verifikasi konektivitas jaringan ke server TACACS+
- Periksa kredensial username/password
- Konfirmasi aksesibilitas perangkat
- Tinjau aturan firewall

#### Error Autentikasi
```
ERROR: TACACS+ authentication failed
```
**Solusi**:
- Verifikasi konfigurasi server TACACS+
- Periksa username/password di konfigurasi XML
- Konfirmasi izin akun
- Test koneksi SSH manual

#### Dependensi Hilang
```
ERROR: Missing dependency: paramiko
```
**Solusi**:
```bash
pip install paramiko openpyxl
pip install --upgrade paramiko openpyxl
```

#### Error Pembuatan Excel
```
ERROR: Failed to create Excel report
```
**Solusi**:
- Pastikan ruang disk yang cukup
- Periksa izin write ke Desktop
- Tutup file Excel yang terbuka dengan nama yang sama
- Verifikasi instalasi openpyxl

### Debug Logging
Aktifkan logging detail dengan memeriksa file debug:
- `connection_errors.log` - Masalah koneksi dan autentikasi
- `chassis_missing_modules.log` - Masalah deteksi hardware
- `processing_details.log` - Informasi pemrosesan data

### Analisis Log
Untuk Windows PowerShell:
```powershell
# Lihat error koneksi terbaru
Get-Content "C:\Users\User\Desktop\FPC Utilisasi Reports\All Debug\Logs\connection_errors.log" -Tail 20

# Periksa status pemrosesan
Select-String "SUCCESS|ERROR" "C:\Users\User\Desktop\FPC Utilisasi Reports\All Debug\Logs\processing_details.log"
```

Atau buka file log langsung dengan Notepad:
```powershell
notepad "C:\Users\User\Desktop\FPC Utilisasi Reports\All Debug\Logs\connection_errors.log"
```

## 📈 Optimasi Performa

### Best Practices
1. **Sequential Processing**: Gunakan `MAX_WORKERS = 1` untuk reliabilitas maksimum
2. **Timeout Tuning**: Sesuaikan timeout berdasarkan kondisi jaringan
3. **Retry Logic**: Konfigurasi jumlah percobaan ulang yang tepat
4. **Memory Management**: Monitor penggunaan memori untuk daftar perangkat yang besar
5. **Debug Cleanup**: Bersihkan file debug secara teratur untuk mencegah masalah disk

### Pertimbangan Scaling
- **Jumlah Perangkat**: Diuji dengan hingga 100+ perangkat
- **Waktu Pemrosesan**: Sekitar 2-3 menit per perangkat
- **Penggunaan Memori**: ~50MB per perangkat selama pemrosesan
- **Kebutuhan Storage**: ~10MB per perangkat untuk file debug

## 📝 Changelog

### Versi 25.0
- Deteksi modul yang ditingkatkan dari XML hardware chassis
- Algoritma inferensi tipe SFP yang diperbaiki
- Penambahan monitoring performa sistem yang komprehensif
- Implementasi analisis pola traffic yang intelligent
- Debug logging dan error handling yang ditingkatkan
- Format laporan Excel yang profesional
- Dukungan timezone Indonesia (WIB/WITA/WIT)

## 🤝 Kontribusi

### Setup Development
1. Fork repository
2. Buat feature branch
3. Install dependensi development
4. Buat perubahan Anda
5. Test secara menyeluruh
6. Submit pull request

### Code Style
- Ikuti guideline PEP 8
- Gunakan nama variabel yang bermakna
- Tambahkan docstring yang komprehensif
- Sertakan error handling
- Tulis unit test untuk fitur baru

## 📄 Lisensi

Proyek ini dilisensikan di bawah MIT License - lihat file [LICENSE](LICENSE) untuk detail.

## 🆘 Dukungan

### Mendapatkan Bantuan
1. Periksa bagian troubleshooting di atas
2. Tinjau log debug untuk error spesifik
3. Verifikasi pengaturan konfigurasi
4. Test konektivitas jaringan secara manual

### Melaporkan Masalah
Ketika melaporkan masalah, silakan sertakan:
- Versi Python dan sistem operasi
- Pesan error lengkap
- File konfigurasi (dengan data sensitif dihilangkan)
- Excerpts log debug
- Langkah-langkah untuk mereproduksi masalah

## 🔒 Keamanan & Best Practices

### Manajemen Kredensial
**CRITICAL**: Jangan pernah commit kredensial ke version control!

#### File yang Harus Dilindungi:
- ✅ `_telkom_access.xml` - **SUDAH DI .gitignore**
- ✅ `*.log` files - **SUDAH DI .gitignore**
- ✅ `*.xlsx` reports - **SUDAH DI .gitignore**
- ✅ `.env` files - **SUDAH DI .gitignore**

#### Checklist Keamanan:
```bash
# 1. Pastikan .gitignore aktif
cat .gitignore

# 2. Verifikasi file yang akan di-commit
git status

# 3. Pastikan tidak ada kredensial
git diff

# 4. Check file sensitif tidak ter-track
git ls-files | grep -E "_telkom_access.xml|\.env$|\.log$"
# Output harus kosong!
```

### Best Practices
1. **Kredensial**:
   - Gunakan password yang kuat dan unik
   - Rotasi kredensial secara berkala
   - Batasi akses file dengan permissions (chmod 600)
   - Jangan share kredensial via email/chat

2. **Version Control**:
   - Selalu gunakan template files (.example)
   - Review `git status` sebelum commit
   - Gunakan `git diff` untuk verifikasi
   - Jangan force push ke main branch

3. **Network Security**:
   - Pastikan koneksi SSH yang aman
   - Monitor log akses secara reguler
   - Gunakan least privilege principle
   - Audit keamanan berkala

---

## 📁 Struktur Project

```
monitoing/
├── .gitignore                      # File dan folder yang diabaikan Git
├── README.md                       # Dokumentasi project
├── fpc_utilisasi.py               # Script utama monitoring
├── list_cnop.txt                  # Daftar perangkat target
├── _telkom_access.xml.example     # Template kredensial (COMMIT INI)
├── _telkom_access.xml             # Kredensial aktual (JANGAN COMMIT!)
├── .env.example                   # Template environment variables
└── .env                           # Config lokal (JANGAN COMMIT!)

# Output (auto-generated, tidak di-commit):
reports/                           # Laporan Excel
logs/                              # File log
debug/                             # Debug output
```

**Catatan**: File yang ditandai "JANGAN COMMIT" sudah otomatis diabaikan oleh `.gitignore`

---

**Catatan**: Tool ini dirancang untuk monitoring jaringan enterprise dan harus digunakan sesuai dengan kebijakan dan prosedur keamanan organisasi Anda. Selalu test di lingkungan non-production sebelum deploy ke infrastruktur kritis.