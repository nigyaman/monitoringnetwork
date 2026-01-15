# Examples - Sample Output

## 📊 Output Samples

Folder ini berisi contoh output dari network monitoring tool untuk tujuan dokumentasi.

## ⚠️ IMPORTANT - Data Masking

**Semua data sensitif telah di-redact/di-mask:**
- ❌ IP Addresses (loopback, management, dll)
- ❌ Hostname yang spesifik
- ❌ Serial numbers
- ❌ Lokasi geografis detail
- ❌ Interface names yang sensitive

**Data yang ditampilkan:**
- ✅ Format dan struktur report
- ✅ Column headers dan layout
- ✅ Sample data dengan nilai dummy/masked
- ✅ Tampilan visual dashboard

## 📁 Struktur Folder

```
examples/
├── README.md                          # File ini
├── screenshots/                       # Screenshot dengan data di-mask
│   ├── dashboard_summary.png
│   ├── fpc_utilization.png
│   ├── port_utilization.png
│   ├── hardware_inventory.png
│   └── alarm_status.png
└── sample_output_template.xlsx        # Template Excel (COMING SOON)
```

## 🖼️ Screenshots

### 1. Dashboard Summary
Ringkasan eksekutif dengan key metrics dan status overview.

![Dashboard Summary](screenshots/dashboard_summary.png)

### 2. FPC Utilization
Detail utilisasi FPC per node dengan traffic analysis.

![FPC Utilization](screenshots/fpc_utilization.png)

### 3. Port Utilization
Analisis detail level port dengan status dan performance metrics.

![Port Utilization](screenshots/port_utilization.png)

### 4. Hardware Inventory
Inventori komponen hardware dengan part numbers dan status.

![Hardware Inventory](screenshots/hardware_inventory.png)

### 5. Alarm Status
Real-time alarm monitoring dan status tracking.

![Alarm Status](screenshots/alarm_status.png)

## 📝 Catatan

- Screenshots diambil dari environment testing dengan data dummy
- Tidak ada data production yang ditampilkan
- Semua IP addresses diganti dengan format `x.x.x.x` atau `192.0.2.x` (TEST-NET)
- Semua serial numbers diganti dengan `XXXXXXXXXXXX`
- Hostname menggunakan format generic `NODE-XX`

## 🔒 Security Notice

**JANGAN upload screenshot production tanpa masking!**

Pastikan sebelum upload screenshot:
1. ✅ Semua IP addresses di-blur atau di-redact
2. ✅ Serial numbers di-hide
3. ✅ Hostname sensitive di-ganti
4. ✅ Lokasi geografis tidak terlihat
5. ✅ Management interface tidak visible

## 🛠️ Cara Membuat Screenshot yang Aman

### Menggunakan Paint / Photo Editor:

1. Buka file Excel report
2. Screenshot sheet yang ingin ditampilkan
3. Buka di Paint atau photo editor
4. Gunakan rectangle tool untuk blur/block:
   - IP addresses
   - Serial numbers
   - Hostname specific
   - Lokasi sensitive

### Menggunakan Windows Snipping Tool:

1. Buka Excel report
2. Gunakan Snipping Tool (Win + Shift + S)
3. Capture area yang ingin ditampilkan
4. Edit dan blur data sensitive
5. Save dengan nama deskriptif

### Tools Recommended:

- **Windows**: Paint, Snipping Tool, Paint 3D
- **Online**: Photopea (free Photoshop alternative)
- **Software**: GIMP (free), Photoshop

## 📋 Checklist Sebelum Upload

Sebelum upload screenshot ke repository:

- [ ] Tidak ada IP address production visible
- [ ] Tidak ada serial numbers visible
- [ ] Hostname sudah di-mask/generalize
- [ ] Tidak ada informasi lokasi specific
- [ ] File size reasonable (<1MB per image)
- [ ] Format PNG atau JPG
- [ ] Nama file descriptive dan clear

## 🎯 Alternative: Sample Data Excel

Sebagai alternatif screenshot, Anda bisa:

1. Copy Excel output
2. Replace semua data sensitive dengan dummy data
3. Save sebagai `sample_output_template.xlsx`
4. Upload template tersebut

Contoh replacement:
- IP: `10.1.1.1` → `192.0.2.1` (TEST-NET range)
- Hostname: `R3.JKT.PE-CORE.1` → `R3.XXX.PE-SAMPLE.1`
- Serial: `ABC123456789` → `XXXXXXXXXXXX`

---

**Remember**: Keamanan data production adalah prioritas!
