# Screenshot Placeholder

Untuk menambahkan screenshot contoh output:

## 📸 Cara Membuat Screenshot yang Aman

### Step 1: Buka Excel Report
```powershell
# Buka file Excel hasil monitoring
start FPC-Occupancy_Report_*.xlsx
```

### Step 2: Screenshot Sheet
1. Pilih sheet yang ingin di-screenshot (Dashboard, FPC Utilization, dll)
2. Tekan `Win + Shift + S` untuk Snipping Tool
3. Select area yang ingin di-capture

### Step 3: Mask Data Sensitif ⚠️ PENTING!

**WAJIB di-blur/redact:**
- ❌ IP Loopback (127.x.x.x, 10.x.x.x, 172.x.x.x, 192.168.x.x)
- ❌ Management IP addresses
- ❌ Hostname yang spesifik (misal: R3.JKT.PE-CORE.1)
- ❌ Serial numbers perangkat
- ❌ Lokasi geografis detail
- ❌ Interface names yang sensitive

**Tools untuk blur:**
- Paint (Windows built-in)
- Paint 3D (Windows 10/11)
- Photopea.com (online, gratis)
- GIMP (software gratis)

### Step 4: Save File
```
Nama file:
- dashboard_summary.png
- fpc_utilization.png
- port_utilization.png
- hardware_inventory.png
- alarm_status.png
- system_performance.png
```

Save di folder ini: `examples/screenshots/`

### Step 5: Verifikasi Keamanan

Sebelum commit, cek:
- [ ] Tidak ada IP address production visible
- [ ] Hostname di-mask (contoh: NODE-01, ROUTER-XX)
- [ ] Serial numbers hidden
- [ ] Lokasi tidak specific
- [ ] File size < 1MB

### Step 6: Commit Screenshot
```bash
git add examples/screenshots/*.png
git commit -m "docs: add sample output screenshots (data masked)"
git push
```

---

## 📋 Template Data Masking

### IP Address Masking:
```
Production:  10.62.170.56    →  192.0.2.1 (TEST-NET)
Production:  10.60.190.15    →  192.0.2.2
Management:  172.16.10.5     →  192.0.2.10
Loopback:    127.0.0.1       →  192.0.2.254
```

### Hostname Masking:
```
Production:  R3.KYA.PE-MOBILE.1  →  R3.XXX.PE-SAMPLE.1
Production:  R5.GYG.ASBR-TSEL.1  →  R5.YYY.ASBR-SAMPLE.1
Generic:     Use NODE-01, ROUTER-01, SWITCH-01
```

### Serial Number Masking:
```
Production:  ABC123456789XYZ  →  XXXXXXXXXX
Production:  JNP1234567890    →  JNPXXXXXXXXX
```

---

## 🎨 Example dengan Paint

### Cara blur di Paint:
1. Open screenshot di Paint
2. Pilih **Select** tool → Rectangle
3. Select area dengan IP/hostname
4. Pilih warna solid (hitam/putih)
5. Fill area dengan `Ctrl+Shift+F` atau Paint bucket
6. Atau tambah text box dengan "x.x.x.x" atau "192.0.2.x"
7. Save As PNG

### Cara blur di Paint 3D:
1. Open screenshot
2. Menu → **Magic select**
3. Select sensitive area
4. Apply **Blur** effect
5. Or use **Stickers** → Add blur stickers
6. Export as PNG

---

## 💡 Tips

1. **Use consistent masking**: Pakai format yang sama untuk semua IP
2. **Keep layout intact**: Hanya ganti data, bukan struktur
3. **Preserve readability**: Orang harus bisa lihat format report
4. **Test before commit**: Review screenshot sebelum upload
5. **Ask if unsure**: Kalau ragu, jangan upload dulu

---

**Status**: Waiting for screenshots to be added

Setelah add screenshots, hapus file ini dan uncomment baris di `examples/README.md`
