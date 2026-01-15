# 🎯 PANDUAN CEPAT: Upload Screenshot Hasil Script ke GitHub

## ⚠️ LANGKAH PENTING - MASK DATA SENSITIF!

### Step 1️⃣: Buka File Excel
```powershell
# Buka file Excel hasil monitoring dari folder project Anda
cd <path-to-your-project>
start FPC-Occupancy_Report_*.xlsx
```

### Step 2️⃣: Screenshot Sheet yang Ingin Ditampilkan

**Sheet yang bagus untuk ditampilkan:**
1. Dashboard Summary (overview)
2. Utilisasi FPC (main data)
3. Utilisasi Port (detail)
4. Hardware Inventory
5. Alarm Status

**Cara Screenshot:**
1. Tekan `Win + Shift + S` (Snipping Tool)
2. Pilih area sheet yang ingin di-capture
3. Screenshot otomatis ke clipboard

### Step 3️⃣: BLUR/REDACT Data Sensitif ⚠️ WAJIB!

**Buka Paint dan Edit Screenshot:**

```powershell
# Cara cepat: paste langsung ke Paint
# 1. Buka Paint
mspaint

# 2. Paste screenshot (Ctrl+V)
# 3. Gunakan Rectangle tool untuk tutup data sensitif
```

**WAJIB DI-TUTUP:**
- ❌ **IP Loopback** (127.x.x.x, 10.x.x.x, 192.168.x.x)
- ❌ **IP Management** 
- ❌ **Hostname Spesifik** (R3.KYA.PE-MOBILE.1 → ganti jadi "R3.XXX.PE-SAMPLE.1")
- ❌ **Serial Numbers** perangkat
- ❌ **Lokasi Detail** (nama kota bisa diganti)

**CARA TUTUP di Paint:**
1. Pilih **Rectangle** tool
2. Pilih warna **Solid** (hitam atau putih)
3. Gambar rectangle di atas IP/hostname
4. Atau gunakan **Text tool** tulis "x.x.x.x" atau "192.0.2.x"

### Step 4️⃣: Save Screenshot

```powershell
# Save dengan nama yang jelas
# Lokasi: <your-project>/examples/screenshots/

# Nama file:
# - dashboard_summary.png
# - fpc_utilization.png
# - port_utilization.png
# - hardware_inventory.png
# - alarm_status.png
```

Save file PNG ke folder: `examples\screenshots\`

### Step 5️⃣: Verifikasi Keamanan ✅

**CHECKLIST SEBELUM UPLOAD:**
- [ ] Tidak ada IP production visible
- [ ] Hostname sudah di-mask
- [ ] Serial numbers hidden
- [ ] Lokasi tidak terlalu specific
- [ ] File size < 1MB per image

### Step 6️⃣: Commit & Push ke GitHub

```powershell
# Navigasi ke folder project
cd <your-project-path>

# Add screenshot
git add examples/screenshots/*.png

# Commit
git commit -m "docs: add sample output screenshots with masked data"

# Push ke GitHub
git push
```

---

## 🎨 CONTOH MASKING:

### IP Address Replacement:
```
PRODUCTION (JANGAN!):        SAFE (BOLEH):
10.62.170.56              →  192.0.2.1
10.60.190.15              →  192.0.2.2
127.0.0.1                 →  192.0.2.254
172.16.10.5               →  x.x.x.x
```

### Hostname Replacement:
```
PRODUCTION (JANGAN!):        SAFE (BOLEH):
R3.KYA.PE-MOBILE.1        →  R3.XXX.PE-SAMPLE.1
R5.GYG.ASBR-TSEL.1        →  R5.YYY.ASBR-SAMPLE.1
R6.BJB.PE-MOBILE.1        →  ROUTER-01
```

### Serial Number Replacement:
```
PRODUCTION (JANGAN!):        SAFE (BOLEH):
ABC123456789XYZ           →  XXXXXXXXXXXX
JNP1234567890             →  JNPXXXXXXXXX
```

---

## 📸 VISUAL GUIDE:

### ❌ JANGAN (Data Mentah):
```
+------------------+----------------+
| Hostname         | IP Address     |
+------------------+----------------+
| R3.KYA.PE-MOBILE | 10.62.170.56  |  ← SENSITIF!
| R5.GYG.ASBR-TSEL | 10.60.190.15  |  ← JANGAN!
+------------------+----------------+
```

### ✅ BOLEH (Data Masked):
```
+------------------+----------------+
| Hostname         | IP Address     |
+------------------+----------------+
| R3.XXX.PE-SAMPLE | 192.0.2.1     |  ← AMAN
| R5.YYY.ASBR-DEMO | 192.0.2.2     |  ← AMAN
+------------------+----------------+
```

---

## 🛠️ TOOLS ALTERNATIF:

### Paint (Built-in Windows):
- Pros: Sudah ada, simple
- Cons: Agak basic

### Paint 3D (Windows 10/11):
- Pros: Ada blur effect
- Cons: Agak lebih kompleks

### Photopea.com (Online):
- Pros: Gratis, powerful, seperti Photoshop
- Cons: Perlu internet

### GIMP (Free Software):
- Pros: Sangat powerful
- Cons: Perlu download/install

---

## ⚡ QUICK COMMANDS:

```powershell
# 1. Buka Excel
start FPC-Occupancy_Report_*.xlsx

# 2. Screenshot (Win+Shift+S), edit di Paint, save ke examples\screenshots\

# 3. Add & commit dari folder project Anda
cd <your-project-path>
git add examples/screenshots/*.png
git commit -m "docs: add sample screenshots (data masked)"
git push
```

---

## 🆘 TROUBLESHOOTING:

### Q: Screenshot terlalu besar (>1MB)?
**A:** Resize di Paint atau compress:
```powershell
# Di Paint: Image → Resize → 75% atau 50%
```

### Q: Lupa blur IP address?
**A:** JANGAN push! Edit ulang:
```powershell
git reset HEAD examples/screenshots/file.png
# Edit ulang screenshot, blur IP
git add examples/screenshots/file.png
git commit -m "docs: add sample screenshots (data masked)"
```

### Q: Sudah ter-push dengan data sensitif?
**A:** SEGERA remove:
```powershell
git rm examples/screenshots/file.png
git commit -m "remove: screenshot with sensitive data"
git push
```

---

## 📋 FINAL CHECKLIST:

✅ Screenshot dari Excel report
✅ IP addresses DI-MASK (x.x.x.x atau 192.0.2.x)
✅ Hostname DI-MASK (XXX atau SAMPLE)
✅ Serial numbers DI-HIDE
✅ File size < 1MB
✅ Format PNG atau JPG
✅ Nama file descriptive
✅ Saved di `examples\screenshots\`
✅ Git add, commit, push

---

**Remember**: Jangan terburu-buru! Review screenshot sebelum upload. 
Lebih baik lambat tapi aman daripada cepat tapi bocor data! 🔒
