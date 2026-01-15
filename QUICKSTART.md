# 🚀 Quick Start Guide - Network Monitoring Tool

## ⚡ Setup Cepat (5 Menit)

### 1. Clone Repository
```bash
git clone https://github.com/wokding/monitoringnetwork.git
cd monitoringnetwork
```

### 2. Install Dependencies
```bash
pip install paramiko openpyxl
```

### 3. Setup Credentials ⚠️ PENTING!
```bash
# Copy template
copy _telkom_access.xml.example _telkom_access.xml

# Edit dengan kredensial Anda (Windows)
notepad _telkom_access.xml

# Atau (Linux/Mac)
cp _telkom_access.xml.example _telkom_access.xml
nano _telkom_access.xml
```

### 4. Jalankan!
```bash
python fpc_utilisasi.py
```

---

## 📝 Checklist Sebelum Upload ke Git

**CRITICAL**: Selalu jalankan checklist ini sebelum `git push`:

```bash
# ✅ 1. Check status
git status

# ✅ 2. Verify .gitignore bekerja
git check-ignore _telkom_access.xml
# Output: _telkom_access.xml (GOOD!)

# ✅ 3. Pastikan tidak ada file sensitif
git ls-files | findstr /R "_telkom_access.xml$ .env$ .log$ .xlsx$"
# Output harus KOSONG!

# ✅ 4. Review changes
git diff

# ✅ 5. Commit hanya jika aman
git add .
git commit -m "Your commit message"
git push
```

---

## 🔒 File Sensitif yang DILARANG di-commit

| File | Status | Alasan |
|------|--------|--------|
| `_telkom_access.xml` | ❌ JANGAN | Berisi username & password |
| `.env` | ❌ JANGAN | Config lokal & secrets |
| `*.log` | ❌ JANGAN | Mungkin ada data operasional |
| `*.xlsx` | ❌ JANGAN | Laporan dengan info network |
| `reports/` | ❌ JANGAN | Output monitoring |

## ✅ File yang AMAN di-commit

| File | Status | Alasan |
|------|--------|--------|
| `_telkom_access.xml.example` | ✅ BOLEH | Template tanpa kredensial |
| `.env.example` | ✅ BOLEH | Template config |
| `*.py` | ✅ BOLEH | Source code |
| `*.md` | ✅ BOLEH | Dokumentasi |
| `.gitignore` | ✅ BOLEH | Git configuration |

---

## 🆘 Troubleshooting Cepat

### ❌ Error: "Missing dependency: paramiko"
```bash
pip install paramiko openpyxl
```

### ❌ Error: "Cannot find _telkom_access.xml"
```bash
copy _telkom_access.xml.example _telkom_access.xml
# Edit dengan kredensial Anda
```

### ❌ Accidentally committed credentials!
```bash
# SEGERA hapus dari Git history!
git rm --cached _telkom_access.xml
git commit -m "Remove accidentally committed credentials"
git push --force

# Kemudian rotasi credentials ASAP!
```

### ❌ Git push rejected
```bash
# Update local dulu
git pull origin main --rebase
git push
```

---

## 📚 Dokumentasi Lengkap

- **README.md** - Dokumentasi utama lengkap
- **SECURITY.md** - Panduan keamanan detail
- **CONTRIBUTING.md** - Cara berkontribusi
- **LICENSE** - MIT License

---

## 🎯 Tips Pro

1. **Selalu review** sebelum commit:
   ```bash
   git diff
   git status
   ```

2. **Test di environment non-production** dulu

3. **Rotasi credentials** secara berkala

4. **Backup reports** penting ke lokasi aman

5. **Monitor logs** untuk troubleshooting

---

## 📞 Butuh Bantuan?

1. Baca **README.md** untuk detail lengkap
2. Check **SECURITY.md** untuk masalah keamanan
3. Lihat **CONTRIBUTING.md** untuk kontribusi
4. Open issue di GitHub

---

**Remember**: Keamanan adalah prioritas #1! 🔐

Never commit credentials. Always double-check before `git push`.
