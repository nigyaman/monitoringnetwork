# Contributing to Network Monitoring Tool

Terima kasih atas minat Anda untuk berkontribusi! Dokumen ini berisi panduan untuk berkontribusi ke project ini.

## 🚀 Cara Berkontribusi

### 1. Fork dan Clone
```bash
# Fork repository di GitHub
# Kemudian clone fork Anda
git clone https://github.com/YOUR_USERNAME/monitoing.git
cd monitoing
```

### 2. Setup Development Environment
```bash
# Install dependencies
pip install paramiko openpyxl

# Copy template credentials (JANGAN commit yang asli!)
cp _telkom_access.xml.example _telkom_access.xml
cp .env.example .env

# Edit dengan kredensial testing Anda
```

### 3. Buat Branch Baru
```bash
git checkout -b feature/nama-fitur-anda
# atau
git checkout -b fix/nama-bug-yang-diperbaiki
```

### 4. Lakukan Perubahan
- Tulis kode yang bersih dan terdokumentasi
- Ikuti style guide Python (PEP 8)
- Tambahkan docstring untuk fungsi baru
- Test perubahan Anda secara menyeluruh

### 5. Commit Changes
```bash
# PENTING: Pastikan tidak ada kredensial!
git status
git diff

# Commit dengan pesan yang jelas
git add .
git commit -m "feat: tambah fitur X untuk meningkatkan Y"
```

### 6. Push dan Create Pull Request
```bash
git push origin feature/nama-fitur-anda
```

Kemudian buat Pull Request di GitHub dengan deskripsi yang jelas.

## 📝 Code Style Guidelines

### Python Code Style
- Ikuti [PEP 8](https://pep8.org/)
- Gunakan 4 spasi untuk indentasi
- Maksimum 100 karakter per baris
- Gunakan nama variabel yang deskriptif

### Docstring Format
```python
def function_name(param1, param2):
    """
    Brief description of function.
    
    Args:
        param1 (type): Description of param1
        param2 (type): Description of param2
    
    Returns:
        type: Description of return value
    
    Raises:
        ErrorType: Description of when error occurs
    """
    pass
```

### Commit Message Format
Gunakan conventional commits:
- `feat:` untuk fitur baru
- `fix:` untuk bug fixes
- `docs:` untuk perubahan dokumentasi
- `refactor:` untuk refactoring code
- `test:` untuk menambah tests
- `chore:` untuk maintenance tasks

Contoh:
```
feat: add support for custom timeout configuration
fix: resolve connection timeout issue on slow networks
docs: update README with new installation steps
```

## 🔒 Security Checklist

**CRITICAL**: Sebelum commit atau push, selalu verifikasi:

```bash
# 1. Check status - pastikan tidak ada file sensitif
git status

# 2. Review changes - pastikan tidak ada credentials
git diff

# 3. Verify no sensitive files are tracked
git ls-files | grep -E "_telkom_access.xml|\.env$|\.log$"
# Output harus KOSONG!

# 4. Check .gitignore is working
git check-ignore -v _telkom_access.xml
# Harus menunjukkan bahwa file di-ignore
```

### File yang TIDAK BOLEH di-commit:
- ❌ `_telkom_access.xml` (kredensial aktual)
- ❌ `.env` (local config)
- ❌ `*.log` (log files)
- ❌ `*.xlsx` (reports)
- ❌ Any file with credentials or passwords

### File yang BOLEH di-commit:
- ✅ `_telkom_access.xml.example` (template)
- ✅ `.env.example` (template)
- ✅ Source code (`*.py`)
- ✅ Documentation (`*.md`)
- ✅ Configuration templates

## 🧪 Testing

### Manual Testing
Sebelum submit PR, pastikan:
1. Script berjalan tanpa error
2. Semua fitur yang terpengaruh masih berfungsi
3. Tidak ada regression pada fitur existing
4. Output format tetap konsisten

### Test Checklist
- [ ] Script runs without errors
- [ ] Connection to devices successful
- [ ] Excel reports generated correctly
- [ ] Logging works properly
- [ ] Error handling works as expected
- [ ] No credentials in debug output
- [ ] Documentation updated

## 📚 Documentation

Jika menambah fitur baru:
1. Update README.md
2. Tambah docstring yang jelas
3. Update changelog jika perlu
4. Tambah contoh penggunaan

## 🐛 Reporting Bugs

### Before Reporting
1. Check existing issues
2. Verify it's not a configuration issue
3. Test with latest version

### Bug Report Should Include
- Python version dan OS
- Complete error message
- Steps to reproduce
- Expected vs actual behavior
- Log files (with sensitive data removed!)
- Configuration (without credentials!)

## 💡 Suggesting Features

Feature requests welcome! Please include:
- Clear description of the feature
- Use case / problem it solves
- Examples of how it would work
- Any implementation ideas

## 📋 Pull Request Process

1. **Fork dan branch** dari `main`
2. **Update documentation** jika diperlukan
3. **Test thoroughly** di environment Anda
4. **Security check** - no credentials!
5. **Submit PR** dengan deskripsi yang jelas
6. **Respond to feedback** dari maintainers

### PR Checklist
- [ ] Code follows style guidelines
- [ ] Documentation updated
- [ ] No credentials or sensitive data
- [ ] Tested locally
- [ ] Commit messages are clear
- [ ] Branch is up to date with main

## 🤝 Code Review

Semua contributions akan di-review untuk:
- Code quality dan style
- Security considerations
- Performance implications
- Documentation completeness
- Test coverage

## 📞 Getting Help

Butuh bantuan? Anda bisa:
1. Check existing documentation
2. Open an issue untuk diskusi
3. Ask in PR comments

## 🙏 Recognition

Contributors akan di-list di README. Terima kasih atas kontribusi Anda!

---

**Remember**: Keamanan adalah prioritas #1. Selalu double-check sebelum commit!
