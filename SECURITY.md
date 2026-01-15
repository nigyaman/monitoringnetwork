# Security Policy

## Reporting Security Issues

If you discover a security vulnerability in this project, please report it by:

1. **DO NOT** create a public GitHub issue
2. Email the maintainer directly with details
3. Include steps to reproduce the vulnerability
4. Allow reasonable time for the issue to be addressed

## Security Best Practices

### Credential Management

**CRITICAL**: Never commit credentials to version control!

#### Protected Files (Already in .gitignore):
- `_telkom_access.xml` - Contains TACACS+ and router credentials
- `.env` - Local environment configuration
- `*.log` - May contain sensitive operational data
- `*.xlsx` - Reports may contain network topology information

#### Before Committing:
```bash
# Always verify what you're about to commit
git status
git diff

# Ensure no sensitive files are tracked
git ls-files | grep -E "_telkom_access.xml|\.env$|\.log$|\.xlsx$"
# Output should be EMPTY!
```

### Credential Storage

1. **Use Template Files**:
   - `_telkom_access.xml.example` - Template for credentials
   - `.env.example` - Template for environment variables
   
2. **Local Setup**:
   ```bash
   # Copy templates and configure with your credentials
   cp _telkom_access.xml.example _telkom_access.xml
   cp .env.example .env
   
   # Set restrictive permissions (Linux/Mac)
   chmod 600 _telkom_access.xml
   chmod 600 .env
   ```

3. **Never**:
   - Commit actual credentials
   - Share credentials via email/chat
   - Store passwords in plaintext outside secure locations
   - Include credentials in screenshots or documentation

### Network Security

1. **SSH Connections**:
   - Always use encrypted connections
   - Verify host keys
   - Use strong authentication
   - Monitor connection logs

2. **Access Control**:
   - Apply least privilege principle
   - Rotate credentials regularly
   - Audit access logs periodically
   - Restrict file access with proper permissions

3. **Data Protection**:
   - Reports may contain sensitive network topology
   - Store reports in secure locations
   - Delete old reports regularly
   - Encrypt reports if transmitting over network

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 25.x    | :white_check_mark: |
| < 25.0  | :x:                |

## Security Updates

Security patches will be released as soon as possible after a vulnerability is confirmed and fixed. Users should update to the latest version promptly.

## Audit Trail

- All credential access should be logged
- Review logs regularly for suspicious activity
- Implement alerting for authentication failures
- Maintain audit trail for compliance

## Compliance

This tool should be used in accordance with:
- Your organization's security policies
- Applicable data protection regulations
- Network access control policies
- Change management procedures

## Contact

For security concerns, please contact the repository maintainer.

---

**Remember**: Security is everyone's responsibility. Always verify before committing!
