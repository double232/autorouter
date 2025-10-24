# IMAP Authentication Troubleshooting

## Error: "login command bad error 12"

This means Office 365 is rejecting your password. This is VERY common and easy to fix.

---

## ✅ SOLUTION 1: App-Specific Password (Recommended)

Office 365 requires app-specific passwords for IMAP when you have 2FA enabled or certain security settings.

### Step-by-Step:

1. **Go to Microsoft Account Security**
   - Visit: https://account.microsoft.com/security
   - Sign in with your email

2. **Find Advanced Security Options**
   - Click "Advanced security options"
   - Scroll down to "App passwords"

3. **Create App Password**
   - Click "Create a new app password"
   - You'll see a randomly generated password like: `abcd-efgh-ijkl-mnop`
   - **Copy this password immediately** (you can't see it again)

4. **Use in the GUI**
   - Open Trial Orders Automation
   - Configuration tab
   - Email Password field: Paste the app password
   - Save Configuration
   - Test Connection

### Example:
```
Email Address: john@eazlegal.com
Email Password: abcd-efgh-ijkl-mnop  ← Use this, not your regular password!
```

---

## ✅ SOLUTION 2: Enable IMAP in Outlook

Sometimes IMAP is disabled by default.

### Step-by-Step:

1. **Open Outlook Web (OWA)**
   - Go to: https://outlook.office.com
   - Sign in

2. **Open Settings**
   - Click the gear icon (⚙️) in top right
   - Click "View all Outlook settings"

3. **Enable IMAP**
   - Go to: Mail → Sync email
   - Check "Let devices and apps use IMAP"
   - Click Save

4. **Test Again**
   - Go back to Trial Orders Automation
   - Test Connection

---

## ✅ SOLUTION 3: Check IMAP Settings

Make sure your IMAP settings are correct:

### For Office 365/Outlook.com:
```
IMAP Server: outlook.office365.com
IMAP Port: 993
```

### For Gmail (if using Gmail instead):
```
IMAP Server: imap.gmail.com
IMAP Port: 993
```

---

## ✅ SOLUTION 4: Disable MFA Temporarily (Not Recommended)

If you can't create an app password:

1. Go to https://account.microsoft.com/security
2. Disable two-factor authentication temporarily
3. Try logging in with regular password
4. Re-enable 2FA after testing

**Note:** This is less secure. App-specific password is better.

---

## Still Not Working?

### Check Your Email Provider

**If you're using:**

**Office 365 / Outlook.com:**
- Use Solution 1 (App-Specific Password)
- IMAP Server: `outlook.office365.com`
- Port: `993`

**Gmail:**
- Enable "Less secure app access" OR use app password
- IMAP Server: `imap.gmail.com`
- Port: `993`

**Custom Domain (like @eazlegal.com):**
- Ask IT admin what IMAP server to use
- Ask if app-specific passwords are required
- May need to enable IMAP in admin console

---

## Testing Your IMAP Credentials

### Quick Test (Windows):

Open PowerShell and test:

```powershell
# Test IMAP connection
Test-NetConnection -ComputerName outlook.office365.com -Port 993
```

If this fails, your firewall might be blocking IMAP.

### Full Test with Python:

Create a file `test_imap.py`:

```python
import imaplib

# Your credentials
EMAIL = "your-email@domain.com"
PASSWORD = "your-app-password"  # Use app password!
SERVER = "outlook.office365.com"
PORT = 993

try:
    print(f"Connecting to {SERVER}:{PORT}...")
    imap = imaplib.IMAP4_SSL(SERVER, PORT)
    print("✅ Connected!")

    print(f"Logging in as {EMAIL}...")
    imap.login(EMAIL, PASSWORD)
    print("✅ Login successful!")

    print("Selecting INBOX...")
    imap.select("INBOX")
    print("✅ INBOX selected!")

    print("\n🎉 All tests passed! Your IMAP credentials work!")

    imap.logout()

except imaplib.IMAP4.error as e:
    print(f"❌ IMAP Error: {e}")
    print("\nTry creating an app-specific password!")
except Exception as e:
    print(f"❌ Error: {e}")
```

Run:
```bash
python test_imap.py
```

---

## Common Error Messages

| Error | What It Means | Solution |
|-------|---------------|----------|
| `bad error 12` | Wrong password | Use app-specific password |
| `AUTHENTICATIONFAILED` | Login rejected | Use app-specific password |
| `LOGIN disabled` | Basic auth off | Enable IMAP or use OAuth2 |
| `Connection refused` | Can't reach server | Check IMAP server address |
| `timeout` | Firewall blocking | Check port 993 is open |

---

## Alternative: Microsoft Graph API (Advanced)

If IMAP absolutely won't work, you could use Microsoft Graph API (requires Azure AD setup).

Let me know if you need this option - it's more complex but more reliable.

---

## Need Help?

1. Try Solution 1 first (app-specific password)
2. If that doesn't work, try Solution 2 (enable IMAP)
3. If still stuck, contact IT support with this error:
   - "IMAP authentication failing with error 12"
   - "Need app-specific password or IMAP enabled"

---

**90% of the time, creating an app-specific password fixes this!**

Go to: https://account.microsoft.com/security → Advanced security options → Create app password
