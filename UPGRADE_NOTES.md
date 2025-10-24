# Trial Orders Automation v2.0 - Upgrade Notes

## What Changed

### Big Changes

**1. NO AZURE AD REQUIRED!**
- Old: Required Azure AD app registration (Tenant ID, Client ID, Secret)
- New: Just use your email and SharePoint password

**2. MULTIPLE AI PROVIDERS**
- Old: Only Claude (Anthropic)
- New: Choose from 4 options:
  - Claude (Anthropic)
  - OpenAI (GPT-4o)
  - Google Gemini
  - vLLM (Self-hosted)

**3. ADDED vLLM SUPPORT**
- Run AI locally on your own hardware
- Zero API costs after setup
- Complete data privacy
- No rate limits

### Technical Changes

**Email Access:**
- Old: Microsoft Graph API (requires Azure AD)
- New: IMAP (standard email protocol)

**SharePoint Access:**
- Old: Microsoft Graph API (requires Azure AD)
- New: SharePoint REST API with user credentials

**Authentication:**
- Old: OAuth2 with app registration
- New: Simple username/password

---

## What You Need to Do

### If Upgrading from v1.0:

1. **Install new dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Update configuration:**
   - Open the GUI
   - Go to Configuration tab
   - Remove old Azure AD settings
   - Add:
     - Email address and password
     - SharePoint username and password
     - Choose AI provider
     - Enter API key for chosen provider (or vLLM URL)

3. **Save and test:**
   - Click "Save Configuration"
   - Click "Test Connection"

### If Starting Fresh:

Just follow the Quick Start in README.md!

---

## Configuration Migration

### Old Config Fields (REMOVE):
- ❌ Tenant ID
- ❌ Client ID
- ❌ Client Secret

### New Config Fields (ADD):
- ✅ Email Address
- ✅ Email Password
- ✅ IMAP Server (default: outlook.office365.com)
- ✅ IMAP Port (default: 993)
- ✅ SharePoint Username
- ✅ SharePoint Password
- ✅ AI Provider (dropdown: claude/openai/gemini/vllm)
- ✅ Provider-specific settings (API keys or vLLM URL)

---

## Benefits of v2.0

### Easier Setup
- No Azure AD app registration
- No API permissions to configure
- Just email + password = ready to go

### More Flexible
- Choose AI provider based on needs
- Switch providers anytime
- Try different models

### More Cost-Effective
- Gemini: ~1/3 the cost of Claude
- vLLM: FREE after hardware investment

### More Private
- vLLM: documents never leave your network
- No data sent to cloud (with vLLM)
- Full control over AI processing

---

## Breaking Changes

### Configuration File
The `config.json` format has changed. Old configs will not work.
Just re-enter your settings in the new GUI.

### Dependencies
New Python packages required:
- `Office365-REST-Python-Client` (for SharePoint)
- `openai` (for OpenAI and vLLM)
- `google-generativeai` (for Gemini)

Old packages removed:
- `msal` (Azure AD authentication)

---

## Recommended AI Provider

**For most users:** Start with **Gemini**
- Cheapest cloud option (~$0.30/month for 100 docs)
- Good quality
- Large context window
- Easy setup

**For best quality:** Use **Claude**
- Best document understanding
- Slightly more expensive (~$0.90/month)
- Anthropic API key required

**For privacy & high volume:** Set up **vLLM**
- One-time GPU investment
- Zero ongoing costs
- Complete data privacy
- Requires technical setup

**For vision tasks:** Use **GPT-4o**
- Strong multimodal capabilities
- Mid-range pricing (~$3/month)
- OpenAI API key required

---

## vLLM Quick Start

If you have a GPU (RTX 4080/4090 or better):

```bash
# Install vLLM
pip install vllm

# Start server (one command)
vllm serve Qwen/Qwen2-VL-7B-Instruct --host 0.0.0.0 --port 8000

# Configure GUI
# AI Provider: vllm
# vLLM Base URL: http://localhost:8000/v1
# Model: Qwen/Qwen2-VL-7B-Instruct
```

That's it! Now you have FREE, PRIVATE AI processing.

---

## Support

If you have issues:

1. Check the Activity Log in the GUI
2. Click "Test Connection" in Configuration tab
3. Review README.md
4. Try a different AI provider
5. Contact IT support

---

## FAQ

**Q: Do I need to re-install?**
A: Just run `pip install -r requirements.txt` to update dependencies.

**Q: Will my old config work?**
A: No, you'll need to re-enter settings in the new GUI format.

**Q: Can I keep using Claude?**
A: Yes! Just select "claude" in the AI Provider dropdown.

**Q: Should I try vLLM?**
A: If you have a good GPU and process many documents, definitely worth trying!

**Q: Is this tested?**
A: Code is ready, but you'll need to test with your credentials.

---

**Ready to upgrade?** Run `Launch.bat` and configure your new settings!
