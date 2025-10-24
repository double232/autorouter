# Trial Orders Automation - Multi-Provider AI Edition

Professional desktop application for automated court document processing with AI-powered date extraction.

**Version 2.0** - No Azure Required | Multiple AI Providers | vLLM Support

---

## Quick Start (60 seconds)

### Step 1: Double-click `Launch.bat`

That's it! The launcher will:
- ✅ Check if Python is installed
- ✅ Install dependencies automatically
- ✅ Launch the GUI

### Step 2: Configure (first time only)

1. Click the **Configuration** tab
2. Enter your credentials:
   - **Email (IMAP)**: Your Office 365 email and password
   - **SharePoint**: Your SharePoint username and password
   - **AI Provider**: Choose one:
     - **Claude** (Anthropic) - Best for documents
     - **OpenAI** (GPT-4o) - Strong vision
     - **Gemini** (Google) - Large context
     - **vLLM** (Self-hosted) - Free, private, fast
3. Click **Save Configuration**
4. Click **Test Connection** to verify

### Step 3: Run Automation

1. Click the **Main** tab
2. Click **▶ Start Processing**
3. Watch the magic happen!

---

## What's New in v2.0

### No Azure AD Required!
- ✅ Uses standard IMAP for email (no app registration)
- ✅ Uses SharePoint REST API with your credentials
- ✅ No tenant ID, client ID, or secrets needed
- ✅ Works with any Office 365 account

### Multiple AI Providers
Choose the best AI for your needs:

| Provider | Best For | Cost | Privacy |
|----------|----------|------|---------|
| **Claude 3.5** | Document understanding | ~$0.003/page | Cloud |
| **GPT-4o** | Vision & reasoning | ~$0.01/page | Cloud |
| **Gemini 1.5 Pro** | Large documents | ~$0.001/page | Cloud |
| **vLLM (Self-hosted)** | Privacy & cost | FREE* | Complete |

*After initial hardware setup

### vLLM - The Game Changer

Run AI **locally on your own hardware**:

**Advantages:**
- 💰 **Zero API costs** after setup
- 🔒 **Complete data privacy** (never leaves your network)
- 🚀 **No rate limits** (as fast as your GPU)
- 🎛️ **Full control** over models and parameters
- 🌐 **Works offline** (no internet needed)

**Recommended Models:**
- **Qwen2-VL-72B** - Excellent vision, best quality
- **Qwen2-VL-7B** - Fast, good quality, lower GPU requirements
- **LLaVA-NeXT** - Open source, solid performance
- **CogVLM2** - Strong document understanding

**Hardware Requirements:**
- **For 7B models**: RTX 4090 (24GB VRAM) or similar
- **For 72B models**: 2x RTX 4090 or A100 GPU
- **CPU fallback**: Slower but works without GPU

**Setup Guide**: See [docs.vllm.ai](https://docs.vllm.ai)

---

## What It Does

### Fully Automated Workflow:

1. **📧 Email Monitoring**: Connects via IMAP to your Office 365 inbox
2. **🔍 Smart Filtering**: Finds court document emails (subject: "SERVICE OF COURT DOCUMENT")
3. **📥 PDF Download**: Extracts download links from email body and downloads PDFs
4. **🤖 AI Processing**: Your chosen AI extracts:
   - Calendar Call date
   - Trial Start date
   - Trial End date
   - Document type (CMO/UTO/Other)
5. **📁 SharePoint Filing**: Saves to case-specific folders:
   ```
   /Cases/[Client]/[Matter] - [Style]/09 Orders/
   ```
6. **📊 Record Tracking**: Creates tracking records with all extracted data
7. **✅ Cleanup**: Marks emails as processed

---

## Features

### 🎨 Beautiful GUI
- Clean, modern interface
- Real-time activity log with color coding
- Progress indicators
- Statistics dashboard
- AI provider selection dropdown

### 🤖 Multiple AI Providers
- **Claude 3.5 Sonnet** - Best document understanding
- **GPT-4o** - Excellent vision capabilities
- **Gemini 1.5 Pro** - 1M+ token context window
- **vLLM** - Self-hosted, unlimited usage

### ⚙️ Easy Configuration
- Save/load settings
- Test connection button
- No command line needed
- Provider-specific settings shown/hidden automatically

### 🔒 Secure
- Credentials saved locally only
- Passwords hidden in UI
- No cloud storage of credentials
- vLLM option for complete data privacy

### 📊 Statistics
- Emails processed count
- PDFs downloaded count
- Last run timestamp
- Real-time status updates

---

## Requirements

- **Windows 10/11** (or macOS/Linux)
- **Python 3.9+**
- **Internet connection** (unless using vLLM locally)

### Email Access (IMAP)
- Office 365 email account
- IMAP enabled (usually enabled by default)
- Email password or app-specific password

### SharePoint Access
- SharePoint username (usually same as email)
- SharePoint password (usually same as email)
- Access to the LitigationOperations site

### AI Provider (Choose One)

#### Option 1: Claude (Anthropic)
Get API key from: [console.anthropic.com](https://console.anthropic.com/)
- Cost: ~$0.003 per page
- Best for: Complex legal documents

#### Option 2: OpenAI
Get API key from: [platform.openai.com](https://platform.openai.com/)
- Cost: ~$0.01 per page
- Best for: Vision and general tasks

#### Option 3: Google Gemini
Get API key from: [ai.google.dev](https://ai.google.dev/)
- Cost: ~$0.001 per page
- Best for: Large documents (1M+ tokens)

#### Option 4: vLLM (Self-Hosted)
Setup guide: [docs.vllm.ai](https://docs.vllm.ai)
- Cost: Free (after hardware)
- Best for: Privacy, high volume, cost savings

---

## File Structure

```
TrialOrdersAutomation/
├── Launch.bat          ← Double-click to start
├── gui.py              ← GUI application (supports all providers)
├── automation.py       ← Core automation logic (multi-provider)
├── requirements.txt    ← Python dependencies
├── config.json         ← Your saved settings (created on first save)
└── README.md          ← This file
```

---

## Setting Up vLLM (Optional)

### Why vLLM?

If you process many documents per month, vLLM can save significant costs:
- **100 documents/month**: Save ~$30/month vs Claude
- **1000 documents/month**: Save ~$300/month
- **Complete privacy**: Documents never leave your network

### Quick Setup (Ubuntu/Windows WSL2)

```bash
# Install vLLM
pip install vllm

# Download model (one-time, ~40GB for 7B model)
# This will auto-download on first use

# Start vLLM server
vllm serve Qwen/Qwen2-VL-7B-Instruct \
  --host 0.0.0.0 \
  --port 8000 \
  --dtype auto \
  --api-key EMPTY

# In the GUI, configure:
# - AI Provider: vllm
# - vLLM Base URL: http://localhost:8000/v1
# - Model: Qwen/Qwen2-VL-7B-Instruct
```

### Hardware Recommendations

| Model | VRAM | GPU Example | Speed | Quality |
|-------|------|-------------|-------|---------|
| Qwen2-VL-7B | 16GB | RTX 4080 | Fast | Good |
| Qwen2-VL-7B | 24GB | RTX 4090 | Very Fast | Good |
| Qwen2-VL-72B | 80GB | A100 | Medium | Excellent |
| Qwen2-VL-72B | 48GB | 2x RTX 4090 | Medium | Excellent |

---

## Scheduling Automation

### Windows Task Scheduler

1. Open **Task Scheduler**
2. **Create Basic Task**
3. Name: "Trial Orders Automation"
4. Trigger: **Daily** at 9:00 AM
5. Action: **Start a program**
   - Program: `Launch.bat`
   - Start in: `C:\Users\...\TrialOrdersAutomation`

The automation will run automatically every day!

---

## Troubleshooting

### "Python not found"
**Solution**: Install Python from [python.org](https://python.org) and check "Add to PATH"

### "Module not found" error
**Solution**: Run:
```bash
python -m pip install -r requirements.txt
```

### "Authentication failed" (Email)
**Solution**:
1. Verify your email and password are correct
2. If using 2FA, create an app-specific password
3. Ensure IMAP is enabled in your email settings

### "SharePoint connection failed"
**Solution**:
1. Verify username and password
2. Test access by opening SharePoint site in browser
3. Check if your account has access to LitigationOperations site

### Connection test fails
**Solution**:
1. Check internet connection
2. Verify email/SharePoint credentials
3. Check IMAP server (outlook.office365.com) and port (993)

### No emails processed
**Solution**:
1. Check inbox for unread emails with "SERVICE OF COURT DOCUMENT" in subject
2. Verify emails are unread
3. Check the Activity Log for errors

### vLLM connection error
**Solution**:
1. Ensure vLLM server is running (`vllm serve ...`)
2. Check Base URL (http://localhost:8000/v1)
3. Verify model name matches what you started vLLM with
4. Check GPU memory isn't exhausted

### AI extraction errors
**Solution**:
1. Try a different AI provider (dropdown in config)
2. Check API key is valid and has credits
3. For vLLM, ensure model is fully loaded
4. Review Activity Log for specific error messages

---

## Cost Comparison

### Monthly Cost Examples (100 documents × 3 pages = 300 pages)

| Provider | Cost | Notes |
|----------|------|-------|
| **Claude** | ~$0.90/month | Best quality for legal docs |
| **GPT-4o** | ~$3.00/month | Good general purpose |
| **Gemini** | ~$0.30/month | Most affordable cloud option |
| **vLLM** | **FREE** | After initial GPU investment |

### vLLM Break-Even Analysis

| Usage | Break-Even Time | Savings/Year |
|-------|----------------|--------------|
| 100 docs/month | ~24 months* | $10.80 |
| 500 docs/month | ~5 months* | $54.00 |
| 1000 docs/month | ~2.5 months* | $108.00 |

*Assuming used RTX 4090 (~$1200)

---

## Support

### Getting Help:

1. Check the **Activity Log** tab for error details
2. Click **Test Connection** to verify setup
3. Review this README
4. Contact your IT administrator

### Common Issues:

| Issue | Solution |
|-------|----------|
| No configuration saved | Click Save Configuration |
| Connection fails | Check credentials in config tab |
| No PDFs downloaded | Verify email links are valid |
| SharePoint upload fails | Check folder structure exists |
| AI extraction fails | Try different AI provider |
| vLLM not responding | Check if server is running |

---

## Comparison: GUI v2.0 vs Power Automate

| Feature | GUI v2.0 | Power Automate |
|---------|----------|----------------|
| **Setup** | 5 minutes | 15+ minutes |
| **Azure Required** | ❌ No | ✅ Yes |
| **Interface** | Desktop app | Web browser |
| **AI Options** | 4 providers | AI Builder only |
| **AI Quality** | Excellent (Claude/vLLM) | Good (AI Builder) |
| **Debugging** | Easy (live logs) | Difficult |
| **Cost** | $0-3/month | Included (limited runs) |
| **Control** | Full | Limited |
| **Offline** | Yes (with vLLM) | No |
| **Privacy** | Complete (vLLM) | Microsoft cloud |
| **Scalability** | Unlimited (vLLM) | Run limits |

---

## Version History

**v2.0** (2025-10-23)
- 🎉 Removed Azure AD requirement
- 🤖 Added multi-provider AI support (Claude, OpenAI, Gemini, vLLM)
- 🔒 Added vLLM self-hosted option
- 📧 Switched to IMAP for email access
- 🌐 Switched to SharePoint REST API
- 🎨 Enhanced GUI with provider selection
- 📖 Updated documentation

**v1.0** (2025-10-23)
- Initial release
- GUI application
- Claude AI integration
- SharePoint filing
- Email automation

---

## License

Internal use only - Vernis & Bowling Law Firm / EAZ Legal PLLC

---

## Credits

**Powered by:**
- Python & Tkinter
- IMAP (email access)
- SharePoint REST API
- Anthropic Claude AI / OpenAI / Google Gemini / vLLM

**AI Models:**
- Claude 3.5 Sonnet (Anthropic)
- GPT-4o (OpenAI)
- Gemini 1.5 Pro (Google)
- Qwen2-VL / LLaVA-NeXT / CogVLM (vLLM)

**Built for:**
EAZ Legal PLLC Litigation Operations Team

---

**🚀 Ready to automate your trial orders?**

**Choose your path:**
- **Quick & Easy**: Use Claude/OpenAI/Gemini cloud APIs
- **Maximum Privacy & Savings**: Set up vLLM locally

**Get started:** Double-click `Launch.bat`!

---

## FAQ

**Q: Do I need Azure?**
A: No! Version 2.0 removed this requirement. Just use your regular email and SharePoint credentials.

**Q: Which AI provider should I choose?**
A: Start with Claude (best quality) or Gemini (best price). Try vLLM if you want privacy or high volume.

**Q: Is my data safe?**
A: With cloud providers (Claude/OpenAI/Gemini), documents are sent to their APIs. With vLLM, everything stays on your hardware.

**Q: Can I switch AI providers?**
A: Yes! Just select a different provider in the Configuration tab and enter the required credentials.

**Q: What GPU do I need for vLLM?**
A: Start with an RTX 4090 (24GB) for 7B models. For 72B models, you'll need 2x RTX 4090 or an A100.

**Q: Does vLLM work without a GPU?**
A: Yes, but it's very slow. GPU is strongly recommended.

**Q: How do I get an app-specific password for email?**
A: In Outlook/Office 365, go to Security settings → App passwords → Create new app password.
