# 🇲🇦 Darija Translator - Cloud Deployment Guide

## Architecture

```
┌─────────────────────────────────────────────┐
│  FRONTEND (GitHub Pages - FREE)             │
│  https://ascertain.github.io/darija-translator │
│                                             │
│  • Static HTML/JS/CSS                       │
│  • MediaRecorder captures audio             │
│  • Calls Groq API directly from browser     │
│  • Calls HF Space for Whisper (optional)    │
└─────────────────┬───────────────────────────┘
                  │
        ┌─────────┴─────────┐
        ▼                   ▼
┌───────────────┐   ┌──────────────────────┐
│ Groq API      │   │ HF Space (optional)  │
│ (FREE)        │   │ (FREE)               │
│               │   │                      │
│ • Translation │   │ • Whisper AI         │
│ • Whisper STT │   │ • Audio → Text       │
│ • LLM (Llama) │   │ • Arabic dialects    │
└───────────────┘   └──────────────────────┘
```

**Total cost: $0** — Everything runs on free tiers.

---

## Step 1: Deploy Frontend (GitHub Pages)

The frontend is already in your portfolio repo.

```bash
cd C:\Users\MOKAS10\vcs\ascertain.github.io
git add darija-translator/
git commit -m "feat: add Darija Translator web app"
git push origin main
```

After push, it's live at: **https://ascertain.github.io/darija-translator/**

---

## Step 2: Get Groq API Key (Required)

1. Go to https://console.groq.com/keys
2. Sign up (free)
3. Create a new API key
4. Copy the key (starts with `gsk_...`)

**Groq provides FREE:**
- Whisper Large V3 transcription (speech → text)
- Llama 3.3 70B translation (text → text)
- 14,400 requests/day on free tier

---

## Step 3: Configure the App

1. Open https://ascertain.github.io/darija-translator/
2. Tap **⚙️ Setup**
3. Paste your Groq API key
4. Save

That's it! The app now works with:
- 🎤 Voice → Groq Whisper (transcription) → Groq Llama (translation) → 🔊 Speech
- ⌨️ Text → Groq Llama (translation) → 🔊 Speech

---

## Step 4 (Optional): Deploy Whisper Backend on Hugging Face

If you want a dedicated Whisper backend (for more control, custom models, or as backup):

### 4.1 Create a Hugging Face Space

1. Go to https://huggingface.co/spaces
2. Click **"Create new Space"**
3. Settings:
   - **Name:** `darija-backend`
   - **SDK:** Docker
   - **Hardware:** CPU Basic (free)
   - **Visibility:** Public
4. Click **"Create Space"**

### 4.2 Upload Backend Files

Upload these files from `darija-translator/hf-backend/`:
- `Dockerfile`
- `app.py`
- `requirements.txt`
- `README.md`

Or via git:
```bash
git clone https://huggingface.co/spaces/YOUR_USERNAME/darija-backend
cd darija-backend
cp /path/to/darija-translator/hf-backend/* .
git add . && git commit -m "Initial deploy" && git push
```

### 4.3 Configure Frontend to Use HF Space

1. Wait for HF Space to build (2-3 minutes)
2. Your Space URL will be: `https://YOUR_USERNAME-darija-backend.hf.space`
3. Open the translator → ⚙️ Setup → paste the URL in "HF Space URL" field
4. Save

Now the app uses your HF Space for Whisper transcription.

---

## How It Works (No Backend Needed!)

The **simplest setup** needs ONLY the Groq API key:

| Feature | Provider | Cost |
|---------|----------|------|
| Speech → Text | Groq Whisper API | FREE |
| Translation | Groq Llama 3.3 70B | FREE |
| Text → Speech | Browser SpeechSynthesis | FREE |
| Hosting | GitHub Pages | FREE |

**Groq's Whisper API** handles Arabic/Darija transcription directly from the browser.
No server needed. No HF Space needed (it's optional for advanced use).

---

## Security Notes

- API keys are stored in `localStorage` (your browser only)
- Keys are sent directly to Groq's API (HTTPS encrypted)
- No intermediate server sees your keys
- Nothing is stored on GitHub Pages (it's static HTML)
- Audio is processed and immediately discarded

---

## Limitations of GitHub Pages Version

| Feature | Local Version | GitHub Pages Version |
|---------|--------------|---------------------|
| Whisper model | Local (faster-whisper) | Groq API or HF Space |
| Translation | Argos (offline) | Groq API (online) |
| Dictionary | SQLite (2658 words) | None (API only) |
| Learning | TinyDB (NoSQL) | localStorage only |
| TTS | Edge-TTS (high quality) | Browser built-in |
| Offline mode | ✅ Full offline | ❌ Needs internet |
| Speed | ~1-2s | ~1-3s |

---

## Quick Commands

```bash
# Deploy frontend
cd C:\Users\MOKAS10\vcs\ascertain.github.io
git add darija-translator/ && git commit -m "update translator" && git push

# Test locally
cd darija-translator
python -m http.server 8000
# Open: http://localhost:8000
```
