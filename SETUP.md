# 🚀 SETUP — Deploy the Tool

This guide is for whoever sets up the tool the first time. Once deployed, your team just opens a URL.

---

## Option A: Streamlit Community Cloud (Recommended — FREE)

### Step 1: Push code to GitHub

```bash
cd attitude_report_app

# Initialize git
git init
git add .
git commit -m "Initial commit"

# Create a new repo on GitHub (private recommended), then:
git remote add origin https://github.com/YOUR_USERNAME/attitude-report-generator.git
git branch -M main
git push -u origin main
```

### Step 2: Deploy on Streamlit Cloud

1. Go to https://share.streamlit.io
2. Sign in with GitHub
3. Click **"New app"**
4. Select your repo, branch `main`, and main file `app.py`
5. Click **Deploy**

In ~2 minutes, you'll get a URL like:
`https://attitude-report-generator.streamlit.app`

### Step 3: Share with team

Send the URL to your team. That's it.

**Optional security**: Add a simple password by editing `app.py` to require an env-var password
(see "Adding Password Protection" below).

---

## Option B: Local Computer (For Testing)

### Prerequisites

- Python 3.10+
- Tesseract OCR installed on your system:
  - **macOS**: `brew install tesseract tesseract-lang`
  - **Ubuntu/Debian**: `sudo apt install tesseract-ocr tesseract-ocr-chi-sim tesseract-ocr-eng`
  - **Windows**: Download installer from https://github.com/UB-Mannheim/tesseract/wiki

### Run

```bash
# Install Python deps
pip install -r requirements.txt

# Run
streamlit run app.py
```

Opens at `http://localhost:8501`.

---

## Google Drive Upload Setup

The app can upload the generated PPTX back into the same campaign Drive folder.
This requires Google OAuth because public folder links only allow reading.

Add these secrets in Streamlit Community Cloud:

```toml
[google_oauth]
client_id = "YOUR_GOOGLE_OAUTH_CLIENT_ID"
client_secret = "YOUR_GOOGLE_OAUTH_CLIENT_SECRET"
redirect_uri = "https://YOUR-STREAMLIT-APP.streamlit.app"
```

In Google Cloud Console, create an OAuth Client ID for a Web application and add
the same `redirect_uri` under Authorized redirect URIs. The Google account used
to connect must have edit permission on the campaign folder.

---

## Option C: Railway / Render (If You Outgrow Free Streamlit)

Streamlit Community Cloud has a small memory limit. If you generate big reports often,
consider Railway ($5/mo) or Render ($7/mo). They run the same code with no changes —
just connect the GitHub repo.

---

## 🔑 Adding Password Protection

If you want to restrict access, add this to the top of `app.py`:

```python
import streamlit as st
import os

# Simple password gate
PASSWORD = os.getenv("APP_PASSWORD", "attitude2026")

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    pwd = st.text_input("Password", type="password")
    if st.button("Login"):
        if pwd == PASSWORD:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("Wrong password")
    st.stop()
```

Then in Streamlit Cloud → App Settings → Secrets, add:
```
APP_PASSWORD = "your_secure_password"
```

---

## 🔄 Updating the Tool

To deploy a new version:

```bash
# Edit code locally, then:
git add .
git commit -m "Describe what you changed"
git push
```

Streamlit Cloud auto-redeploys in ~1 minute.

---

## 🆘 Troubleshooting Deployment

**"ModuleNotFoundError: pytesseract"**
→ Make sure `requirements.txt` includes `pytesseract` and `packages.txt` includes
  `tesseract-ocr` (Streamlit Cloud reads this for system packages).

**"Tesseract not found in PATH"**
→ Same as above. Streamlit Cloud uses `packages.txt` for apt packages.

**App is slow / crashes on big folders**
→ Streamlit Community Cloud has 1GB memory limit. Consider Railway/Render upgrade if
  reports have many large screenshots.

**"Cannot access folder" even though it's public**
→ Open the folder URL in a private/incognito browser window (logged out). If you can see
  the contents there, the tool can read it. If not, the share setting is wrong.
