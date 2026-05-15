# PeelAnalyzer Web — Streamlit Edition

Run PeelAnalyzer entirely in your browser (Android Chrome, iPhone Safari, any device).

---

## 🚀 Deploy to Streamlit Community Cloud (free, 10 minutes)

### Step 1 — GitHub account
Go to **github.com** on your phone, sign up (free).

### Step 2 — Create a repository
1. Tap **+** → **New repository**
2. Name it `peel-analyzer-web`, set it to **Public**
3. Tap **Create repository**

### Step 3 — Upload the web files
In your new repository, tap **Add file → Upload files** and upload these files:
```
streamlit_app.py
data_utils_web.py
database_web.py
requirements.txt
```
Tap **Commit changes**.

### Step 4 — Deploy on Streamlit Community Cloud
1. Go to **share.streamlit.io** (sign in with GitHub)
2. Tap **New app**
3. Select your `peel-analyzer-web` repository
4. Main file path: `streamlit_app.py`
5. Tap **Deploy** — wait ~2 minutes
6. Your app is live at `https://yourname-peel-analyzer-web-xxxx.streamlit.app`

### Step 5 — Use on your phone
Open the URL in Chrome on Android. The app works fully in the browser.

---

## 📱 How to use

| Step | What to do |
|------|-----------|
| **Import** | Go to 🏠 Home → Upload your CSV files |
| **Analyse** | Go to 🔬 Analyse Tests → select a test, set smoothing + regions, click Analyse |
| **Batch** | In Analyse page, scroll down to Batch Operations |
| **Figures** | Go to 📊 Figure Designer → assign conditions → generate |
| **Save session** | Sidebar → Download session (.db) — upload it next time to resume |
| **Export** | In Analyse page, scroll to Export → Download Excel |

---

## 🗂️ Files in this package

### Web app (Streamlit — use these 4 files on Streamlit Cloud):
| File | Purpose |
|------|---------|
| `streamlit_app.py` | Main web app (all pages) |
| `data_utils_web.py` | Qt-free data processing |
| `database_web.py` | In-memory SQLite with CSV blob storage |
| `requirements.txt` | Python dependencies for Streamlit Cloud |

### Desktop app (PyQt6 — run on Windows/Mac/Linux with Python):
| File | Purpose |
|------|---------|
| `main.py` | Entry point — run this |
| `main_window.py` | Main window |
| `test_view.py` | Per-test analysis tab |
| `figure_designer.py` | Figure designer tab |
| `widgets.py` | Reusable Qt widgets |
| `batch_dialog.py` | Shared-parameter batch dialog |
| `sequential_dialog.py` | One-by-one analysis dialog |
| `normalize_dialog.py` | Width normalization dialog |
| `settings_dialog.py` | Settings dialog |
| `database.py` | Desktop SQLite database |
| `data_utils.py` | Desktop data utilities |
| `app_settings.py` | User preferences |
| `constants.py` | UI styles and constants |

### To run the desktop app:
```bash
pip install PyQt6 matplotlib numpy scipy pandas openpyxl
python main.py
```

---

## 💡 Session persistence

Streamlit Community Cloud has **no persistent storage** — your data lives in memory during the session. Always download your `.db` session file before closing the tab. Upload it again next time to restore all your tests, analyses, conditions, and figure configs.
