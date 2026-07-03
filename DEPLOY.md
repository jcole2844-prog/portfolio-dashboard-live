# Deploying the Portfolio Dashboard to the web (access from anywhere + iPad)

> **Data source of truth: the Google Sheet "Portfolio Dashboard"** (tabs:
> Holdings, Watchlist, Fixed Income). The dashboard reads — and, via the in-app
> editor, writes — to this Sheet. You can also edit the Sheet directly (e.g. the
> Google Sheets app on your iPad) and the dashboard picks it up. The local
> `.xlsx` files in Downloads are only an offline fallback if the Sheet is
> unreachable.


This puts the dashboard online at a private `https://` URL, protected by Google
login, reading your Excel files from Google Drive. You open it in Safari on your
iPad (or any browser) and sign in with your Google account.

There are 4 one-time setup stages. Budget ~45–60 minutes the first time.

---

## Stage 1 — Put your two Excel files in Google Drive

1. In Google Drive, create a folder (e.g. **Portfolio Dashboard Data**) and put
   both workbooks in it:
   - `Holdings Listing ……xlsx`
   - `CD_Holdings.xlsx`
2. For **each** file, get its **file ID**: right-click → **Share** → **Copy link**.
   The link looks like
   `https://drive.google.com/file/d/`**`1AbCdEfGhIjKlMnOpQrStUvWxYz`**`/view`.
   The bold part is the **file ID** — save both IDs.
3. You do **not** need to make them public. Keep them private; we'll grant access
   to a service account in Stage 2.

> To update your holdings later, just replace/overwrite the file in Drive — the
> dashboard reads the latest version automatically (cached up to ~1 minute).

---

## Stage 2 — Create a Google service account (so the app can read the files)

1. Go to <https://console.cloud.google.com/> and create a project (e.g.
   **portfolio-dashboard**).
2. Enable the **Google Drive API**: APIs & Services → Library → search
   "Google Drive API" → **Enable**.
3. Create the service account: IAM & Admin → **Service Accounts** → **Create**.
   Name it e.g. `portfolio-dashboard`. Skip the optional role steps → **Done**.
4. Open the service account → **Keys** → **Add Key** → **Create new key** →
   **JSON** → **Create**. A `.json` file downloads — keep it safe; this is the
   credential.
5. Copy the service account's **email** (looks like
   `portfolio-dashboard@your-project.iam.gserviceaccount.com`).
6. Back in Google Drive, **Share each Excel file** (or the whole folder) with
   that service-account email, **Viewer** access. This is what lets the app read
   your private files.

---

## Stage 3 — Put the code on GitHub

1. Create a **private** repo on GitHub (e.g. `portfolio-dashboard`).
2. Upload the contents of the `portfolio_dashboard` folder to it:
   `app.py`, `data_feed.py`, `requirements.txt`, and the `.streamlit/` folder
   (with `config.toml` and `secrets.toml.example` — **but NOT** a real
   `secrets.toml`). The included `.gitignore` already blocks secrets and `.xlsx`
   files from being committed.
3. Do **not** commit your Excel files or the service-account JSON.

---

## Stage 4 — Deploy on Streamlit Community Cloud (free) + Google login

1. Go to <https://share.streamlit.io>, sign in with GitHub, and authorize it to
   see your private repo.
2. **Create app** → pick your repo, branch `main`, main file path **`app.py`**.
   Under **Advanced settings**, choose **Python 3.12**.
3. Open **Advanced settings → Secrets** and paste your secrets in TOML form —
   use `.streamlit/secrets.toml.example` as the template:
   - `[gdrive_files]` → the two file IDs from Stage 1.
   - `[gdrive_service_account]` → the contents of the JSON key from Stage 2,
     converted to TOML (each JSON field becomes `key = "value"`; keep the
     `private_key` exactly, including the `\n` characters).
4. Click **Deploy**. First build takes a few minutes (installing packages).
5. Make it **private**: app menu (•••) → **Settings** → **Sharing** → set to
   *"Only specific people can view"* and add your Google email (and any others).
   Now only those emails can open it, and they sign in with Google.

You'll get a URL like `https://your-app.streamlit.app`. Open it on your iPad,
sign in with Google, and you're in. Add it to your iPad Home Screen (Safari →
Share → **Add to Home Screen**) for an app-like icon.

---

## Notes & gotchas

- **Local still works.** With no secrets present, the app reads the local files
  on your PC exactly as before — so `run_dashboard.bat` is unchanged.
- **yfinance rate limits.** Cloud servers share IPs, so Yahoo may throttle more
  than your home network. Expect occasional blank cells that fill on the next
  refresh. If it becomes a problem, a free **Finnhub** API key would make the
  data much steadier — ask and I'll wire it in.
- **iPad layout.** The wide tables scroll sideways with a finger-swipe. Cards
  stack/condense on the narrower screen. It's usable but designed desktop-first;
  tell me if you want a mobile-tuned layout.
- **Updating holdings.** Replace the file in Google Drive — no redeploy needed.
- **Updating the code.** Push to GitHub; Streamlit Cloud redeploys automatically.
</content>
