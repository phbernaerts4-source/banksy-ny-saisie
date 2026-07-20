# Banksy NYC — Daily Report (cloud)

Standalone Streamlit form for the NYC manager/cashier. Writes each day into the
dedicated **MasterSheet NY** Google Sheet (source of truth). No install for the
user — just a URL.

## What's in here
- `saisie_manager_ny.py` — the form (cloud mode: no local DB → writes to the Sheet)
- `Master_Template_XLS/MasterSheet_NY_Test.xlsx` — colored template reference
- `requirements.txt` — slim deps
- `.streamlit/config.toml` — theme
- `.streamlit/secrets.toml.example` — secrets template (real values go in the Cloud UI)

## Deploy on Streamlit Community Cloud

1. **Create a private GitHub repo** and push this folder:
   ```bash
   cd ~/Desktop/SaisieNY_Deploy
   git init && git add . && git commit -m "Banksy NYC daily report — cloud"
   git branch -M main
   git remote add origin https://github.com/<you>/banksy-ny-saisie.git
   git push -u origin main
   ```
   (`.streamlit/secrets.toml` is gitignored — it will NOT be pushed. Good.)

2. Go to **https://share.streamlit.io** → sign in with GitHub → **Create app**.
   - Repository: `<you>/banksy-ny-saisie`
   - Branch: `main` · Main file path: `saisie_manager_ny.py`

3. **Advanced settings → Secrets**: paste the entire contents of your local
   `.streamlit/secrets.toml` (access_code + gsheet_id + [gcp_token]).

4. **Deploy**. You get a URL like `https://banksy-ny-saisie.streamlit.app`.
   Send it (with the access code) to the manager.

## Notes
- The Google Sheet is owned by you; the app authenticates as you (OAuth token in
  secrets). The manager/cashier never touches Google.
- Update the form later = `git push` → auto-redeploy.
