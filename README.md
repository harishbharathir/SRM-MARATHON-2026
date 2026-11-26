# TSRM Run 2026 — Official Participant Ticket Portal

Web application for **Trichy SRM Run - 6th Edition (27 November 2026)** allowing participants to search by registered mobile number, switch between multiple participants sharing the same mobile number, view their runner details with a highlighted **T-Shirt Size**, and download their **PDF** or **PNG** ticket featuring a scannable **Code 128** Bib ID barcode.

## Architecture
- **Backend (`app.py`)**: Flask + Gunicorn API loading participant records server-side from `SRM RUN 6.0 FINAL 1.xlsx` (`Male S+E` and `Female S+E` sheets), so full participant datasets are never shipped to the browser.
  - `GET /` — Participant ticket portal UI (`templates/index.html`)
  - `POST /api/search` — Normalized mobile lookup returning only matching participant(s)
  - `GET /api/barcode/<bib>.svg` — Server-side Code 128 SVG barcode for any registered Bib ID
  - `GET /api/health` — Render health check endpoint (`200` when dataset is loaded)
- **Frontend (`templates/index.html`)**: Clean search-first UI with high-DPI HTML5 Canvas ticket rendering, crisp Code 128 barcode generation, PNG download, and pure-JS A4 Landscape PDF export with vector Code 128 barcode overlay.

## Deploy on Render
1. Push this repository to GitHub (recommended: set the repository to **Private** in GitHub Settings if you do not want participant spreadsheet rows publicly browsable).
2. In Render, choose **New + -> Blueprint** (uses `render.yaml`) or **New + -> Web Service** connected to your repository:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 4 --timeout 60`
   - **Health Check Path**: `/api/health`
3. *(Optional Privacy Hardening)* You can remove `SRM RUN 6.0 FINAL 1.xlsx` from Git and upload it as a Render **Secret File** at `/etc/secrets/SRM_RUN_6_0_FINAL_1.xlsx` (or set the `EXCEL_PATH` environment variable). `app.py` checks both automatically.

## Run Locally
```bash
pip install -r requirements.txt
python app.py
```
Then open `http://127.0.0.1:5000`.
