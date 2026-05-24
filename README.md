# AR Aging Report Generator

A Streamlit app hosted on Railway that converts a raw invoices export into a formatted **AR Aging Summary** Excel workbook — matching the Y&S Group reporting layout.

## Features

- Upload any invoices `.xlsx` export (full or pre-filtered format)
- Select an **As of Date** — aging buckets recalculate automatically
- Generates a two-tab Excel workbook:
  - **AR Aging Summary** — pivot by network × aging bucket, whole-dollar currency with live SUMIFS formulas
  - **Invoice Details** — full filtered invoice list with frozen header row
- Only networks with outstanding balances appear on the summary tab
- Company name normalization (YSA 2/3 → YSA, YS Tickets Spec → YS Tickets, etc.)
- Each report run is logged to Supabase

## Aging Buckets

| Bucket | Days Outstanding |
|---|---|
| Current | 0 or fewer |
| 1 to 30 | 1–30 days |
| 31 to 60 | 31–60 days |
| 61 to 90 | 61–90 days |
| 91 and Over | 91+ days |

## Local Setup

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_ORG/ar-aging-report.git
cd ar-aging-report

# 2. Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Set environment variables
export SUPABASE_URL=https://your-project.supabase.co/rest/v1/
export SUPABASE_KEY=your-anon-key

# 5. Run the app
streamlit run app.py
```

The app will open at `http://localhost:8501`.

## Deploying to Railway

1. Push this repo to GitHub.
2. Go to [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**.
3. Select your repo — Railway auto-detects Python and installs `requirements.txt`.
4. Under **Settings → Deploy**, set the start command:
   ```
   streamlit run app.py --server.port=$PORT --server.address=0.0.0.0
   ```
5. Under **Variables**, add:
   - `SUPABASE_URL` → `https://your-project.supabase.co/rest/v1/`
   - `SUPABASE_KEY` → your Supabase anon key
6. Click **Deploy** — Railway redeploys automatically on every GitHub push.

## Input File Requirements

Two export formats are supported. The app auto-detects which one was uploaded.

**Full format** (23 columns) — includes `Paid`, `IsCancelled`; app filters to unpaid/active invoices.

**Light format** (8 columns) — pre-filtered export; all rows treated as unpaid and finalized.

Both formats must include:

| Column | Description |
|---|---|
| `Bal.` | Outstanding balance (numeric) |
| `Client` | Network / marketplace name |
| `Company` | Broker entity name |
| `Inv#` | Invoice number |
| `Ext Order #` | External order reference |
| `Status` | Invoice status |
| `Created` | Invoice creation datetime |

## Project Structure

```
ar-aging-report/
├── app.py              # Streamlit UI
├── report_builder.py   # Excel generation logic
├── logger.py           # Supabase usage logging
├── requirements.txt    # Python dependencies
├── Procfile            # Railway start command
├── .gitignore
└── README.md
```
