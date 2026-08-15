# AI-Powered EDA Report Generator

A Streamlit app that takes a CSV or Excel file, runs a full statistical
exploratory data analysis (EDA) locally with pandas, and then uses the
**Groq API** to turn those statistics into a plain-English narrative
report — overview, key observations, data quality issues, relationships,
and recommended next steps.

> **Why Groq instead of Gemini?** This project originally used the Gemini
> API. During setup we hit three separate Google-side account issues in a
> row (a new `AQ.`-format key rejected by the API, a retired model, and a
> project-level `403 "denied access"` block) — all confirmed as active,
> unresolved problems on Google's end, not code bugs. Groq's free tier
> doesn't have that problem and is also extremely fast, so the AI layer
> was switched over.

## How it works

```
 Upload CSV/XLSX
        │
        ▼
 pandas: shape, dtypes, missing values, describe(), correlations,
 outliers (IQR rule), top categorical values      ──▶ shown in Streamlit tabs
        │
        ▼
 Statistics only (NOT raw rows) serialized to compact text
        │
        ▼
 Groq API (groq SDK, chat completions) → Markdown narrative
        │
        ▼
 Displayed in-app + downloadable as a Markdown report
```

**Privacy/cost note:** only the already-computed summary statistics are
sent to Groq — never the raw dataset rows. This keeps the request small
and cheap regardless of how many rows you upload, and avoids sending
potentially sensitive row-level data to a third-party API.

## Project structure

```
eda_report_generator/
├── app.py             # Streamlit UI and orchestration
├── eda_utils.py        # pandas/matplotlib EDA logic (no API calls)
├── groq_utils.py        # Groq API integration
├── requirements.txt
├── .env.example
└── README.md
```

## Setup

1. **Install dependencies** (Python 3.10+ recommended):

   ```bash
   python -m venv venv
   source venv/bin/activate        # Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

2. **Get a Groq API key**: create one for free at
   [console.groq.com/keys](https://console.groq.com/keys) — no credit card
   required.

3. **Set the key in code, not in the browser.** Copy `.env.example` to
   `.env` in the project folder and paste your key into it:

   ```bash
   cp .env.example .env
   # then edit .env:
   # GROQ_API_KEY=your_actual_key_here
   ```

   `app.py` loads this automatically at startup via `python-dotenv`. There
   is no API-key field in the UI — the key lives only in your local `.env`
   file (which should never be committed to version control) or in a
   `GROQ_API_KEY` environment variable set on the host/server.

4. **Run the app**:

   ```bash
   streamlit run app.py
   ```

   Streamlit will open the app at `http://localhost:8501`. If
   `GROQ_API_KEY` isn't set, the app still loads and the statistical EDA
   tabs still work — you'll just see a banner explaining the AI narrative
   section needs the key.

## Using the app

1. Upload a `.csv`, `.xlsx`, or `.xls` file from the sidebar.
2. Browse the auto-generated tabs: Overview, Missing Values, Statistics,
   Correlations, Distributions, Categorical — all computed instantly with
   pandas, no API key required for this part.
3. Click **Generate AI Report** to have Groq write the narrative
   sections based on those statistics.
4. Click **Download Markdown report** to save the full report (narrative
   + underlying stats) to disk.

## Model

The app is hardcoded to `llama-3.3-70b-versatile` — a well-established,
fast, free-tier model on Groq that's a strong fit for report-writing.
There's no in-app selector by design. If you want a different trade-off
later, change the `DEFAULT_MODEL` constant at the top of `groq_utils.py`:

| Model | Trade-off |
|---|---|
| `llama-3.3-70b-versatile` (current default) | Best balance of quality and speed |
| `llama-3.1-8b-instant` | Faster still, slightly less nuanced narrative |
| `deepseek-r1-distill-llama-70b` | Stronger reasoning, lower free-tier daily limit |

Groq's free tier (per their published limits) runs up to ~14,400
requests/day on most Llama models — far more than this app needs for
normal use. See [console.groq.com/docs/models](https://console.groq.com/docs/models)
for the current full list if a model here has been retired.

## Extending this project

- **PDF export**: convert the Markdown report to PDF with a library like
  `markdown2` + `weasyprint`, or export the Streamlit page as-is.
- **Bigger files**: for very large files, sample the dataframe (e.g.
  `df.sample(n=50_000)`) before computing statistics to keep things fast.
- **Column-level Q&A**: add a chat box that lets users ask Groq
  follow-up questions about specific columns, using the same
  stats-only-context approach.
- **Caching**: wrap the statistic functions in `@st.cache_data` if you
  expect users to re-run the same file often.

## Troubleshooting

- **"GROQ_API_KEY is not set" banner** — you haven't created `.env` yet,
  or it's not in the same folder as `app.py`. Copy `.env.example` to
  `.env` and paste in your key, then restart `streamlit run app.py`.
- **"Groq request failed: ..."** — usually an invalid API key, or the
  `DEFAULT_MODEL` in `groq_utils.py` has been deprecated on Groq's side;
  check [console.groq.com/docs/models](https://console.groq.com/docs/models)
  for a current model name.
- **Rate limit errors** — Groq's free tier has per-minute and per-day
  limits per model; wait a minute and retry, or switch to a different
  model in `groq_utils.py`.
- **Excel files fail to load** — make sure `openpyxl` installed correctly
  (`pip install openpyxl`); it's required for `.xlsx` support.
- **App is slow on large files** — the histogram/bar-chart generation is
  capped at 12 numeric and 8 categorical columns by default; you can
  lower `MAX_NUMERIC_PLOTS` / `MAX_CATEGORICAL_PLOTS` in `eda_utils.py`
  for very wide datasets.
