# AI Claim Amount Calculator

Streamlit app that classifies claim payment-status rows, totals pending and paid amounts, and exports a reviewable report. Built as a colourful dashboard for pasted text, Excel/CSV uploads, or screenshot input via the OpenAI Responses API.

## Purpose

Turn messy claim status lists (lines like `pending on $2,866.02`, `paid $500.00`, `pending no paid amnt`) into:

- Counts for **Pending On Amount**, **Paid Amount**, **Pending No Paid Amount**, and **Pending Only** rows
- Counts for blank rows, black-highlighted rows, and **Other / Ignored** rows
- Totals for pending amount, paid amount, and net open amount (pending − paid)
- An editable results table so you can correct classifications before export
- CSV and Excel downloads of the edited results (`ai_claim_amount_report.csv`, `ai_claim_amount_report.xlsx`)

## Features

- **Paste text** — line-by-line classification with `$` currency parsing; empty lines become blank rows
- **Upload Excel (`.xlsx`) or CSV** — choose the status/text column; optionally map a separate pending amount column and/or paid amount column. Only the active Excel sheet is read. `.xls` is not accepted
- **Excel row alerts** — blank rows (every cell empty) and black-highlighted rows. A row is black-highlighted when at least one cell has a black or very dark fill (openpyxl `fgColor`, perceived brightness under 45). CSV has no fill colour, so black-row detection is Excel-only; CSV still flags blank rows
- **Screenshot mode** — accepts PNG, JPG, JPEG, or WEBP, sends the image to the OpenAI Responses API (`responses.create` with `input_image`), and parses the returned JSON rows (requires `OPENAI_API_KEY`)
- **Interactive dashboard** — Streamlit `data_editor` for category, amounts, row alert, and include flags; formula breakdown expander
- **Export** — download CSV or Excel of the edited results

Text and Excel/CSV paths work **without** an API key. Only the screenshot tab needs OpenAI.

## Tech stack

| Piece | Choice |
|---|---|
| UI | Streamlit |
| Data | pandas |
| Excel | openpyxl |
| Config | python-dotenv (`load_dotenv()`) |
| AI (screenshots) | OpenAI Python SDK (`responses.create` with image input) |

## Project structure

```text
ai-claim-amount-calculator/
├── app.py              # Full Streamlit application (UI, parsing, dashboard, AI)
├── requirements.txt    # streamlit, pandas, openpyxl, python-dotenv, openai
├── .env.example        # Placeholder OPENAI_API_KEY and OPENAI_MODEL
├── .gitignore          # ignores .env, .venv/, __pycache__/, *.pyc
└── README.md
```

There is no package layout or tests.

## Environment variables

Copy the example file and edit it locally (`.env` is gitignored):

```bash
cp .env.example .env
```

```env
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=gpt-4o-mini
```

| Variable | Required | Default | Used for |
|---|---|---|---|
| `OPENAI_API_KEY` | Only for screenshot tab | — | OpenAI client |
| `OPENAI_MODEL` | No | `gpt-4o-mini` | Vision / Responses model name |

## How to run

```bash
git clone https://github.com/Santo250499/ai-claim-amount-calculator.git
cd ai-claim-amount-calculator
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env        # optional; set OPENAI_API_KEY for screenshot mode
streamlit run app.py
```

## How classification works (high level)

Rule-based rows (paste, Excel, CSV) are labelled **Pending On Amount**, **Paid Amount**, **Pending No Paid Amount**, **Pending Only**, **Blank Row**, or **Other / Ignored**.

- Mapped amount columns override the text parse when the chosen value is greater than 0. If both a pending column and a paid column are mapped and both are greater than 0, the pending amount wins.
- Text containing `no paid` (including `pending no paid amnt`) is **Pending No Paid Amount** with both amounts set to 0, before any `$` amount is read.
- Otherwise a `$` amount together with `paid` is **Paid Amount**. A `$` amount together with `pending`, and no `paid`, is **Pending On Amount**.
- Text that contains `pending` but no `$` amount is **Pending Only**. Anything else is **Other / Ignored**.

Black highlighting is a **row alert**, not a category. The text category stays as classified, and the status notes that a black row was found. A fully empty Excel row is **Blank Row** even if it also has a dark fill. Blank rows are both category **Blank Row** and row alert **Blank Row**. The screenshot prompt may also return a **Black Row** category; the results table lets you change category and row alert before export.

Pending and paid totals only include rows whose `include_pending` / `include_paid` flags are true (editable in the table). Net open amount is pending total minus paid total.

## Suggested follow-ups for this repo

- Add a short sample CSV under `samples/` so someone can try paste/upload without real claim data
- Consider a thin test module for `extract_currency_amount` / `classify_claim_row` (pure functions already exist in `app.py`)
