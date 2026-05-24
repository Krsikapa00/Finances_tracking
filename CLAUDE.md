# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

A personal finance tool that processes exported bank/investment CSVs and writes formatted Excel workbooks (`finances.xlsx`, `investments.xlsx`) with categorized transactions, summaries, and portfolio history.

## Commands

```powershell
# Install dependencies
pip install -r requirements.txt

# Generate bank transaction workbook (finances.xlsx)
python generate_excel.py

# Generate investment portfolio workbook (investments.xlsx)
python generate_investments.py
```

`run.bat` is a convenience wrapper for `generate_excel.py` on Windows.

## Architecture

### generate_excel.py → finances.xlsx

Reads all CSVs under `Statements/{AccountName}/` and produces five sheets:

| Sheet | Contents |
|---|---|
| Transactions | Chronological log with category, amount, balance |
| Monthly Summary | Pivot: category × month spending |
| Category Rules | Editable keyword→category mapping |
| Conflicts | Transactions matched by multiple rules |
| Uncategorized | Transactions with no matching rule |

**Two CSV formats are supported:**
- *Chequing*: columns `Date, Description, Debit, Credit, Balance` (date format `YYYY-MM-DD`)
- *Visa/Credit*: columns `Date, Description, Charge, Payment, Balance` (date format `MM/DD/YYYY`)

**Categorization** uses keyword rules stored inside the workbook's "Category Rules" sheet and backed up to `category_rules_backup.json`. On each run, rules are loaded from the JSON backup (if the xlsx is absent) or from the xlsx sheet, then re-saved to JSON. This prevents rule loss if the xlsx is deleted.

### generate_investments.py → investments.xlsx

Reads CSVs under `Investments/{AccountType}_{Currency}_{ID}/` and produces five sheets:

| Sheet | Contents |
|---|---|
| Portfolio Summary | Balance per account |
| Holdings | Current positions: qty, cost basis, market value, unrealized gain |
| Transactions | Full trade history with commissions |
| Dividends | Income events (DIV, WHTX02, NRT, TXPDDV) |
| Account History | Historical snapshots from `holdings_history.csv` |

Account folder naming: `{Type}_{Currency}_{ID}` where Type ∈ {FHSA, TFSA, SDRSP} and Currency ∈ {CAD, USD}.

The script preserves any sheets it doesn't own when regenerating the workbook (only the five managed sheets are replaced).

`holdings_history.csv` is appended on each run to track portfolio growth over time.

## Key Data Flow

```
Statements/**/*.csv  ──┐
category_rules_backup.json ─┤→ generate_excel.py → finances.xlsx
                            └→ (updates) category_rules_backup.json

Investments/**/*.csv ──┐
holdings_history.csv ──┤→ generate_investments.py → investments.xlsx
                       └→ (appends) holdings_history.csv
```

## Adding New Statement Files

- Drop Chequing exports into `Statements/TD_Chequing_1234/`
- Drop Visa exports into `Statements/TD_Rewards_Visa_5678/`
- Drop investment activity/holdings CSVs into the matching `Investments/` subfolder
- Re-run the relevant script; the workbook is fully regenerated each time
