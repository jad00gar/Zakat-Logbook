## Disclaimer

This tool is provided for personal organisational use only. It is not a fatwa or religious ruling. Nisab values, calculation methods, and what qualifies as zakatable wealth can vary between scholars and schools of thought. Always consult a qualified Islamic scholar regarding your specific Zakat obligations.

# Zakat Tracker 🌙

A fully-featured Microsoft Excel workbook for tracking annual Zakat obligations, charitable giving, and payment history — generated entirely from a Python script using `openpyxl`.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue) ![openpyxl](https://img.shields.io/badge/openpyxl-3.0%2B-green) ![License](https://img.shields.io/badge/License-MIT-yellow)

---

## What It Does

Zakat Tracker helps you:
- Calculate your annual Zakat obligation based on gold price, Nisab threshold, and net zakatable assets
- Track all charitable payments (Zakat, Sadaqah, Fitrana, Qurbani, and more) in a running ledger
- See how much has been paid, what's still outstanding, and when the next Zakat year is due
- Generate per-person and per-year reports on your giving history

---

## Quick Start

**Requirements:** Python 3.8+ and one library.

```bash
pip install openpyxl
python generate_zakat_tracker.py
```

This produces `Zakat_Tracker.xlsx` in your current directory. Open it in Microsoft Excel (recommended) or LibreOffice Calc.

---

## Workbook Structure

The workbook contains 8 sheets:

| Sheet | Purpose |
|-------|---------|
| **Guide** | Reference sheet — colour coding legend, usage guide, Nisab explanation, FAQ |
| **Settings** | Configure payment types, transfer services, recipients, and Nisab values |
| **Zakat Summary** | Main annual tracker — one row per Zakat year, all calculations automatic |
| **Stocks** | Stock and investment account balances per year |
| **Cash** | Cash and liquid asset balances per year |
| **Debts** | Outstanding debt balances per year |
| **Ledger** | Every individual payment with running totals |
| **Reports** | Filter by recipient and year, breakdown by payment type |

---

## Features

### Zakat Summary Sheet
- **14 columns** tracking every component of your annual Zakat calculation:
  - Stocks, Cash, Gold value, Debts auto-pulled from their sheets by matching date
  - Net Zakatable Assets = Stocks + Cash + Gold − Debts
  - Nisab Threshold = Gold price × 2.7315 oz (= 85 grams) — configurable in Settings
  - Zakat Due = 2.5% of full net wealth if assets ≥ Nisab (majority scholarly position)
  - Paid This Period pulled automatically from Ledger entries
  - Running Balance — cumulative outstanding Zakat across years
  - **Status indicator** — ✅ Paid in Full / ⚠️ Partially Paid / ❌ Not Started
  - **Brought Forward** — unpaid Zakat carried in from the prior year
- **Duplicate date warning** — highlights red if you accidentally enter the same Zakat year twice
- **Summary Dashboard** — 6 totals cards: Total Owed, Total Paid, Sadaqah, Fitrana, Qurbani, Outstanding Balance
- **Hawl Tracker** — calculates your next Zakat due date (last date + 354 days), shows countdown and live status emoji (🕌 Due Now / ⚠️ Due Soon / ✅ In Progress)
- Sort and filter on all columns

### Stocks / Cash / Debts Sheets
- 10 rows for 10 years of data
- 6 configurable account columns per sheet + auto-summing Total column
- Values auto-link to Zakat Summary when dates match
- Sort and filter on all columns

### Ledger Sheet
- **200 rows** for individual payment records
- Columns: Date, Type, Service Used, Given To, Details/Notes, Amount, Fees, Total Paid, Running Total
- **Dropdown validation** on Type (from Settings), Service (from Settings), and Given To / Recipient (from Settings)
- **No-negative validation** on Amount and Fees columns
- **Search bar** — type any keyword to highlight matching rows in yellow across Type, Given To, and Details columns
- Running cumulative total column updates automatically as you enter payments
- Sort and filter on all columns

### Reports Sheet
- **Person filter** — dropdown of all unique names from the Ledger
- **Year filter** — filter by specific year or view all years at once
- **Breakdown by payment type** — SUMPRODUCT totals for every type in Settings, respecting both filters
- **100-row detail table** — shows every matching transaction for the selected person/year
- **Service fee summary table** — all-time fees and payment counts per transfer service

### Settings Sheet
- **Payment Types** — 30 slots (pre-filled: Zakat, Sadaqah, Fitrana, Qurbani). Add your own.
- **Transfer Services** — 30 slots (pre-filled: Remitly, Wise, Bank Transfer, Cash, Zelle). Add your own.
- **Recipients / Given To** — 30 slots (pre-filled: Islamic Relief USA, Zakat Foundation, LaunchGood, Local Mosque, Family Member). Add your own.
- **Nisab Settings** — Gold Nisab oz (default 2.7315 = 85g), Silver Nisab oz (19.1358 = 595g). Change these if your scholar uses a different standard.
- **Live Nisab Calculator** — enter today's gold price to instantly see the current Nisab dollar threshold

---

## How Zakat Is Calculated

```
Net Zakatable Assets (H) = Stocks + Cash + Gold Value − Debts
Gold Value               = Gold Price ($/oz) × Gold Owned (oz)
Nisab Threshold (I)      = Gold Price ($/oz) × 2.7315 oz (= 85g)

If H ≥ I:  Zakat Due = H × 2.5%
If H < I:  Zakat Due = $0
```

**Important:** Zakat is calculated on your **entire** net zakatable wealth, not just the amount above Nisab. This follows the majority scholarly position of all four major Sunni schools (Hanafi, Maliki, Shafi'i, Hanbali). The Nisab is a qualifying threshold — once crossed, 2.5% applies to the full amount.

---

## Customising the Workbook

### Adding new payment types, services, or recipients
Go to the **Settings** sheet and type into any empty slot in the relevant column. The dropdowns in the Ledger update automatically.

### Changing the Nisab oz value
Go to **Settings → Nisab Settings** and edit cell D44 (Gold) or D45 (Silver). All Nisab calculations in Zakat Summary update immediately.

### Checking today's Nisab
Go to **Settings → Current Nisab Calculator** and enter the current gold spot price in cell B49.

### Adding more account columns
The Stocks, Cash, and Debts sheets have 6 account columns each. To add more, insert a column before the Total column — the SUM formula will expand automatically.

---

## Regenerating the Workbook

If you want a fresh copy (e.g. to reset all data while keeping the structure):

```bash
python generate_zakat_tracker.py
```

To customise the default values (account names, preset types/services/recipients, Nisab defaults), edit the constants at the top of each `build_*` function in `generate_zakat_tracker.py`.

---

## Requirements

```
openpyxl>=3.0.0
```

Install with:
```bash
pip install openpyxl
# or
pip install -r requirements.txt
```

---

## Compatibility

- **Microsoft Excel 2016+** — fully supported, all features work
- **Microsoft Excel 365** — fully supported
- **LibreOffice Calc** — mostly works; some conditional formatting and emoji rendering may differ
- **Google Sheets** — import works but dropdowns linked to named ranges may need manual re-linking

---



---

## License

MIT License — see [LICENSE](LICENSE) for details.

You are free to use, modify, and share this for personal or commercial purposes. The only requirement is that you keep the copyright notice in place.

---

## Contributing

Contributions welcome. If you find a bug, have a feature request, or want to add support for a different Nisab calculation method, open an issue or submit a pull request.

