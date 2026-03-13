## Disclaimer

This tool is provided for personal organisational use only. It is not a fatwa or religious ruling. Nisab values, calculation methods, and what qualifies as zakatable wealth can differ between scholars and schools of thought. Always consult a qualified Islamic scholar regarding your specific Zakat obligations.

# Zakat-LogBook 🌙
Happy to help if any one have any questions or need help.


> A comprehensive Microsoft Excel workbook for tracking annual Zakat obligations and all charitable giving.
I customized this over the years and felt it might be a good thing to share incase anyone is intrested.

This workbook allows you to keep track of your zakat based on every year how much you have paid and how much is owe.

![Version](https://img.shields.io/badge/Version-v3.0-green)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Excel](https://img.shields.io/badge/Excel-2016%2B-orange)

---

## Overview

Zakat-LogBook helps you:

- Calculate your annual Zakat obligation based on live gold price, Nisab threshold, and net zakatable assets
- Track all charitable payments (Zakat, Sadaqah, Fitrana, Qurbani, and more) in a searchable running ledger
- Monitor payment status per year — paid in full, partially paid, or not started
- See unpaid Zakat carried forward across years
- Generate detailed per-person and per-year reports on all giving history
- Get a live Nisab dollar threshold by entering today's gold price

---

## Quick Start
Simply download the Zakat-Logbook.xlsx file and start using it. Check the guide page for any questions or reach out to me.

Always start with Adding entry in Zakat Summary workbook. To add more, insert rows before the last data row — formulas will expand automatically.


## Workbook Structure

8 sheets, purpose-built and fully linked to each other:

| Sheet | Purpose |
|---|---|
| **Guide** | Reference sheet — colour legend, usage guide, Nisab explanation, FAQ |
| **Settings** | Configure payment types, transfer services, recipients, and Nisab values |
| **Zakat Summary** | Main annual tracker — one row per Zakat year, all calculations automatic |
| **Stocks** | Stock and investment account balances, one row per year |
| **Cash** | Cash and liquid asset balances, one row per year |
| **Debts** | Outstanding debt balances, one row per year |
| **Ledger** | Every individual payment with running totals — the core log |
| **Reports** | Filter by recipient and year, breakdown by type, service fee summary |

---

## Features

### Zakat Summary — 14 Columns

One row per Zakat year. Enter your date and gold price — everything else calculates automatically.

| Column | What it does |
|---|---|
| A — Zakat Date | Enter your annual Zakat calculation date |
| B — Stock Portfolio | Auto-pulled from Stocks sheet (matched by row position) |
| C — Cash & Liquid | Auto-pulled from Cash sheet |
| D — Total Debts | Auto-pulled from Debts sheet |
| E — Gold Price ($/oz) | Enter the gold spot price on this date |
| F — Gold Owned (oz) | Enter your troy oz of gold |
| G — Value of Gold | = Gold Price × oz (auto) |
| H — Net Zakatable Assets | = Stocks + Cash + Gold − Debts (auto) |
| I — Nisab Threshold | = Gold Price × 2.7315 oz (auto, configurable in Settings) |
| J — Zakat Due (2.5%) | = 2.5% of Net Assets if ≥ Nisab, else $0 (auto) |
| K — Paid This Period | Auto-summed from Ledger entries for this date range |
| L — Running Balance | Cumulative unpaid balance across all years (auto) |
| M — Status | ✅ Paid in Full / ⚠️ Partially Paid / ❌ Not Started (auto) |
| N — Brought Forward | Unpaid balance carried in from the prior year (auto) |

**Additional Summary features:**
- **Duplicate date warning** — highlights red if you enter the same year twice
- **Dashboard** — 6 totals cards: Total Owed, Total Paid, Sadaqah, Fitrana, Qurbani, Outstanding Balance
- **Hawl Tracker** — calculates your next Zakat due date (last date + 354 days), shows live countdown and status emoji

### Stocks / Cash / Debts Sheets

- 10 rows for 10 years of data
- 6 named account columns per sheet with an auto-summing Total column that links automatically to Zakat Summary
- Sort and filter on all columns
- Default account names (all editable in the script):
  - **Stocks:** TD Ameritrade, Charles Schwab, Fidelity, Vanguard, Robinhood, Other Account
  - **Cash:** Chase Checking, Chase Savings, Bank of America, Money Market, Cash on Hand, Other Liquid
  - **Debts:** Chase Credit Card, Citi Credit Card, Amex Credit Card, Car Loan, Personal Loan, Other Debt

### Ledger — The Core Log

200 rows for recording every individual payment you make.

| Column | Details |
|---|---|
| Date | Payment date |
| Type | Dropdown — Zakat, Sadaqah, Fitrana, Qurbani + any you add in Settings |
| Service Used | Dropdown — Remitly, Wise, Bank Transfer, Cash + any you add in Settings |
| Given To | Dropdown — Islamic Relief USA, Zakat Foundation + any you add in Settings |
| Details / Notes | Free text |
| Amount ($) | Payment amount (validated: no negatives) |
| Fees ($) | Transfer fees (validated: no negatives) |
| Total Paid ($) | = Amount + Fees (auto) |
| Running Total ($) | Cumulative sum of all payments (auto) |

**Ledger features:**
- **Search bar** — type any keyword to instantly highlight matching rows in yellow across Type, Given To, and Notes columns
- **4 dropdown validations** — Type, Service Used, and Given To all link to Settings lists; Amount/Fees reject negative values
- **Running total** — updates automatically as you add entries
- Sort and filter on all columns

### Reports Sheet

- **Person filter** — dropdown of all unique recipient names extracted automatically from the Ledger
- **Year filter** — select a specific year or view all years at once; all cards update instantly
- **Breakdown by payment type** — shows total given to the selected person per type (Zakat, Sadaqah, etc.), respecting both filters
- **100-row transaction detail table** — every matching payment for the selected person and year
- **Service fee summary** — all-time fees, amounts, and payment counts per transfer service

### Settings Sheet

Everything configurable in one place. Changes take effect immediately in all dropdowns.

| Column | Contents | Pre-filled examples |
|---|---|---|
| B — Payment Types | 30 slots | Zakat, Sadaqah, Fitrana, Qurbani |
| D — Transfer Services | 30 slots | Remitly, Western Union, Wise, PayPal, Zelle, Bank Transfer, Cash, Check, Venmo, CashApp, MoneyGram, Other |
| F — Recipients / Given To | 30 slots | Islamic Relief USA, Zakat Foundation, LaunchGood, Local Mosque, Family Member |

**Nisab Settings:**
- Gold Nisab oz: **2.7315** (= 85g ÷ 31.1035 g/oz) — stored in cell D44
- Silver Nisab oz: **19.1358** (= 595g ÷ 31.1035 g/oz) — stored in cell D45
- Both values are editable if your scholar uses a different standard

**Live Nisab Calculator:**
- Enter today's gold price → instantly see the current Nisab dollar threshold
- Silver threshold shown alongside for comparison

---

## How Zakat Is Calculated

```
Net Zakatable Assets = Stocks + Cash + (Gold Price × Gold oz) − Debts
Nisab Threshold      = Gold Price × 2.7315 oz  (≈ value of 85 grams of gold)

If Net Assets ≥ Nisab:   Zakat Due = Net Assets × 2.5%
If Net Assets < Nisab:   Zakat Due = $0
```

**Important:** Zakat is calculated on your **entire** net zakatable wealth — not just the amount above Nisab. This follows the majority position of all four major Sunni schools (Hanafi, Maliki, Shafi'i, Hanbali). The Nisab is a qualifying threshold only. Once crossed, 2.5% applies to the full amount.

---

## Customising the Workbook

### Add a new payment type, service, or recipient
Go to **Settings** and type into any empty slot in the relevant column (B, D, or F). The corresponding dropdown in the Ledger updates immediately — no formulas to edit.

### Change the Nisab oz value
**Settings → Nisab Settings → cell D44** (gold) or **D45** (silver). All Nisab calculations in Zakat Summary update instantly.

### Check today's Nisab threshold
**Settings → Current Nisab Calculator → enter gold price in B49.** The dollar threshold shows immediately.

### Add more years
The Zakat Summary, Stocks, Cash, and Debts sheets each have 10 rows (for 10 years). To add more, insert rows before the last data row — formulas will expand automatically.

### Change account names
Open `generate_zakat_logbook.py` and edit the account lists in `main()`:

```python
build_asset_sheet(wb, 'Stocks',
    ["TD Ameritrade", "Charles Schwab", "Fidelity", "Vanguard", "Robinhood", "Other Account"],
    "Total Portfolio")
```

Then re-run the script to get a fresh workbook with your account names.

---

## Regenerating the Workbook

Always Start by downloading the Sheet and **SAVE AS** new file

> **Tip:** Keep your personal working copy saved under a different name (e.g. `My_Zakat_2025.xlsx`) so you never accidentally overwrite your data by re-running the script.

---



## Compatibility
I might be wrong but it should work with these platforms

| Platform | Status |
|---|---|
| Microsoft Excel 2016+ | ✅ Fully supported |
| Microsoft Excel 365 | ✅ Fully supported |
| LibreOffice Calc | ⚠️ Mostly works — some emoji rendering and conditional formatting may differ |
| Google Sheets | ⚠️ Import works — dropdown validations linked to named ranges may need re-linking |
| macOS Excel | ✅ Fully supported |

---

## Versioning

| Version | Status | Notes |
|---|---|---|
| v3.0 | ✅ Current | Finalized release — Recipients/Given To dropdown added, clean regeneration script |

Future changes release as **v3.1, v3.2**, etc. The version is set in `generate_zakat_logbook.py`:

```python
VERSION = "v3.0"   # bump this for future releases
```

---

## Contributing

Contributions are welcome. If you find a bug, have a feature request, or want to add support for a different Nisab calculation method or currency, open an issue or submit a pull request.

Please keep pull requests focused — one feature or fix per PR.

---


---

## License

MIT License — see [LICENSE](LICENSE) for full text.

Free to use, modify, and share for personal or commercial purposes. The only requirement is keeping the copyright notice in place.

