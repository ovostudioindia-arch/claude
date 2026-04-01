# 50/30/20 Personal Budget Tracker — Google Sheets

## Quick Setup (2 minutes)

1. **Create** a new Google Sheet at [sheets.new](https://sheets.new)
2. **Open** Extensions → Apps Script
3. **Delete** any default code, then paste the entire contents of `budget-tracker.gs`
4. **Run** the `buildBudgetTracker` function (click ▶ or Run → buildBudgetTracker)
5. **Authorize** when prompted — the script only modifies this spreadsheet
6. Wait ~15 seconds for the build to complete

## What Gets Created

| Tab | Purpose |
|---|---|
| **Dashboard** | Main view with month selector, budget vs. actual tables, 3 charts |
| **Transactions** | Data entry — date, amount, sub-category, description |
| **Budget Setup** | Set income, 50/30/20 ratios, and fixed monthly bills |
| **Reference** | Hidden lookup tables (categories, buckets, months) |

## How to Use

- **Budget Setup tab**: Edit the yellow cells to set your income and ratios
- **Budget Setup tab**: Add/edit fixed monthly bills (rows 9+)
- **Transactions tab**: Enter daily spending in columns A–C and E (D and F auto-fill)
- **Dashboard tab**: Use the month dropdown (cell I4) to filter all data and charts

## 50/30/20 Rule

| Bucket | Target | Covers |
|---|---|---|
| **Need** | 50% | Rent, utilities, groceries, insurance, gas, health, phone, internet |
| **Want** | 30% | Dining out, clothes, entertainment, streaming, gym, transport, other |
| **Savings / Debts** | 20% | Loan payments, credit cards, 401k, savings transfers |

## Seed Data

The script pre-loads 40 transactions across January–March 2024 so charts and tables populate immediately.
