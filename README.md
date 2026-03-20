# 🏦 Finance & Banking — Excel Data Analytics Portfolio Project

---

## 📌 Project Overview

A complete data analytics project built entirely in Microsoft Excel, simulating real-world work at a retail bank. The dataset contains **400 bank transactions** covering loans, deposits, and credit cards across 5 branches for FY 2024.

This project demonstrates the full data analyst workflow:
**Raw Data → Cleaning → Analysis → Pivot Tables → Charts → Dashboard**

---

## 📂 File Structure

```
Finance_Banking_Portfolio.xlsx
│
├── 0. Start Here        ← Step-by-step learning guide & interview prep
├── 1. Raw Data          ← 400 transactions with intentional data quality issues
├── 2. Clean Data        ← Cleaned & enriched dataset with derived columns
├── 3. Analysis          ← KPI cards + 5 breakdown tables
├── 4. Pivot Summary     ← 3 multi-dimensional cross-tab tables
├── 5. Charts            ← 4 professional charts
└── 6. Dashboard         ← Executive summary dashboard (fully dynamic)
```

---

## 🗂️ Dataset Description

| Field | Description |
|---|---|
| Transaction ID | Unique identifier (TXN-XXXXXX) |
| Date | Transaction date (FY 2024) |
| Branch | One of 5 branches: Downtown, Uptown, Westside, Eastgate, Northpark |
| Product | Checking/Savings Account, Personal Loan, Mortgage, Credit Card, etc. |
| Category | Deposits · Loans · Cards |
| Relationship Manager | One of 5 managers |
| Status | Active · Inactive · Pending |
| Risk Level | Low · Medium · High |
| Amount ($) | Transaction value |
| Interest Rate (%) | Applied interest rate |
| Term (Months) | Loan/deposit term |
| Customer Segment | Basic · Standard · Premium |

---

## 🧹 Data Cleaning (Sheet 2)

The raw dataset contained **5 data quality issues**, all fixed using Excel formulas:

| Issue | Fix Applied | Formula Used |
|---|---|---|
| Branch names in ALL CAPS | Standardised to Title Case | `=PROPER(TRIM(A2))` |
| Leading spaces in Status column | Removed whitespace | `=TRIM(A2)` |
| ~10 missing Amount values | Replaced with column average | `=IFERROR(IF(A2="", AVERAGEIF(...), VALUE(A2)), 0)` |
| ~8 missing Interest Rate values | Same approach | `=IFERROR(IF(A2="", AVERAGEIF(...), VALUE(A2)), 0)` |
| No date parts (Month/Quarter/Year) | Extracted from date | `=TEXT(B2,"mmm")` · `=INT((MONTH(B2)-1)/3+1)` |

**Bonus derived column:** Monthly loan payment calculated using:
```
=(Amount × Rate/12) / (1 − (1 + Rate/12)^−Term)
```

---

## 📊 Analysis (Sheet 3)

**KPI Cards:**
- Total Active Portfolio Value
- Total Number of Transactions
- Average Loan Amount
- Average Interest Rate

**Breakdown Tables:**

| Table | Dimensions | Formulas Used |
|---|---|---|
| A. By Branch | 5 branches × Amount/Count/Avg/Share | `SUMIF`, `COUNTIF` |
| B. By Category | Deposits/Loans/Cards × Amount/Rate/Term | `SUMIF`, `AVERAGEIF` |
| C. Monthly Trend | Jan–Dec × Amount/Count/Growth | `SUMIF`, MoM formula |
| D. By Risk Level | Low/Medium/High × Amount/Count | `SUMIF`, `COUNTIF` |
| E. Manager Performance | 5 managers × Portfolio/Deals/Avg | `SUMIF`, `RANK` |

Monthly growth cells are **conditionally formatted** — green for positive, red for negative.

---

## 🔄 Pivot Summary (Sheet 4)

Three cross-tab tables built with `SUMPRODUCT` (multi-condition aggregation):

1. **Branch × Product Category** — portfolio value matrix
2. **Risk Level × Customer Segment** — deal count matrix
3. **Quarter × Branch** — quarterly performance by location

> 💡 `SUMPRODUCT` allows multi-condition analysis without PivotTables — a highly valued skill in finance analyst interviews.

---

## 📈 Charts (Sheet 5)

| Chart | Type | Shows |
|---|---|---|
| Portfolio by Branch | Column | Which branch holds the most value |
| Category Split | Pie | Deposits vs Loans vs Cards share |
| Monthly Trend | Line | Portfolio growth across 2024 |
| Risk Level Comparison | Horizontal Bar | Exposure by risk category |

---

## 🎯 Dashboard (Sheet 6)

An executive-level summary dashboard featuring:
- **4 KPI headline cards** (Active Portfolio, Total Deals, Avg Loan, High Risk Count)
- **Branch, Category & Risk mini-tables**
- **Top performing manager** (dynamically calculated)
- **Fully dynamic** — all values update automatically when source data changes

---

## 🔧 Excel Skills Demonstrated

### Formulas
| Formula | Used For |
|---|---|
| `SUMIF()` | Aggregate amounts by branch, category, risk |
| `COUNTIF()` | Count transactions by dimension |
| `AVERAGEIF()` | Average rates/amounts by category |
| `SUMPRODUCT()` | Multi-condition pivot-style analysis |
| `IFERROR()` | Handle missing/error values gracefully |
| `PROPER()` + `TRIM()` | Data cleaning — fix casing & spaces |
| `TEXT()` | Extract month names from dates |
| `RANK()` | Rank manager performance |
| `INDEX()` + `MATCH()` | Dynamic lookups (dashboard top manager) |
| PMT-style formula | Calculate monthly loan repayments |

### Other Skills
- ✅ Conditional Formatting (colour scales, data bars, cell rules)
- ✅ Excel Tables with structured references
- ✅ Auto-filters and Freeze Panes
- ✅ 4 Chart types (Column, Pie, Line, Bar)
- ✅ Multi-sheet workbook design
- ✅ Dynamic dashboard with no hardcoded values

---

## 💼 Interview Talking Points

> **"How did you handle missing data?"**
> I used `IFERROR` combined with `AVERAGEIF` to replace ~18 missing values with the column average — a simple but defensible imputation strategy for this context.

> **"What's the most complex formula you used?"**
> `SUMPRODUCT` for multi-dimensional analysis. For example: summing all transaction amounts where Branch = "Downtown" AND Category = "Loans" simultaneously.

> **"How is the dashboard built?"**
> All KPI cards and tables use formulas that reference the Clean Data sheet directly. There are no hardcoded values in the dashboard — change anything in the data and everything updates automatically.

> **"What finance-specific analysis did you do?"**
> I calculated monthly loan repayment amounts using a PMT-style formula, tracked month-over-month portfolio growth, and segmented exposure by risk level — metrics a real bank analyst would report on.

---

## 🚀 How to Use

1. Download `Finance_Banking_Portfolio.xlsx`
2. Open in Microsoft Excel (2016 or later recommended)
3. Start on the **"0. Start Here"** tab for a step-by-step walkthrough
4. Follow the tabs in order: Raw → Clean → Analysis → Pivot → Charts → Dashboard

---

## 🔮 Possible Extensions

- [ ] Add **Power Query** to automate the cleaning steps
- [ ] Add **PivotTable + Slicers** for interactive filtering
- [ ] Build an `IF + AND` risk flagging column: `=IF(AND(Amount>100000, Risk="High"), "Flag", "OK")`
- [ ] Add a **VLOOKUP** product rate table
- [ ] Add **Sparklines** inside the monthly trend table
- [ ] Model 3 scenarios with **Scenario Manager** (base / optimistic / pessimistic)

---

## 👤 About

Admir Istrefi
Data Analysis Enthusiast
📧 admir.istrefi@gmail.com
📍 Prishtina, Kosovo

---

*If you found this useful, feel free to ⭐ star the repository!*
