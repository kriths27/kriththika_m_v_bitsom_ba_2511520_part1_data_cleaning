# Part 1: Business Data Cleaning, Validation & Excel Reporting

**Student:** Kriththika M V | **ID:** `bitsom_ba_2511520`  
**Repository:** `kriththikamv_bitsom_ba_2511520_part1_data_cleaning`  
**Tool Stack:** Python 3 (pandas, xlsxwriter, matplotlib) · Excel · Markdown

---

## Problem Statement

The raw transactional dataset (`raw_orders.xlsx`) for an Indian e-commerce business contained **932 order records** across multiple regions, categories, and customer segments. The data suffered from:

- Mixed and inconsistent date formats (8+ variations)
- Missing values in critical fields (`region`, `ship_mode`, `discount`)
- Duplicate `order_id` entries (data entry errors)
- Invalid discount values (negative or > 100%)
- Logical date errors (ship date earlier than order date)
- Mixed order statuses (Completed, Cancelled, Returned, Failed) requiring segmentation

The goal was to **clean, validate, and standardize** the dataset, generate calculated columns for business analysis, build pivot-style summaries, and produce a full data quality audit report.

---

## Repository Structure

```
kriththikamv_bitsom_ba_2511520_part1_data_cleaning/
├── data/
│   ├── raw_orders.xlsx           ← Original dataset (DO NOT MODIFY)
│   └── cleaned_orders.xlsx       ← Cleaned dataset with 29 columns
├── outputs/
│   ├── data_quality_report.xlsx  ← 5-sheet audit report
│   ├── pivot_summary.xlsx        ← 4-sheet pivot analysis
│   └── cleaning_log.md           ← Issue log, rules, assumptions
├── screenshots/
│   ├── raw_data_preview.png      ← Raw data snapshot with issue summary
│   ├── cleaned_data_preview.png  ← Cleaned data with calculated columns
│   ├── pivot_summary_1.png       ← Region × Category analysis
│   └── pivot_summary_2.png       ← Monthly trend + Segment + Shipping
└── README.md
```

---

## Dataset Description

| Attribute | Value |
|---|---|
| Source file | `raw_orders.xlsx` |
| Total rows | 932 |
| Columns | 21 (order_id, dates, customer, geography, product, financials, status) |
| Date range | 2024–2025 |
| Missing values | region (26), ship_mode (22), discount (18) |
| Duplicates | 32 rows with repeated order_id |

---

## Tools & Technologies

| Tool | Purpose |
|---|---|
| **Python 3** | Core processing language |
| **pandas** | Data loading, transformation, pivoting |
| **xlsxwriter** | Excel file generation with rich formatting |
| **matplotlib** | Screenshot/visual generation |
| **openpyxl** | Excel reading |

---

## Cleaning Steps Applied

### Step 1: Load & Inspect
- Loaded raw data with `dtype=str` to avoid premature type conversion
- Assessed shape, column types, null counts, duplicate counts

### Step 2: Date Parsing (Multi-Format)
- Applied 8-format date parser + fallback `pd.to_datetime(dayfirst=True)`
- Handled formats: `21 Jul 2024`, `06/08/2024`, `28-11-2024`, `2025-09-01`, etc.
- **Zero unparseable dates** — all 932 dates successfully parsed

### Step 3: Numeric Conversion
- Cast `quantity`, `unit_price`, `discount`, `sales`, `cost`, `profit` to float
- Non-numeric discount values coerced to NaN before imputation

### Step 4: Duplicate Detection
- Identified **32 duplicate order_ids** (keeping first occurrence as authoritative)
- Marked with `DUPLICATE` flag; excluded from cleaned dataset and analyses

### Step 5: Data Quality Flagging
- Assigned pipe-separated `data_quality_flag` per row: `CLEAN | MISSING_REGION | MISSING_SHIP_MODE | MISSING_DISCOUNT | INVALID_DISCOUNT | SHIP_BEFORE_ORDER | EXTREME_MARGIN`

### Step 6: Business Rules
| Rule | Implementation |
|---|---|
| Missing `region` → `"Unknown"` | `fillna('Unknown')` |
| Missing `ship_mode` → `"Unknown"` | `fillna('Unknown')` |
| Missing `discount` → `0` | `fillna(0)` where other key fields valid |
| Invalid discount → flagged + clamped | `clip(0, 1)` → `cleaned_discount` |
| Cancelled/Failed orders → excluded from sales | Filtered by `order_status` |
| Returned orders → summarized separately | Separate refund sheet |
| Ship date before order date → flagged | Date comparison logic |

### Step 7: Calculated Columns

| Column | Formula |
|---|---|
| `cleaned_discount` | `CLAMP(discount, 0, 1)`, null → 0 |
| `calculated_sales` | `quantity × unit_price × (1 − cleaned_discount)` |
| `calculated_profit` | `sales − cost` |
| `profit_margin` | `calculated_profit / calculated_sales` |
| `shipping_delay_days` | `ship_date − order_date` (calendar days) |
| `order_month` | `YYYY-MM` period |
| `order_year` | Year extracted from order_date |
| `data_quality_flag` | Pipe-delimited audit flags |

---

## Output Files

### `cleaned_orders.xlsx`
- 29 columns including all original fields + 7 calculated columns
- Color-coded rows: 🟩 Green = CLEAN, 🟨 Yellow = minor issue, 🟥 Red = critical issue
- Excludes duplicate records

### `data_quality_report.xlsx` (5 sheets)
1. **Summary Overview** — Aggregate issue counts and percentages
2. **Row-Level Flags** — Every row with its specific flags
3. **Duplicates** — All 32 duplicate order records isolated
4. **Date Issues** — Records with date parsing or logic errors
5. **Refunds Summary** — All 164 Returned orders with financial impact

### `pivot_summary.xlsx` (4 sheets)
1. **Region × Category** — Sales, profit, order count, avg margin, avg discount by region and product category
2. **Monthly Trend** — Month-by-month sales and profit totals with margin trend
3. **Segment Breakdown** — Consumer, Corporate, Small Business, Home Office analysis
4. **Ship Mode Performance** — Average shipping delay and sales by shipping method

---

## Key Business Insights

| Insight | Finding |
|---|---|
| **Data Quality Rate** | 85.5% of records are flagged CLEAN (797/932) |
| **Duplicate Rate** | 3.4% — indicates data entry pipeline issues |
| **Active vs. Cancelled** | 622 active vs. 146 cancelled/failed (23.5% loss rate) |
| **Refund Impact** | 164 returned orders; financial impact tracked separately |
| **Discount Validity** | 16 records had out-of-range discounts — potential pricing system bug |
| **Shipping Anomaly** | 22 records show ship date before order date — data entry error |
| **Missing Region** | 26 records (2.8%) — likely from orders placed via a new channel |

---

## Limitations

1. Duplicate resolution uses "keep first" — may not always select the most recent/accurate record
2. Returned orders are excluded from active sales but their original revenue entry is not reversed in the raw data
3. Discount clamping preserves data volume at cost of analysis accuracy for 16 records
4. Missing region imputed as "Unknown" — regional analysis will understate true regional totals
5. Calculated profit (`sales − cost`) may differ slightly from the raw `profit` column due to source rounding

---

## Screenshots

| File | Description |
|---|---|
| `raw_data_preview.png` | First 8 rows of raw data with issue summary bar |
| `cleaned_data_preview.png` | First 10 rows of cleaned data with calculated columns |
| `pivot_summary_1.png` | Sales by region/category, margin heatmap, profit by category |
| `pivot_summary_2.png` | Monthly trend line, segment bar chart, shipping delay chart |

---

*Part of BITSOM BA Capstone Project | Generated: June 2026*
