# Ventas (Bsale) — Sales ETL & Automated Excel Reporting

A Python project that automates **sales data processing and reporting** using exports from **Bsale**.  
It cleans and standardizes raw transaction files, optionally merges product data (e.g., **SKU + cost**), computes key metrics (sales, cost, margin), and produces an **Excel report** with business-ready summaries.

---

## What this project does

Many sales reports end up as manual Excel work spread across multiple files. This project aims to make the workflow:

- **Reproducible** (same inputs → same report)
- **Automated** (minimal manual steps)
- **Scalable** (works as sales volume grows)

---

## Key features

- Load and normalize Bsale exports (sales documents such as invoices/receipts/credit notes — depending on your setup)
- Merge transactions with a product master table using **SKU**
- Compute business KPIs:
  - Total sales
  - Total cost
  - Gross margin
  - Margin %
  - Units sold
- Generate an Excel report including:
  - Summary by **product type/category**
  - Summary by **marketplace/channel** (if available)
  - **Low-margin alerts** (e.g., margin < 20%)
  - Top / bottom products by units sold

---

## Outputs

Typical output is an Excel file containing:
- Cleaned/standardized tables
- Summary tables (pivot-like)
- Low-margin lists
- Product rankings

> Tip: Add 2–3 screenshots of the final Excel report in `docs/img/` and embed them here to improve portfolio impact.

---

## Repository structure

Suggested layout (adapt t
