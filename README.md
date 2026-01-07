# Ventas (Bsale) — Sales ETL & Excel Reporting (Python)

Python pipeline to process **Bsale sales exports**, enrich them with product information (e.g., **SKU + cost**), compute key sales metrics (sales, cost, margin), and generate an **Excel report** with business-ready summaries.


## What you get

- Cleaned and standardized sales tables from Bsale exports
- Product enrichment by **SKU** (costs, categories, product type, etc.)
- Key metrics:
  - Total sales
  - Total cost
  - Gross margin
  - Margin %
  - Units sold
- Excel report outputs such as:
  - Summary by product type/category
  - Summary by marketplace/channel (if present in the data)
  - Low-margin products list (e.g., margin < 20%)
  - Top / bottom products by units sold


