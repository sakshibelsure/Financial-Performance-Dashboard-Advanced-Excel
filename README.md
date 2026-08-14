# 📊 Executive Financial Performance Dashboard (Advanced Excel)

## 📌 Project Overview
An end-to-end financial analysis and MIS reporting dashboard built in Microsoft Excel. This project processes transactional financial data to track organizational KPIs, variance against budget targets, regional profitability, and product line margins across FY 2025–2026.

---

## 📸 Dashboard Preview

![Executive Dashboard](Dashboard.png)

---

## 🛠️ Data Architecture & Analytical Workflow
1. **Raw Data Ingestion:** 1,000+ line items of multi-region operational financial records.
2. **Data Cleaning & Standardization:** Handled text normalization (`TRIM`, `PROPER`), dynamic fiscal period mapping (`FY`, fiscal quarters), and structured Excel tables (`tbl_FinancialData`).
3. **Calculation & Audit Controls:**
   - Real-time audit checks (`PASS`/`FAIL`) for revenue balancing, profit integrity, and budget reconciliation.
   - Dynamic conditional aggregation using `SUMIFS` and structured table references.
4. **Pivot Reporting Layer:** Pivot tables segmented by Department, Region, Product Category, and Timeline.
5. **Executive MIS Dashboard:** Interactive KPI cards, dynamic slicers, variance indicators, and strategic management callouts.

---

## 📈 Key Business Metrics Analyzed
- **Revenue & Profitability:** ₹2.81 Cr Total Revenue, ₹40.49 L Net Profit (13.34% Net Margin).
- **Budget Performance:** ₹1.05 Cr Budget Variance (96.39% Target Achievement).
- **Expense Breakdown:** COGS (60.14%), OPEX (9.09%), Salary (7.98%), Marketing (5.70%).

---

## 💡 Key Excel Formulas & Functions Used
- **Fiscal Calendar Calculations:** `="Q"&INT(MOD(MONTH(Date)-4,12)/3)+1`
- **Dynamic Conditional Aggregation:** `SUMIFS()`, `AVERAGEIFS()`, `COUNTA()`
- **Structured References:** `tbl_FinancialData[Revenue]`, `tbl_FinancialData[Net Profit]`
- **Audit Controls:** `IF()`, `ISNUMBER()`, Nested Logical Validation Checks
- **Interactive Visuals:** Dynamic KPI Cards, Slicers, Custom Formatted Variance Charts
