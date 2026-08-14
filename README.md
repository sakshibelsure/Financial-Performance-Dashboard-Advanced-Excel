# 📊 Executive Financial Performance Dashboard (Advanced Excel & MIS)

## 📌 Project Overview
An end-to-end financial analysis and MIS reporting dashboard built in Microsoft Excel. This project processes transactional financial data to track organizational KPIs, variance against budget targets, regional profitability, and product line margins across FY 2025–2026.

---

## 📸 Dashboard Preview
![Dashboard Preview](dashboard_preview.png)

---

## 🛠️ Architecture & Workflow
1. **Raw Data Ingestion:** 1,000+ line items of multi-region transactional financial records.
2. **Data Cleaning & Standardization:** Handled text normalization (`TRIM`, `PROPER`), dynamic fiscal period mapping (`FY`, fiscal quarters), and structured Excel tables (`tbl_FinancialData`).
3. **Calculation & Audit Controls:**
   - Real-time audit checks (`PASS`/`FAIL`) for revenue balancing, profit integrity, and budget reconciliation.
   - Dynamic conditional aggregation using `SUMIFS`, structured table formulas, and dynamic array indexing.
4. **Pivot Reporting Layer:** Pivot tables segmented by Department, Region, Product Category, and Timeline.
5. **Executive MIS Dashboard:** Interactive KPI cards, dynamic slicers, variance indicators, and strategic management callouts.

---

## 📈 Key Business Metrics Analyzed
- **Total Revenue:** ₹2.81 Cr
- **Net Profit:** ₹40.49 L (Net Margin: 13.34%)
- **Budget Variance:** ₹1.05 Cr (96.39% Target Achievement)
- **Expense Breakdown:** COGS (60.14%), OPEX (9.09%), Salary (7.98%), Marketing (5.70%)
- **Top Regional Drivers:** West & North regions contributing ~49% of total revenue.

---

## 💡 Key Excel Formulas & Functions Used
- **Fiscal Calendar Calculations:** `="Q"&INT(MOD(MONTH(Date)-4,12)/3)+1`
- **Dynamic Conditional Aggregation:** `SUMIFS()`, `AVERAGEIFS()`, `COUNTA()`
- **Structured Table References:** `tbl_FinancialData[Revenue]`, `tbl_FinancialData[Net Profit]`
- **Audit Controls:** `IF()`, `ISNUMBER()`, `MATCH()`, Nested Logical Validation Checks
