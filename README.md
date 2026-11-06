# Excel Portfolio: Data Analysis & Formulas

*(Note: Download the `.xlsx` file from this repository to view the formulas, datasets, and calculations in action).*

## Foundations & Aggregations
* **Basic Functions:** Explored foundational formulas including `SUM`, `AVERAGE`, and basic arithmetic for revenue calculations.
* **Data Structuring:** Applied Excel tables, formatted datasets, and calculated total & average metrics by region and sales rep.

## Task Previews
*(<img width="1917" height="1018" alt="Screenshot 2026-10-09 220018" src="https://github.com/user-attachments/assets/2fbc0ddd-08cd-4dff-b0d4-ffac6b940f7a" />
)*

## Logical Functions & Error Handling
* **Error Management:** Implemented `IFERROR` and `IFNA` to clean up broken formulas and maintain data integrity.
* **Logical Checks:** Used `IF` statements to categorize data and handle edge cases dynamically.

## Task Previews
*(<img width="1917" height="1020" alt="Screenshot 2026-10-09 215957" src="https://github.com/user-attachments/assets/c4a19f8a-546a-413a-ab34-edc7919692eb" />
)*

## Lookup Methods Compared (XLOOKUP vs INDEX+MATCH)
Executed 9 specific lookup tasks to evaluate the strengths and weaknesses of different lookup functions.

* **VLOOKUP Vulnerability:** Demonstrated how it breaks entirely if new columns are inserted or deleted. Strictly requires the lookup column to be the leftmost column.
* **INDEX+MATCH Resilience:** Utilized for backward compatibility and structural resilience. Successfully implemented **Two-Way Lookups** to find intersection points between row and column headers.
* **XLOOKUP Superiority:** Applied the modern standard for:
  * Left-side lookups.
  * Multi-criteria array lookups using boolean logic (`*`).
  * Approximate matches for commission tiers (`match_mode = -1`).
  * Reverse lookups to find the latest entries (`search_mode = -1`).
  * Spill arrays to return multiple columns of data from a single formula.

## Task Previews
*(<img width="1917" height="1017" alt="Screenshot 2026-10-12 182625" src="https://github.com/user-attachments/assets/47da8f3e-c490-4466-aabc-f4dd28f76c46" />)*

# Data Cleaning & Power Query (ETL Pipeline)

## 📌 Overview
This module demonstrates end-to-end data sanitization using Excel's native cleaning tools and automated ETL (Extract, Transform, Load) pipelines using **Power Query**.

---

## 🛠️ Part 1: Native Data Cleaning & Validation

* **Data Validation & Auditing:** Restricted `Region` inputs via list validation and `Units` to positive decimals (`> 0`). Audited legacy anomalies (`"hi"`, `-5`) using **Circle Invalid Data**.
* **Visual Profiling (Data Bars):** Applied Conditional Formatting Data Bars on `Revenue` to visually identify high-performing records and outliers.
* **Pattern Extraction (Flash Fill):** Extracted 3-digit area codes from formatted phone strings (`(044) 234-5678`) using Flash Fill (`Ctrl + E`), enforcing text pre-formatting (`'044`) to preserve leading zeros.
* **Deduplication:** Flagged duplicate employee records using structured `COUNTIFS([Emp ID], [@Emp ID])` logic prior to executing `Remove Duplicates`.
* **Type Coercion:** Converted text-formatted numerics (`"80"`) into true numeric values.

## Task Previews
*(<img width="842" height="195" alt="image" src="https://github.com/user-attachments/assets/f6d8f6ef-6aa0-45b3-be5d-8d32beea0380" />)*

---

## ⚙️ Part 2: Power Query (M-Engine) Transformations

### 1. Merge Queries (Left Outer Join )
*(<img width="596" height="73" alt="image" src="https://github.com/user-attachments/assets/25b66612-ba50-4a83-9e7d-9c7c66299e85" />)*

---

### 2. Append Queries, Custom Column & Refresh (Tasks 7, 9, 10)
*(<img width="486" height="145" alt="image" src="https://github.com/user-attachments/assets/43dfb3c4-fe9a-4ad9-a5b9-2ac6d6f9ce1c" />)*

---

### 3. Unpivot Columns - Wide to Tall Normalization (Task 8)
*(<img width="242" height="169" alt="image" src="https://github.com/user-attachments/assets/3e63aa12-b209-4b25-b082-209d6ed8a273" />)*

# Pivot Tables & Executive Charting (Interactive Analytics)

## 📌 Overview
This module demonstrates rapid multidimensional data aggregation, custom metric creation, and executive dashboard design using **Excel Pivot Tables**, **Slicers/Timelines**, and **Combo Charts**.

---

## 🛠️ Part 1: Pivot Table Engine & Custom Metrics

- **Data Type Auditing & Aggregation:** Diagnosed and resolved text-formatted numeric anomalies (`'600`) in the `Raw_Sales` source table that forced `Revenue` aggregations to default to `COUNT` (`9`) instead of `SUM` (`7,900`).
- **Time-Series Matrix (Region × Month):** Built cross-tabulated revenue breakdowns utilizing automatic date hierarchies (`Jan`–`Apr`).
- **Relative Contribution & Cumulative Tracking:** Configured `Show Values As` rules to compute rep-level revenue share (`% of Grand Total`) and monthly cumulative performance (`Running Total In`).
- **Calculated Fields (Profitability):** Engineered an in-memory `Margin` metric (`=(Revenue - Cost) / Revenue`) directly within the Pivot cache to evaluate true weighted profitability (`43%`–`44%`) without bloating the source table.
- **Stable Model Lookups (`GETPIVOTDATA`):** Extracted dynamic summary KPIs (`1,700` for `Arun`) using metadata-driven `GETPIVOTDATA` references rather than fragile cell coordinates.

---

## 📊 Part 2: Interactive Visuals & Dynamic Refresh

- **Multi-Pivot Slicers & Timelines:** Deployed interactive `Region` Slicers and `Date` Timelines linked across all worksheet Pivot Tables via **Report Connections** for synchronized filtering.
- **Dual-Axis Combo Charting:** Visualized monthly performance by pairing `Sum of Revenue` (Clustered Columns) with `Average of OrderValue` / AOV (Line on a **Secondary Axis**) to compare volume against deal size.
- **Dynamic Source Expansion (`Refresh All`):** Appended new transaction records (`OrderID 110`, `Revenue 2000`) to the structured `Raw_Sales` Excel Table (`Ctrl + T`) and executed `Refresh All` (`Ctrl + Alt + F5`), automatically scaling Grand Total revenue from `7,900` to `9,900` across all linked tables and charts.

---

## 📸 Task Previews

### 1. Structured Source Data (`Raw_Sales` Tab)
*(<img width="944" height="265" alt="image" src="https://github.com/user-attachments/assets/c5369838-8c6d-4daf-8b16-0067404db071" />)*

---

### 2. Interactive Analytics & Combo Chart (`Pivots` Tab)
*(<img width="1917" height="991" alt="Screenshot 2026-11-06 235024" src="https://github.com/user-attachments/assets/ae7a66fc-5fcb-4726-82ca-1780aaec3c89" />)*
