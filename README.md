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


