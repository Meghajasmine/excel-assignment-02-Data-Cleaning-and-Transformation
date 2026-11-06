# Excel Assignment 2: Data Cleaning and Transformation

## Project Overview
This project focuses on data cleaning, preparation, and transformation using Microsoft Excel. Raw, real-world datasets often present defects such as missing records, inconsistent text formatting, typographical errors, duplicate entries, and concatenated fields. 

Using the **Product Dataset**, this project executes an end-to-end preprocessing workflow to audit, standardize, restructure, and format data into an analysis-ready structure.

---

## Repository & Project Structure

```text
Excel_Assignment_2_Submission/
│
├── Assignment_2_Cleaned_Solution.xlsx     # Cleaned solution workbook (Cleaned Dataset, Raw Dataset, Audit Log)
├── Raw_Dataset.xlsx                       # Unaltered source dataset
├── Project_Summary_One_Page.pdf           # One-page executive project summary
├── Excel_Assignment_2_Report.pdf          # Step-by-step verification report with visual evidence
├── README.md                              # Project documentation
└── screenshots/                           # Verification screenshots of cleaning steps
    ├── 01_missing_value_handling.png
    ├── 02_find_replace_corrections.png
    ├── 03_duplicate_removal_process.png
    ├── 04_column_splitting_and_merging.png
    └── 05_conditional_formatting_results.png
```

---

## Detailed Data Cleaning & Transformation Workflow

### 1. Handling Missing Values
* **Missing Category Attributes (4 records):**
  * Handled via contextual catalog imputation by mapping each missing value to matching product lines:
    * `Sneakers` $\rightarrow$ **Fashion** (consistent with Nike Sneakers)
    * `Coffee Maker` $\rightarrow$ **Kitchen** (consistent with Keurig Coffee Maker)
    * `Fitness Tracker` $\rightarrow$ **Electronics** (consistent with Garmin Fitness Tracker)
    * `Backpack` $\rightarrow$ **Outdoor** (consistent with North Face gear)
* **Missing Price Information (3 records):**
  * Imputed using category benchmarks and product equivalents:
    * `Sony Headphones` $\rightarrow$ **$250.00** (benchmark from Bose Headphones)
    * `Ray-Ban Sunglasses` $\rightarrow$ **$150.00** (benchmark from Oakley Sunglasses)
    * `Coleman Camping Tent` $\rightarrow$ **$130.00** (Outdoor category benchmark)

### 2. Correcting Inconsistent Data
* **Product Name Casing:** Identified lowercase inconsistencies (`laptop`, `smartphone`, `headphones`) and standardized them into Title Case (`Laptop`, `Smartphone`, `Headphones`) using Excel's `=PROPER()` function.
* **Category Typographical Errors:** Identified repeated misspelling `Electroni` across the dataset and executed **Find & Replace** (`Ctrl + H`) to standardize all instances to `Electronics`.

### 3. Removing Duplicates
* Applied Excel's **Data $\rightarrow$ Remove Duplicates** tool evaluating uniqueness across all columns.
* Identified and removed **3 exact duplicate records**:
  * `HP Laptop` (`17-JUN-IN`)
  * `Bose Headphones` (`16-APR-ES`)
  * `Samsonite Laptop Bag` (`21-AUG-CA`)
* Successfully reduced the row count from **34 records to 31 unique rows**.

### 4. Splitting and Merging Data
* **Splitting `Product ID`:** Extracted composite codes (`DD-MMM-CC`) using hyphen delimiter `-` into two distinct dimensional attributes:
  * **Manufacturing Date:** Structured into standardized date entries.
  * **Country Code:** Extracted 2-letter ISO country codes (`US`, `UK`, `CA`, etc.).
* **Merging Columns:** Combined `Brand Name` and `Product Name` into a single **`Product Brand`** column using concatenation formula:
  ```excel
  =[@[Brand Name]] & " " & [@[Product Name]]
  ```

### 5. Number and Date Formatting
* **Price:** Formatted as standard **Currency** (`$#,##0.00`).
* **Manufacturing Date:** Formatted explicitly into the standard **`DD-MM-YYYY`** display convention.

### 6. Conditional Formatting
* **Category Custom Rule:** Highlighted cells where the Category equals `"Electronics"` with a soft green fill (`#E2EFDA`) and dark green text.
* **Price Gradient Data Bars:** Applied gradient Data Bars across the `Price` column to visually represent relative price levels across the catalog.

---

## Key Metrics Summary

| Attribute / Stage | Before Cleaning | After Preprocessing | Primary Operation Applied |
| :--- | :--- | :--- | :--- |
| **Total Rows** | 34 Records | 31 Unique Records | Full-row deduplication |
| **Missing Prices** | 3 nulls (8.8%) | 0 nulls (100% complete) | Equivalent product & benchmark imputation |
| **Missing Categories** | 4 nulls (11.8%) | 0 nulls (100% complete) | Contextual product lookup mapping |
| **Category Typos** | 5 rows (`Electroni`) | 0 errors (`Electronics`) | Find & Replace text standardization |
| **Product Casing** | Inconsistent lowercase | 100% Proper Title Case | Text transformation via `=PROPER()` |
| **Product Identifier** | Composite string | 2 atomic fields | Delimiter split into Date & Country Code |
| **Brand Representation** | 2 separate columns | 1 unified column | Concatenation (`&`) into `Product Brand` |
| **Formatting** | General / Unformatted | Standardized types | Currency (`$`), Date (`DD-MM-YYYY`), Data Bars |
