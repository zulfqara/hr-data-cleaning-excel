# HR Data Cleaning & Quality Assurance (DQA) Project

## Project Overview
This project focuses on cleaning and validating an HR dataset using Microsoft Excel. The main goal was to perform comprehensive data cleaning and standardization across the dataset by identifying formatting inconsistencies, mapping incorrect entries to valid ones, and validating the final output using standard Excel functions. 

## What Was Actually Done (Step-by-Step)

### 1. File Setup & Import
*   Imported the raw CSV data into Excel.
*   Created a working file where we copied the raw data into a new tab named "raw snapshot".

### 2. Initial Data Quality Assurance (DQA)
*   Performed an Initial DQA on the raw Excel file to profile the dataset.
*   Used `SUMPRODUCT`, `COUNTIFS`, `COUNTBLANK`, and `COUNTA` functions to figure out the problems and inconsistencies in the data (such as 190 blanks in the Email column and 376 pattern failures in the Phone Number column).
*   Used `MIN` and `MAX` functions to know the corners and numerical boundaries of the data. 

### 3. Data Cleaning & Mapping Architecture
After finding out the anomalies, we started cleaning the data using a dedicated mapping strategy and specific formulas:
*   **The Mapping Tab:** We created a separate tab where we first extracted the unique invalid values (the "Discarded" columns) and then later decided on their corrected versions.
*   **Dynamic Lookups:** To apply these corrections to the main dataset, we used `XLOOKUP` to look up the correct value from the mapping tab. This approach made the cleaning process efficient and allowed us to avoid writing complex, nested formulas.

### 4. Text and Number Standardization Formulas
Alongside the lookups, we applied the following core formulas to standardize the dataset:
*   **`IFERROR`**: Used to handle errors seamlessly (e.g., if a lookup didn't find a match).
*   **`VALUE`**: Used to convert numbers that were incorrectly stored as text back into real, usable numbers.
*   **`TRIM`**: Used to remove invisible leading and trailing spaces from the text.
*   **`LOWER`**: Used to convert the email addresses into lowercase.
*   **`PROPER`**: Used to format the names with correct capitalization.

### 5. Final Validation (Post DQA)
*   After the cleaning steps were complete, we performed a final Post DQA to validate the work.
*   We re-ran the exact same Initial DQA formulas on the cleaned data. 
*   This time, the numbers showed zero for those anomalies (e.g., 0 blanks, 0 pattern failures, and 0 non-numeric errors), confirming the data was fully standardized.
