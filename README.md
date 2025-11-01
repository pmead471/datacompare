# Data Compare Tool

A Python-based data comparison utility for validating data migrations and identifying discrepancies between two Excel files. This tool generates comprehensive comparison reports showing field-level accuracy metrics and detailed record-by-record differences.

## Overview

This tool was developed to automate data quality validation during data migrations. It compares two Excel files based on configurable field mappings and a primary key, then generates an Excel report with three tabs:
- **Summary**: Field-level comparison statistics including match percentages
- **Differences**: Detailed view of all records with discrepancies
- **Similarities**: Records that match between both files

## Features

- **Flexible Column Mapping**: Compare fields with different names between source and target files
- **Primary Key Validation**: Ensures data integrity by matching records on a unique identifier
- **Error Handling**: Validates file existence, primary keys, and mapped columns
- **Statistical Analysis**: Calculates match percentages and identifies data quality issues
- **Auto-formatted Reports**: Generated Excel reports with auto-adjusted column widths for readability
- **Handles Edge Cases**: Properly treats null values, empty strings, and missing records

## Usage

### Basic Configuration

1. Update the **column mapping** dictionary to define which fields to compare:
```python
column_mapping = {
    "First Name": "First Name",      # Source Column: Target Column
    "Last Name": "Last Name",
    "Phone": "Phone",
    "Address": "Address",
    "Company": "Company",
    "DateAdded": "DateAdded",
    "Status": "Status",
}
```

2. Set the **primary keys** for both files:
```python
primary_key_file1 = "ContactID"  # Primary key in FILE ONE.xlsx
primary_key_file2 = "ContactID"  # Primary key in FILE TWO.xlsx
```

3. Update the **file paths** in the function call:
```python
compare_excel_files("FILE ONE.xlsx", "FILE TWO.xlsx", column_mapping, primary_key_file1, primary_key_file2)
```

## Use Cases

This tool is ideal for:
- **Data Migration Validation**: Verify data integrity after migrating between systems
- **ETL Quality Assurance**: Validate transformations in data pipelines
- **Database Reconciliation**: Compare production vs. backup data
- **API Testing**: Validate data consistency between source and target APIs
- **Regression Testing**: Ensure data remains unchanged after system updates

## Sample Files

- `FILE ONE.xlsx` - Source data file (example)
- `FILE TWO.xlsx` - Target data file (example)
- `comparison_report.xlsx` - Generated output report (example)

## Author

**Patrick Mead**  
Data Quality Analyst | 7+ years experience in data validation and quality assurance

