import pandas as pd
from openpyxl import Workbook
import os

# function for adjusting cell widths in the final report
def adjust_column_widths(writer, sheet_name):
    worksheet = writer.sheets[sheet_name]
    # iterate through the columns in the worksheet
    for col in worksheet.iter_cols():
        # set where to start, max length starts at 0 and start at the first column
        max_length = 0
        column = col[0].column_letter  # Get the column letter
        # iterate through the rows in each column and set max length if the it is longer than the previous max
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))  # Added str() to ensure cell value is a string
            except:
                pass
        # Add a bit to the cell length
        adjusted_width = (max_length + 2)
        # Set the width
        worksheet.column_dimensions[column].width = adjusted_width

def compare_excel_files(file1, file2, column_mapping, primary_key_file1, primary_key_file2):
    # Error if file not found
    if not os.path.isfile(file1) or not os.path.isfile(file2):
        raise ValueError(f"file {file1} or {file2} not found")
    
    # Read the Excel files and create the dataframes
    df1 = pd.read_excel(file1, engine ='openpyxl')
    df2 = pd.read_excel(file2, engine ='openpyxl')

    # Ensure the primary key exists in both DataFrames
    if primary_key_file1 not in df1.columns or primary_key_file2 not in df2.columns:
        raise ValueError(f"Primary key {primary_key_file1} or {primary_key_file2} not found in the respective Excel files")

    # Set the primary key as the index for both DataFrames
    df1.set_index(primary_key_file1, inplace=True)
    df2.set_index(primary_key_file2, inplace=True)
    
    # Align the DataFrames based on the index (primary key)
    df1, df2 = df1.align(df2, join='inner', axis=0)

    # Converting Dates to the same format in both files
    # df1['CloseDate'] = pd.to_datetime(df1['CloseDate']).dt.strftime('%Y-%m-%d')
    # df2['CloseDate'] = pd.to_datetime(df2['CloseDate']).dt.strftime('%Y-%m-%d')

    # replace a value with another value
    # df2['Content__c'] = df2['Content__c'].str.replace("l&#39;","l'")
    # df2['Content__c'] = df2['Content__c'].str.replace("d&#39;","d'")
    # df2['Content__c'] = df2['Content__c'].str.replace("&#39;","'")
    # df2['Content__c'] = df2['Content__c'].str.replace("&amp;","&")
    # df2['Content__c'] = df2['Content__c'].str.replace("&lt;","<")
    # df2['Content__c'] = df2['Content__c'].str.replace("&gt;",">")

    # Initialize an empty list to store the results and differences
    results = []
    differences = []
    similarities = []

    # Iterate through each column in the column mapping
    for column1, column2 in column_mapping.items():
        # error if a mapping columns item not found in either files
        if column1 not in df1.columns or column2 not in df2.columns:
            raise ValueError(f"Column {column1} or {column2} not found in the respective excel files")

        # Calculate values
        total_rows = len(df1)
        correct_rows = ((df1[column1] == df2[column2]) | (df1[column1].isna() & df2[column2].isna()) | ((df1[column1] == "") & (df2[column2] == ""))).sum()
        incorrect_rows = total_rows - correct_rows
        percentage_correct = (correct_rows / total_rows) * 100
        percentage_incorrect = (incorrect_rows / total_rows) * 100
        if len(df1) > len(df2):
            fallout = len(df1) - len(df2)
        else:
            fallout = len(df2) - len(df1)

        # Add to the lists
        results.append({
            "Field Comparison": f"{column1} vs {column2}",
            "Total Rows": total_rows,
            "Correct Rows": correct_rows,
            "Incorrect Rows": incorrect_rows,
            "Percentage Correct": percentage_correct,
            "Percentage_Incorrect": percentage_incorrect,
            'Fallout': fallout
        })

        # Find differences and append to differences list all others append to the similarities list
        for idx in df1.index:
            source_value = df1.at[idx, column1]
            target_value = df2.at[idx, column2]
            
            if source_value != target_value and not (pd.isna(source_value) and pd.isna(target_value)):
                differences.append({
                    "Primary Key": idx,
                    "Source Column": column1,
                    "Source Value": source_value,
                    "Target Value": target_value
                })
            else:
                # make a similarities list
                similarities.append({
                    "Primary Key": idx,
                    "Source Column": column1,
                    "Source Value": source_value,
                    "Target Value": target_value
                })
    
    # Create a DataFrame for the results and differences and similarities
    results_df = pd.DataFrame(results)
    differences_df = pd.DataFrame(differences)
    similarities_df = pd.DataFrame(similarities)

    # Filter similarities to include only primary keys not in differences, error thrown if the df is empty
    if not differences_df.empty:
        differences_primary_keys = set(differences_df["Primary Key"])
        similarities_df = pd.DataFrame([sim for sim in similarities if sim["Primary Key"] not in differences_primary_keys])

    # Write the results/differences/similarities to a new Excel file
    with pd.ExcelWriter("comparison_report.xlsx", engine='openpyxl') as writer:
        results_df.to_excel(writer, sheet_name='Summary', index=False)
        adjust_column_widths(writer, 'Summary')
        # Check if differences_df is empty
        if differences_df.empty:
            # Create an empty DataFrame with the required columns
            empty_df = pd.DataFrame(columns=['Primary Key','Source Column', 'Source Value', 'Target Value'])
            empty_df.to_excel(writer, sheet_name='Differences', index=False)
        else:
            # Pivot the differences DataFrame and reset the index
            differences_pivot = differences_df.pivot(index='Primary Key', columns='Source Column', values=['Source Value','Target Value']).reset_index()
            # Correctly flatten the MultiIndex columns
            differences_pivot.columns = ['_'.join(col).strip() if isinstance(col, tuple) else col for col in differences_pivot.columns.values]
            # Write the pivot table to the Excel file
            differences_pivot.to_excel(writer, sheet_name='Differences', index=False)
        adjust_column_widths(writer, 'Differences')

        if similarities_df.empty:
            # Create an empty DataFrame with the required columns
            empty_df = pd.DataFrame(columns=['Primary Key','Source Column', 'Source Value', 'Target Value'])
            empty_df.to_excel(writer, sheet_name='Similarities', index=False)
        else:
            # Pivot the similarities DataFrame and reset the index
            
            similarities_pivot = similarities_df.pivot(index='Primary Key', columns='Source Column', values=['Source Value','Target Value']).reset_index()

            # Correctly flatten the MultiIndex columns
            similarities_pivot.columns = ['_'.join(col).strip() if isinstance(col, tuple) else col for col in similarities_pivot.columns.values]

            # Write the pivot table to the Excel file without adding a new index column
            similarities_pivot.to_excel(writer, sheet_name='Similarities', index=False)
        adjust_column_widths(writer, 'Similarities')

    print("Comparison report generated: comparison_report.xlsx")

# Example usage
column_mapping = {
    "First Name": "First Name",
    "Last Name": "Last Name",
    "Phone": "Phone",
    "Address": "Address",
    "Height": "Height",
    "Company": "Company",
    "ContactID": "ContactID",
    "DateAdded": "DateAdded",
    "Status": "Status",
}

primary_key_file1 = "ContactID"
primary_key_file2 = "ContactID"

compare_excel_files("FILE ONE", "FILE TWO", column_mapping, primary_key_file1, primary_key_file2)

