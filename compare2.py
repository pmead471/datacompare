import pandas as pd
from openpyxl import Workbook
import os

# function for adjusting cell widths in the final report
def adjust_column_widths(writer, sheet_name):
    worksheet = writer.sheets[sheet_name]
    for col in worksheet.iter_cols():
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column].width = adjusted_width

def compare_excel_files(file1, file2, column_mapping, primary_key_file1, primary_key_file2):
    if not os.path.isfile(file1):
        raise ValueError(f"file {file1} not found")
    if not os.path.isfile(file2):
        raise ValueError(f"file {file2} not found")
    
    df1 = pd.read_excel(file1, engine='openpyxl')
    df2 = pd.read_excel(file2, engine='openpyxl')

    if primary_key_file1 not in df1.columns or primary_key_file2 not in df2.columns:
        raise ValueError(f"Primary key {primary_key_file1} or {primary_key_file2} not found in the respective Excel files")

    df1.set_index(primary_key_file1, inplace=True)
    df2.set_index(primary_key_file2, inplace=True)
    
    df1, df2 = df1.align(df2, join='inner', axis=0)

    # Ensure CloseDate is in the same format for both DataFrames
    df1['CloseDate'] = pd.to_datetime(df1['CloseDate']).dt.strftime('%Y-%m-%d')
    df2['CloseDate'] = pd.to_datetime(df2['CloseDate']).dt.strftime('%Y-%m-%d')

    results = []
    differences = []
    similarities = []

    for column1, column2 in column_mapping.items():
        if column1 not in df1.columns or column2 not in df2.columns:
            raise ValueError(f"Column {column1} or {column2} not found in the respective excel files")

        total_rows = len(df1)
        correct_rows = ((df1[column1] == df2[column2]) | (df1[column1].isna() & df2[column2].isna()) | ((df1[column1] == "") & (df2[column2] == ""))).sum()
        incorrect_rows = total_rows - correct_rows
        percentage_correct = (correct_rows / total_rows) * 100
        percentage_incorrect = (incorrect_rows / total_rows) * 100

        results.append({
            "Field Comparison": f"{column1} vs {column2}",
            "Total Rows": total_rows,
            "Correct Rows": correct_rows,
            "Incorrect Rows": incorrect_rows,
            "Percentage Correct": percentage_correct,
            "Percentage_Incorrect": percentage_incorrect
        })

    # Find differences and similarities based on entire row comparison
    for idx in df1.index:
        source_row = df1.loc[idx]
        target_row = df2.loc[idx]
        
        if source_row.equals(target_row):
            similarities.append({
                "Primary Key": idx,
                **source_row.to_dict()
            })
        else:
            differences.append({
                "Primary Key": idx,
                **source_row.to_dict(),
                **{f"Target_{col}": target_row[col] for col in target_row.index}
            })

    results_df = pd.DataFrame(results)
    differences_df = pd.DataFrame(differences)
    similarities_df = pd.DataFrame(similarities)

    with pd.ExcelWriter("comparison_report.xlsx", engine='openpyxl') as writer:
        results_df.to_excel(writer, sheet_name='Summary', index=False)
        adjust_column_widths(writer, 'Summary')
        
        if differences_df.empty:
            empty_df = pd.DataFrame(columns=['Primary Key'] + list(df1.columns) + [f"Target_{col}" for col in df2.columns])
            empty_df.to_excel(writer, sheet_name='Differences', index=False)
            adjust_column_widths(writer, 'Differences')
        else:
            differences_df.to_excel(writer, sheet_name='Differences', index=False)
            adjust_column_widths(writer, 'Differences')

        if similarities_df.empty:
            empty_df = pd.DataFrame(columns=['Primary Key'] + list(df1.columns))
            empty_df.to_excel(writer, sheet_name='Similarities', index=False)
            adjust_column_widths(writer, 'Similarities')
        else:
            similarities_df.to_excel(writer, sheet_name='Similarities', index=False)
            adjust_column_widths(writer, 'Similarities')

    print("Comparison report generated: comparison_report.xlsx")

# Example usage
column_mapping = {
    "Name": "Name",
    "Amount": "Amount",
    "ForecastCategory": "ForecastCategory",
    "LeadSource": "LeadSource",
    "Type": "Type",
    "Constant.TRUE": "ATG_Migrated__c",
    "CloseDate": "CloseDate",
    "StageName": "StageName"
}

primary_key_file1 = "ATG_Opportunity_External_Id__c"
primary_key_file2 = "ATG_External_ID__c"

compare_excel_files("Opportunity Source Iteration 1 10_11.xlsx", "Opportunity Target Iteration 1 10_11.xlsx", column_mapping, primary_key_file1, primary_key_file2)