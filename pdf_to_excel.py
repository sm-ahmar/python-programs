import tabula
import pandas as pd
import openpyxl

#def pdf_to_excel(pdf_file, excel_file):

# list_of_dfs = tabula.read_pdf('lighting_area_schedule.pdf', pages='all',multiple_tables=True)
# print(tabula.environment_info())
# print(list_of_dfs)
pdf_path='input.pdf'
excel_path = 'output.xlsx'
header_text = 'CAPACITY'

try:
        # Use a high stream/guess setting for better detection of tables
        list_of_dfs = tabula.read_pdf(pdf_path, pages='5-1389',
                      multiple_tables=True, pandas_options={'header': None})
except Exception as e:
        print(f"Error during PDF reading: {e}")
if not list_of_dfs:
        print("No tables found or extraction failed.")


print(f"Found and extracted {len(list_of_dfs)} potential tables across all pages.")

# Combine all DataFrames into a single one for a single Excel sheet output
combined_df = pd.concat(list_of_dfs, ignore_index=True)
column_name = combined_df.columns[0]
combined_df = combined_df[combined_df[column_name] != column_name]
indices_to_drop = combined_df[combined_df[1] == header_text].index

if len(indices_to_drop) > 1:
    combined_df.drop(indices_to_drop[1:], inplace=True)

# Output to an Excel file using pandas' to_excel method
try:
    # Using openpyxl as the engine for .xlsx files
    combined_df.to_excel(excel_path, index=False, engine='openpyxl')
    print(f"Successfully saved all extracted tables to {excel_path}")
except Exception as e:
    print(f"Error during Excel writing: {e}")
