import os
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def read_files(folder_path):
    dataframes = {'da': None, 'catalogue': None, 'sl': None}
    if not os.path.exists(folder_path):
        print(f"Folder '{folder_path}' does not exist")
        return dataframes
    files = [f for f in os.listdir(folder_path) if f.endswith(('.csv', '.xlsx', '.xls'))]
    if not files:
        print(f"No relevant files found in '{folder_path}'")
        return dataframes
    for file in files:
        file_path = os.path.join(folder_path, file)
        if file.endswith('.xlsx') or file.endswith('.xls'):
            df = pd.read_excel(file_path)
        else:
            df = pd.read_csv(file_path)

        if 'da' in file.lower():
            dataframes['da'] = df
            print(f"Identified as DA file: {file}")
        elif 'catalogue' in file.lower():
            dataframes['catalogue'] = df
            print(f"Identified as Catalogue file: {file}")
        elif 'sl' in file.lower():
            dataframes['sl'] = df
            print(f"Identified as Shopping List file: {file}")

    if dataframes['da'] is None:
        print("DA file needs to be present in the folder")
    if dataframes['catalogue'] is None:
        print("Catalogue file needs to be present in the folder")
    if dataframes['sl'] is None:
        print("SL file needs to be present in the folder")
    return dataframes


def clean_string(s):
    s = str(s)
    s = re.sub(r'^[\d.]+', '', s)
    s = re.sub(r'\.(?!:)', '', s)
    # s = re.sub(r'(_[A-Za-z]+_\d+)$', '', s)
    return s


def compare_strings(s1,s2):
    if pd.isna(s1) or pd.isna(s2):
        return False
    return s1 in s2 or s2 in s1


def auto_adjust_columns(worksheet, df):
    # Set column widths to auto adjust
    for idx, col in enumerate(df.columns):
        max_length = max(df[col].astype(str).map(len).max(), len(col))
        worksheet.set_column(idx, idx, max_length)


def apply_conditional_formatting(worksheet, df , workbook):
    format_red = workbook.add_format({'bg_color': 'red', 'font_color': 'white'})
    format_yellow = workbook.add_format({'bg_color': 'yellow'})
    # Apply formatting the "Match" columns
    for row_num, row in df.iterrows():
        for col_num, value in enumerate(row):
            if value in ['nok', 'NULL']:
                worksheet.write(row_num + 1, col_num, value, format_red)
            elif 'DA Price is higher' in str(value):
                worksheet.write(row_num + 1, col_num, value, format_yellow)


def process_files(da_file_path, selected_folder):
    folder_path = os.path.join('input_files', selected_folder)
    folder_name = os.path.basename(folder_path)
    print(f"Reading files from folder: {selected_folder}")
    dataframes = read_files(selected_folder)
    if any(df.empty for df in dataframes.values()):
        print("One or more files are empty. Cannot proceed with comparison")
        return

    supplier_name = folder_name.split('_')[0]
    # Prepare the DataFrames
    da_data = dataframes['da'][['Item', 'Description', 'Price', 'Quantity', '[  ]']]
    catalogue_data = dataframes['catalogue'][['Item', 'Description', 'Price']]
    sl_data = dataframes['sl'][['Item', 'Description', 'Price']]
    # Replace the ? in the Description column
    da_data['Description'] = da_data['Description'].str.replace('¿', '’')
    da_data['[  ]'] = da_data['[  ]'].str.replace('¿', '’')
    catalogue_data['Description'] = catalogue_data['Description'].str.replace('¿', '’')
    # clean the [] column
    da_data['[  ]'] = da_data['[  ]'].apply(clean_string)
    # Rename columns for consistency
    da_data.rename(columns={'Description': 'Description_DA', 'Price': 'Price_DA'}, inplace=True)
    # Perform the Comparison with catalogue
    comparison_results_catalogue = pd.merge(da_data, catalogue_data, on='Item', how='inner')

    comparison_results_catalogue['Supplier'] = supplier_name
    comparison_results_catalogue['Description_Match'] = comparison_results_catalogue.apply(
        lambda row: 'ok' if not pd.isna(row['Description']) and row['Description_DA'] == row['Description']
        else ('NULL' if pd.isna(row['Description']) else 'nok'), axis=1)
    comparison_results_catalogue['[  ]_Match'] = comparison_results_catalogue.apply(
        lambda row: 'ok' if compare_strings(row['[  ]'], row['Description']) else 'nok', axis=1)
    comparison_results_catalogue['Price_Match'] = comparison_results_catalogue.apply(
        lambda row: 'ok' if not pd.isna(row['Price']) and row['Price_DA'] == row['Price']
        else ('NULL' if pd.isna(row['Price']) else 'nok'), axis=1)
    # Warning Column
    comparison_results_catalogue['Warning'] = comparison_results_catalogue.apply(
        lambda row: 'DA Price is higher than Catalogue price' if row['Price_DA'] > row['Price'] else '', axis=1)
    # Handle Missing Values
    # comparison_results_catalogue['Quantity'].fillna(0, inplace=True) (method not working in pandas 3.0)
    comparison_results_catalogue.fillna({'Quantity': 0}, inplace=True)
    # Perform the Comparison with sl
    comparison_results_sl = pd.merge(da_data, sl_data, left_on='Description_DA', right_on='Description', how='left')
    comparison_results_sl.rename(columns={'Item_x': 'Item_DA', 'Item_y': 'Item_SL'}, inplace=True)
    comparison_results_sl['Description_Match'] = comparison_results_sl.apply(
        lambda row: 'ok' if not pd.isna(row['Description']) and row['Description_DA'] == row['Description']
        else ('NULL' if pd.isna(row['Description']) else 'nok'), axis=1)
    comparison_results_sl['[  ]_Match'] = comparison_results_sl.apply(
        lambda row: 'ok' if compare_strings(row['[  ]'], row['Description']) else 'nok', axis=1)
    comparison_results_sl['Price_Match'] = comparison_results_sl.apply(
        lambda row: 'ok' if not pd.isna(row['Price']) and row['Price_DA'] == row['Price']
        else ('NULL' if pd.isna(row['Price']) else 'nok'), axis=1)
    # Warning Column
    comparison_results_sl['Warning'] = comparison_results_sl.apply(
        lambda row: 'DA Price is higher than SL price' if row['Price_DA'] > row['Price'] else '', axis=1)
    # Output the Results
    with pd.ExcelWriter('comparison_results.xlsx', engine='xlsxwriter') as writer:
        comparison_results_catalogue.to_excel(writer, sheet_name='DA_catalogue_Comparison', index=False)
        comparison_results_sl.to_excel(writer, sheet_name='DA_SL_Comparison', index=False)

        # Access the workbook and the worksheets
        workbook = writer.book
        worksheet_catalogue = writer.sheets['DA_catalogue_Comparison']
        worksheet_sl = writer.sheets['DA_SL_Comparison']

        # Apply auto-adjustment and conditional formatting
        auto_adjust_columns(worksheet_catalogue, comparison_results_catalogue)
        apply_conditional_formatting(worksheet_catalogue, comparison_results_catalogue, workbook)
        auto_adjust_columns(worksheet_sl, comparison_results_sl)
        apply_conditional_formatting(worksheet_sl, comparison_results_sl, workbook)

        print(f"Comparison report saved to comparison_results.xlsx")
