# Chatgpt-3.5-turbo
# script has approved 
# Executte time
# real    0m1.232s
# user    0m1.106s
# sys     0m0.352s

import pandas as pd
import openpyxl

def read_excel_and_check(file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path, engine='openpyxl', header=None)

        # Open the Excel file with openpyxl for direct access to cell values and formulas
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        sheet = workbook.active

        # Iterate through each cell in the DataFrame
        for row_index, row in df.iterrows():
            for col_index, cell_value in enumerate(row):
                # Get the cell object from openpyxl
                cell = sheet.cell(row=row_index + 1, column=col_index + 1)
                
                # Check the value and formula of each cell
                # cell.data_type n for none, s for string, f for formula, must set data_only=False
                # cell.internal_value attribute display formulas, must set data_only=False  
                print(f"Row: {row_index + 1}, Column: {col_index + 1}, Value: {cell_value}, Formula: {cell.data_type}")
                # convert cell content to string to find formula
                if ( str(cell.value).find("=(B5/$F$6") != -1):
                    print("Thats True", cell.value)

    except Exception as e:
        print(f"Error: {e}")

# Replace 'your_excel_file.xlsx' with the path to your Excel file
read_excel_and_check('../test/test.xlsx')
