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
                print(f"Row: {row_index + 1}, Column: {col_index + 1}, Value: {cell_value}, Formula: {cell.internal_value}")

    except Exception as e:
        print(f"Error: {e}")


# Replace 'your_excel_file.xlsx' with the path to your Excel file
if __name__ == "__main__":
    read_excel_and_check('../test/test.xlsx')
