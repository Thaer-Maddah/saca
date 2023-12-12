# google bard script
# script has approved 
# Executte time without pandas module
# real    0m0.733s
# user    0m0.693s
# sys     0m0.287s

import pandas as pd
import openpyxl

def evaluate_cell_formulas(sheet_name):
    # Loop through each cell
    for row in ws.rows:
        for cell in row:
            # Check if the cell has a formula
            if cell.data_type == 'f':
                # Print the cell address and the formula
                print(f"Cell {cell.coordinate} contains a formula: {cell.value}")
            else:
                # Print the cell address and the value
                print(f"Cell {cell.coordinate} does not contain a formula: {cell.value}")


# This method find specific formula by full string
def eval_cell_1(sheet_name):
    # Loop through each cell
    for row in ws.rows:
        for cell in row:
            # Check if the cell has a formula
            if cell.data_type == 'f':
                if cell.value == "=(B5/$F$6)*100":
                    # Print the cell address and the formula
                    print(f"Cell {cell.coordinate} contains a formula: {cell.value}")
            else:
                # Print the cell address and the value
                pass
     

# This method find specific formula by substring
def eval_cell_2(sheet_name):
    # Loop through each cell
    for row in ws.rows:
        for cell in row:
            # Check if the cell has a formula
            if ( str(cell.value).find("=(B5/$F$6") != -1):# and cell.value == 6.157635467980295:
                print(f"Cell {cell.coordinate} contains a formula: {cell.value}")

# This method find specific formula by substring
def eval_cell_3(sheet_name):
    # Loop through each cell
    for row_index, row in df.iterrows():
        for col_index, cell_value in enumerate(row):
            cell = sheet_name.cell(row=row_index + 1, column=col_index + 1)
            # Check if the cell has a formula
            if ( str(cell.value).find("=(B5/$F$6") != -1):# and cell.value == 6.157635467980295
                if round(cell_value, 2) == 6.16: # round the number for 2 decimals
                    print(f"Cell {cell.coordinate} contains a formula: {cell_value}")


if __name__ == "__main__":
    # Define the filename and sheet name
    filename = "test.xlsx"
    #sheet_name = "sheet1"
    # Load the workbook
    df = pd.read_excel(filename, engine='openpyxl', header=None)
    wb = openpyxl.load_workbook(filename, data_only=False)
    ws = wb.active
    #ws = wb[sheet_name]
    evaluate_cell_formulas(ws)
    eval_cell_3(ws)
