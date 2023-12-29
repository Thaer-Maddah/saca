import pandas as pd


def read_excel_and_check(file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)


        # Iterate through each cell in the DataFrame
        for row_index, row in df.iterrows():
            for col_name, cell_value in row.items():
                # Check the value and content of each cell
                print(f"Row: {row_index}, Column: {col_name}, Value: {cell_value}")


    except Exception as e:
        print(f"Error: {e}")


# Replace 'your_excel_file.xlsx' with the path to your Excel file
read_excel_and_check('../test/test.xlsx')
