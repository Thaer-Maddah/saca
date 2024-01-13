#!/usr/bin/python3
##
## Revise excel files
## http://github.com/Thaer-Maddah
##
## Copyright (C) 2023 Thaer Maddah. All rights reserved.

## Permission is hereby granted, free of charge, to any person obtaining
## a copy of this software and associated documentation files (the
## "Software"), to deal in the Software without restriction, including
## without limitation the rights to use, copy, modify, merge, publish,
## distribute, sublicense, and/or sell copies of the Software, and to
## permit persons to whom the Software is furnished to do so, subject to
## the following conditions:
##
## The above copyright notice and this permission notice shall be
## included in all copies or substantial portions of the Software.
##
## THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
## EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
## MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
## IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
## CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
## TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
## SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
##
## [ MIT license: http://www.opensource.org/licenses/mit-license.php ]
##

import sys
import pandas as pd
import numpy as np
import browse_files as bf
import openpyxl
#from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import time
#import win32com.client as client
#import write_grades as wr


#excel_file='files/excel.xlsx'

folder = 'Assign/test'
ext = '.xlsx'
trim_txt = '/mnt/c/code/Assign/'
data = [[1400, 880, 630], [2700, 780, 1080], [2500, 900, 1375], [2700, 805, 1755]]
grade = []

def sheetTitle(ws):
    title1 = 'وظيفة مهارات الحاسوب'
    title2 = 'وظيفه مهارات الحاسوب'
    if ws.title == title1 or ws.title == title2:
        print('Sheet name: 3')
        grade.append(3)
        print(ws.title)
    else:
        print('Sheet name is wrong: 0')
        grade.append(0)

# This method find specific formula by substring
def eval_cell(ws, df, str_formula, str_value ):
    #string = "=SUM("
    # Loop through each cell
    for row_index, row in df.iterrows():
        for col_index, cell_value in enumerate(row):
            cell = ws.cell(row=row_index + 1, column=col_index + 1)
            # Check if the cell has a formula
            if ( str(cell.value).find(str_formula) != -1):# and cell.value == 6.157635467980295
                if round(cell_value, 0) == str_value: # round the number for 2 decimals
                    print(f"Cell {str_formula} {cell.coordinate} contains a formula: {cell_value}")
                    return True
    return False

def is_worksheet_empty(ws):
    """Check if a given worksheet is empty."""
    for row in ws.iter_rows(values_only=True):
        if any(cell for cell in row):
            return False
    return True




def reviseExcel(ws, df):
    str_dict = {
        "=SUM": 196,
        "=AVERAGE": 49,
        "=MAX": 60,
        "=MIN": 37,
        "=LARGE": 53,
        "=SMALL": 46,
        #"=IF(": "D",
        "=COUNT": 15,
        "=COUNTIF": 7,
        "=COUNTIF(": 4
        }

    #sheetTitle(ws)
    for key, value in str_dict.items():
        if eval_cell(ws, df, key, value ):
            grade.append(2)
        else:
            grade.append(0)

    #hasChart(xl, work_book)
    degree = sum(grade[1:len(grade)])
    print(grade)
    #wr.writeExcelGrades([grade]) 
    print('Final degree is:', degree)
    del grade[:]
 

def main():
    counter = 0 
    item = []
    files = []
    files, dirs = bf.browse(ext, folder)
    start = time.time()
    for file, dir in zip(files, dirs):
        path = bf.getFile(file, dir)
        print(path)
        # openpyxl
        # data_only=True return the Cell value
        df = pd.read_excel(path, engine='openpyxl', header=None)
        wb = openpyxl.load_workbook(path, data_only=False)
        ws = wb.active
        grade.append(path.strip(trim_txt))
        reviseExcel(ws, df)
    
        #grade.append(dir + file)
        #grade.append(4)

        # this is the sheet work section 
        #print(f"Active sheet is: {wb.active}")
        #for sheet in wb.worksheets:
        #    #print(f"Processing sheet '{sheet.title}'")
        #    wb.active = wb[sheet.title]
        #    if is_worksheet_empty(wb[sheet.title]):
        #        pass
        #    else:
        #        wb.active = ws
        #        reviseExcel(ws, df)
        #        print(ws)

        counter += 1
        print(counter, 'Excel file revised!')
    
        sep = '='
        print(sep*120)
        #time.sleep(1)
    
    end = time.time()
    print(f"Total time: {round(end - start)} seconds")


if __name__ == '__main__':
    sys.exit(main())
