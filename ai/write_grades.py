##
## Write grades into excel file
## http://github.com/Thaer-Maddah
##
## Copyright (C) 2023 Thaer Maddah. All rights reserved.

import pandas as pd

def writeDocGrades(data):
    excel_headers = ['student',
                    'font face', 
                    'Heading font size',  
                    'Paragraph font size',
                    'Heading font weight',
                    'Heading font underline',
                    'Heading font color',
                    'Paragraph font black color',
                    'Indentation',
                    'Chart',
                    'Lists',
                    'Styling Tables',
                    'Images and shapes',
                    'Header and Footer',
                    'References',
                     'Ideas and Style']
    

    df = pd.DataFrame(data, columns = excel_headers)
    writer = pd.ExcelWriter('files/grades-doc.xlsx', mode='a',if_sheet_exists='overlay', engine='openpyxl')
    df.to_excel(writer, sheet_name='Students Grades', header=None, startrow=writer.sheets['Students Grades'].max_row, index=False)
    writer._save()


def writeExcelGrades(grades):
    excel_headers = ['student',
                    'File name', 
                    'Sheet name',  
                    'Concatenate',
                    'Day',
                    'Month',
                    'Year',
                    'Sum',
                    'Average',
                    'Max',
                    'Min',
                    'Large',
                    'Small',
                    'classification',
                    'Count students',
                    'Count females',
                    'Count names',
                    'Decimal places',
                    'Conditional formatting',
                    'Charts'
                     ]

    df = pd.DataFrame(grades, columns = excel_headers)
    writer = pd.ExcelWriter('files/grades-excel.xlsx', mode='a',if_sheet_exists='overlay', engine='openpyxl')
    df.to_excel(writer, sheet_name='Students Grades', header=None, startrow=writer.sheets['Students Grades'].max_row, index=False)
    writer._save()
