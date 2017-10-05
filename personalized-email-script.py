import openpyxl

 wb = openpyxl.load_workbook('example.xlsx')

 wb.get_sheet_names()

 sheet = wb.get_sheet_by_name('Sheet3')

 sheet['A1']

 print sheet['A1'].value