from openpyxl import Workbook
from openpyxl import load_workbook

def delete_bad_columns():
    wb = load_workbook(filename = 'appointmentsReport.xlsx')
    sheet = wb['appointmentsReport']
    sheet.delete_cols(3, 2)
    sheet.delete_cols(4, 4)
    sheet.delete_cols(5, 7)
    sheet.delete_cols(6, 5)
    sheet.delete_rows(1,1)
    wb.save('appointmentsReport.xlsx')

delete_bad_columns()
