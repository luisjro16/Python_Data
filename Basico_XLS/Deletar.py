import openpyxl as op
import os

name_file = 'plan.xlsx'
wb_open = op.load_workbook(filename=name_file)

sheet_selc = wb_open['Dados']

sheet_selc.delete_rows(3)

wb_open.save(name_file)

os.startfile(name_file)