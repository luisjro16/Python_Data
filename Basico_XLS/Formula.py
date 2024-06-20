import openpyxl as op
import os
from openpyxl.styles import PatternFill

wb_open = op.load_workbook('Vendedores.xlsx')
sheetVendas = wb_open['Vendas']
sheetResumo = wb_open['Resumo']

somarLA = 0
somarEM = 0
somarNP= 0
somarAM = 0

for linha in range(2, len(sheetVendas['A'])+1):
    if sheetVendas['A%s' % linha].value == "Amanda Martins":
        somarAM += sheetVendas['C%s' % linha].value
    if sheetVendas['A%s' % linha].value == "Eliane Moreira":
        somarEM += sheetVendas['C%s' % linha].value
    if sheetVendas['A%s' % linha].value == "Nicolas Pereira":
        somarNP += sheetVendas['C%s' % linha].value
    if sheetVendas['A%s' % linha].value == "Leonardo Almeida":
        somarLA += sheetVendas['C%s' % linha].value

for linha in range(2, len(sheetResumo['A'])+1):
    if sheetResumo['A%s' % linha].value == "Amanda Martins":
        sheetResumo['B%s' % linha] = somarAM
    if sheetResumo['A%s' % linha].value == "Eliane Moreira":
        sheetResumo['B%s' % linha] = somarEM
    if sheetResumo['A%s' % linha].value == "Nicolas Pereira":
        sheetResumo['B%s' % linha] = somarNP
    if sheetResumo['A%s' % linha].value == "Leonardo Almeida":
        sheetResumo['B%s' % linha] = somarLA

print(f"{somarLA}")

wb_open.save('Vendedores.xlsx')

os.startfile('Vendedores.xlsx')