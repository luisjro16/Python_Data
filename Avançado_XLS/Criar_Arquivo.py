from openpyxl import load_workbook
from openpyxl import Workbook as WB
import string
import os

wb_dados = load_workbook(filename='Quebrar.xlsx')
sheetDados = wb_dados['Dados']

nomeCliente = ""
totalLinha = len(sheetDados['A']) + 1

for linha in range(2, totalLinha):
    nomeAtual = sheetDados['A%s' % linha].value

    if nomeAtual == nomeCliente:
        linhaSheetQuebra = len(sheetRelatorio['A']) + 1

        cColumA = "A" + str(linhaSheetQuebra)
        cColumB = "B" + str(linhaSheetQuebra)
        cColumC = "C" + str(linhaSheetQuebra)

        sheetRelatorio[cColumA] = sheetDados['A%s' % linhaSheetQuebra].value
        sheetRelatorio[cColumB] = sheetDados['B%s' % linhaSheetQuebra].value
        sheetRelatorio[cColumC] = sheetDados['C%s' % linhaSheetQuebra].value

        wb_cliente.save(filename=nomeCliente+'.xlsx')
    else:

        wb_cliente = WB()
        sheetRelatorio = wb_cliente.active
        sheetRelatorio.title = 'Relatorio'

        nomeCliente = sheetDados['A%s' % linha].value

        for letra in string.ascii_uppercase:
            coluna = letra+'1'

            if sheetDados[coluna] != "":
                sheetRelatorio[coluna] = sheetDados[coluna].value
            else:
                break

        sheetRelatorio['A2'] = sheetDados['A%s' % linha].value
        sheetRelatorio['B2'] = sheetDados['B%s' % linha].value
        sheetRelatorio['C2'] = sheetDados['C%s' % linha].value

        wb_cliente.save(filename=nomeCliente+'.xlsx')
