import pandas as pd
import os

vendas_DF = pd.read_excel('Vendas_Jan.xlsx')

vendas_DF = vendas_DF.drop(columns=['Produto', 'Data Venda'])

vendas_DF = vendas_DF.groupby('Vendedor').sum()

vendas_DF.to_excel('Relatorio.xlsx')

os.startfile('Relatorio.xlsx')