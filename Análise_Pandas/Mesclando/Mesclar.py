import pandas as pd
import os

vendasJan_DF = pd.read_excel('Vendas_Jan.xlsx')
vendasJan_DF = vendasJan_DF.sort_values(by="Data Venda")

vendasFev_DF = pd.read_excel('Vendas_Fev.xlsx')
vendasFev_DF = vendasFev_DF.sort_values(by="Data Venda")

vendasMar_DF = pd.read_excel('Vendas_Mar.xlsx')
vendasMar_DF = vendasMar_DF.sort_values(by="Data Venda")

relatorio_DF = pd.DataFrame()
relatorio_DF = pd.concat([vendasJan_DF, vendasFev_DF, vendasMar_DF], keys=['Janeiro', 'Fevereiro', 'Marco'])
relatorio_DF = relatorio_DF.drop(columns='Produto')

print(relatorio_DF)