import pandas as pd
import os

vendas_DF = pd.read_excel("Vendas_Jan.xlsx")
relatorio_DF = pd.DataFrame()

#print(vendas_DF)

del vendas_DF["Produto"]

Vendedores = list()

for nome in vendas_DF["Vendedor"]:
    if nome in Vendedores:
        continue
    else:
        Vendedores.append(nome)

relatorio_DF["Vendedor"] = Vendedores

for nome1 in relatorio_DF["Vendedor"]:
    total = 0
    cont = 0

    for nome2 in vendas_DF["Vendedor"]:

        if nome1 == nome2:
            total = total + vendas_DF.loc[cont, "Total Vendas"]
        cont = cont + 1
    
    relatorio_DF.loc[relatorio_DF["Vendedor"] == nome1, "Total"] = total

print(relatorio_DF)
