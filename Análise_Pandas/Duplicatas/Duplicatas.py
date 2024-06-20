import pandas as pd
import os

baseVendas_DF = pd.read_excel('Base_Vendas.xlsx')

baseVendas_DF["Duplicatas"] = baseVendas_DF.duplicated(subset="Vendedor")

#print(baseVendas_DF)

for id in range(len(baseVendas_DF)):
    if baseVendas_DF.loc[id, "Duplicatas"] == True:
        nome = baseVendas_DF.loc[id, "Vendedor"]
        valor_venda = baseVendas_DF.loc[id, "Total Vendas"]

        baseVendas_DF = baseVendas_DF.drop(index=id)

        baseVendas_DF.loc[baseVendas_DF['Vendedor'] == nome, "Total Vendas"] += valor_venda 

baseVendas_DF = baseVendas_DF.drop(columns="Duplicatas")
baseVendas_DF['Situacao'] = None

baseVendas_DF.loc[baseVendas_DF['Total Vendas'] > 800, 'Situacao'] = "Positivo"
baseVendas_DF.loc[baseVendas_DF['Total Vendas'] <= 800, 'Situacao'] = "Negativo"

baseVendas_DF = baseVendas_DF.sort_values(by="Vendedor")

print(baseVendas_DF)
