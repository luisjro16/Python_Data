import pandas as pd

vendas = pd.DataFrame({'produto': ['TV', 'Celular', 'TV', 'Microondas', 'Celular'],
                   'valor_unitario': [1000, 500, 1200, 300, 600],
                   'regiao': ['Sul', 'Sul', 'Sudeste', 'Nordeste', 'Sudeste']})

# Calcula a soma da quantidade vendida por produto e por região
tabela_dinamica = vendas.pivot_table(values='valor_unitario', index='produto', columns='regiao')

print(vendas,'\n')
print(tabela_dinamica)


