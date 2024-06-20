import pandas as pd

vendas = pd.DataFrame({'produto': ['TV', 'Celular', 'Microondas', 'Fogão'], 'valor': [1000, 500, 300, 800]})
produtos = pd.DataFrame({'produto': ['TV', 'Celular', 'Geladeira', 'Fogão'], 'categoria': ['Eletrônicos', 'Eletrônicos', 'Eletrodomésticos', 'Eletrodomésticos'], 'preço': [1200, 600, None, 700]})

print(vendas, '\n', produtos)

resumo = vendas.merge(produtos, how='inner', on='produto')

print(resumo)
