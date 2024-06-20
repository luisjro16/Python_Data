import openpyxl as op
import pyperclip as copy
import pyautogui as auto

product_WB = op.load_workbook('produtos_ficticios.xlsx')
sheet_product = product_WB['Produtos']

auto.press('win')
auto.press('enter')
auto.write('https://cadastro-produtos-devaprender.netlify.app/index.html')
auto.press('enter')
auto.sleep(4)

cont = 0

for linha in sheet_product.iter_rows(min_row=2):
    
    #Nome
    nome_produto = linha[0].value
    copy.copy(nome_produto)
    auto.click(x=767, y=320)
    auto.hotkey('ctrl', 'v')
    
    #Descricao
    descricao = linha[1].value
    copy.copy(descricao)
    auto.click(x=767, y=428)
    auto.hotkey('ctrl', 'v')

    auto.sleep(1)

    #Categoria
    categoria = linha[2].value
    copy.copy(categoria)
    auto.click(x=767, y=569)
    auto.hotkey('ctrl', 'v')
    
    #Codigo
    codigo = linha[3].value
    copy.copy(codigo)
    auto.click(x=767, y=677)
    auto.hotkey('ctrl', 'v')

    #Peso
    peso = linha[4].value
    copy.copy(peso)
    auto.click(x=767, y=781)
    auto.hotkey('ctrl', 'v')

    #Dimensao
    dimensao = linha[5].value
    copy.copy(dimensao)
    auto.click(x=767, y=888)
    auto.hotkey('ctrl', 'v')

    #Proximo
    auto.moveTo(x=170, y=968)
    auto.click()
    auto.sleep(2)

    #Preco
    preco = linha[6].value
    copy.copy(preco)
    auto.moveTo(x=338, y=330)
    auto.click()
    auto.hotkey('ctrl', 'v')

    #Quantidade
    quantidade = linha[7].value
    copy.copy(quantidade)
    auto.moveTo(x=767, y=436)
    auto.click()
    auto.hotkey('ctrl', 'v')

    #Validade
    validade = linha[8].value
    copy.copy(validade)
    auto.click(x=767, y=553)
    auto.hotkey('ctrl', 'v')

    #Cor
    cor = linha[9].value
    copy.copy(cor)
    auto.click(x=767, y=638)
    auto.hotkey('ctrl', 'v')

    #Tamanho
    tamanho = linha[10].value
    copy.copy(tamanho)
    auto.click(x=254, y=757, duration=0.5)
    if tamanho == 'Pequeno':
        auto.click(x=235, y=803, duration=1)
    elif tamanho == 'Médio':
        auto.click(x=204, y=830, duration=1)
    else:
        auto.click(x=202, y=855, duration=1)

    #Material
    validade = linha[11].value
    copy.copy(validade)
    auto.click(x=767, y=865)
    auto.hotkey('ctrl', 'v')

    #Proximo
    auto.moveTo(x=170, y=918)
    auto.click()
    auto.sleep(2)

    #Fabricante
    fabricante = linha[12].value
    copy.copy(fabricante)
    auto.click(x=767, y=351)
    auto.hotkey('ctrl', 'v')

    #Origem
    origem = linha[13].value
    copy.copy(origem)
    auto.click(x=767, y=452)
    auto.hotkey('ctrl', 'v')

    #observacoes
    observacoes = linha[14].value
    copy.copy(observacoes)
    auto.click(x=767, y=566)
    auto.hotkey('ctrl', 'v')

    #Codigo
    codigo = linha[15].value
    copy.copy(codigo)
    auto.click(x=881, y=734)
    auto.hotkey('ctrl', 'v')

    #Localizacao
    localizacao = linha[16].value
    copy.copy(localizacao)
    auto.click(x=767, y=830)
    auto.hotkey('ctrl', 'v')

    #botao concluir
    auto.click(x=169, y=910)
    auto.press('enter')
    auto.sleep(2)
    auto.click(x=919, y=629)
    auto.sleep(3)

    
