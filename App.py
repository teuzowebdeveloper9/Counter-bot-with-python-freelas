##the client asked to automate the process to fill in the informations from a spreadsheet to a website I used a program that tells me the position my mouse is in ##

## Obviously when I got the task I remembered the autogui to automate my clicks and the mouse and when I saw that I was going to use a spreadsheet I remembered openxl##

##the time to sleep is to take breaks from automation##

##I downloaded a plugin to read the spreadsheet but this is not mandatory.##



import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Upload the Excel file and spreadsheet
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Iterate through the lines starting from the second (ignoring the header)
for linha in sheet_produtos.iter_rows(min_row=2, values_only=True):
    nome_produto = linha[0]
    descricao = linha[1]
    categoria = linha[2]
    codigo_produto = linha[3]
    dimensoes = linha[4]
    sleep(5)
    preco = linha[5]
    quantidade = linha[6]
    quantidade_em_estoque = linha[7]
    data_de_validade = linha[8]
    cor = linha[9]
    tamanho = linha[10]
    pyautogui.click(1699,561,duration=0.5)
    if tamanho== 'pequeno'
    1739,563
    elif tamanho=='médio'
    1749,579
     else
    1719,569
    material = linha[11]
    sleep(5)
    fabricante = linha[12]
    pais_fabricante = linha[13]
    pais_de_origem = linha[14]
    observacoes = linha[15]
    codigo_de_barras = linha[16]
    localizacao_armazenamento = linha[17]
    
    # List with values and positions
    valores_posicoes = [
        (nome_produto, (1579, 451)),
        (descricao, (1581, 453)),
        (categoria, (1583, 464)),
        (codigo_produto, (1586, 467)),
        (dimensoes, (1593, 477)),
        (preco, (1604, 496)),
        (quantidade, (1615, 505)),
        (quantidade_em_estoque, (1622, 512)),
        (data_de_validade, (1648, 539)),
        (cor, (1661, 551)),
        (tamanho, (1672, 561)),
        (material, (1682, 573)),
        (fabricante, (1701, 582)),
        (pais_fabricante, (1713, 593)),
        (pais_de_origem, (1726, 608)),
        (observacoes, (1739, 622)),
        (codigo_de_barras, (1749, 634)),
        (localizacao_armazenamento, (1755, 654)),
    ]

    # Fill in the fields automatically
    for valor, posicao in valores_posicoes:
        if valor:  # Evitar valores vazios
            pyperclip.copy(str(valor))
            pyautogui.click(*posicao, duration=1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5) #a little while because I want to prevent mistakes#
    
    
    #I achieved :)