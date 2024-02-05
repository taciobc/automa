import openpyxl
import pyperclip
import pyautogui
from time import sleep

#Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios (1).xlsx')
sheet_produtos = workbook['Produtos']
#Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
   
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(279,375,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(232,472,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(287,591,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(280,681,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(321,764,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(294,844,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Botao de proximo
    pyautogui.click(106,903,duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(245,388,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_em_estoque = linha[7].value 
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(241,489,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(225,553,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(272,657,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(253,738,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(239,743,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(225,737,duration=1)
    else:
        pyautogui.click(232,736,duration=1)

    material= linha[11].value
    pyperclip.copy(material)
    pyautogui.click(151,830,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Botao de proximo
    pyautogui.click(83,891,duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(230,420,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_de_origem = linha[13].value
    pyperclip.copy(pais_de_origem)
    pyautogui.click(214,496,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(288,597,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(243,716,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_no_armazem = linha[16].value
    pyperclip.copy(localizacao_no_armazem)
    pyautogui.click(146,810,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Botao de proximo
    pyautogui.click(109,863,duration=1)

    #Botao de ok
    pyautogui.click(500,226,duration=1)

    #Botao de adicionar mais produtos
    pyautogui.click(357,631,duration=1)



#Repetir esses passos para outros campos até preencher campos daquela página
#Clicar em próxima
#Repetir os mesmos passos e ir para a próxima página
#Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
#Clicar em ok para finalizar
#Clicar e ok de novo
#Clicar em adicionar mais um ate finalizar a planilha