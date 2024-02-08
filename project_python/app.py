import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook ['Produtos']

for linha in sheet_produtos.iter(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(832,339, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(296,439, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(152,565, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(148,647, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(150,736, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(150,822, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(181,879, duration=1)
    sleep(5)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(152,385, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(150,476, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(145,559, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(152,634, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    tamanho = linha[10].value
    if tamanho == 'Pequeno':
        pyautogui.click(168, 759, duration=1)
    elif tamanho == 'Medio':
        pyautogui.click(163, 781, duration=1)
    else:
        pyautogui.click(168, 759, duration=1)
       
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(151,820, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(177,873, duration=1)
    sleep(5)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(154,405, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(148,491, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(149,572, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(150,712, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(154,799, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Botao de concluir!
    pyautogui.click(174,857, duration=1)
    #Botao de Ok!
    pyautogui.click(656,612, duration=1)
    #Iniciar nova adicao de produto!
    pyautogui.click(482,612, duration=1)
    
    

