# 1 - Ler dados da planilha
# 2 - Inserir cada célula de cada linha em um campo do sistema

import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    # nome
    pyautogui.click(345, 624, duration=1.5)
    pyautogui.write(linha[0].value)

    # produto
    pyautogui.click(337, 647, duration=1.5)
    pyautogui.write(linha[1].value)

    # quantidade
    pyautogui.click(304, 676, duration=1.5)
    pyautogui.write(str(linha[2].value))
    # O pyautogui não consegue digitar valores numéricos diretamente, por isso o str().

    # categoria
    pyautogui.click(383, 704, duration=1.5)
    pyautogui.write(linha[3].value)

    pyautogui.click(229, 724, duration=1.5)
    pyautogui.click(660, 424, duration=1.5)
