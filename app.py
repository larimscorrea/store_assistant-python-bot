# Ler os dados da planilha
# Inserir cada célula de cada linha em um campo do sistema
import openpyxl
import pyautogui


workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
sells_sheet = workbook['vendas']

for line in sells_sheet.iter_rows(min_row=2):
    pyautogui.click(x, y, duration=1.5)
    pyautogui.write(line[0].value)
    line[0].value
    line[1].value
    line[2].value
    line[0].value