import pyautogui
import pyperclip
import time
import pandas
import numpy
import openpyxl

#Guia para o PyAutoGUI: pyautogui.readthedocs.io/

pyautogui.PAUSE = 1.5

#Acessar o Drive em questão

linkdrive = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyautogui.press("win")
pyautogui.write("Google Chrome")
pyautogui.press("enter")
time.sleep(4)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy(linkdrive)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

#Carregar o Sistema com Delay de 5 segundos
time.sleep(5)

#Abrir a pasta no Drive e Navegar nela
pyautogui.click(x=322, y= 286, clicks=2)
time.sleep(2)

#Fazer o Download do arquivo
pyautogui.click(x=671, y=193)
pyautogui.click(x=696, y=586)

#Aguardar o Download
time.sleep(20)

#Confirmar o Download
pyautogui.press("enter")

#Ler o arquivo em Excel e Importar para o Código
tabela = pandas.read_excel(r"C:\Users\Joao Pedro\Documents\General Programming\Python Projects\Automatização de Vendas\Aula 1-20220110T230732Z-001\Aula 1\Exportar\Vendas - Dez.xlsx")
print(tabela)

#Calcular o faturamento e quantidade de produtos
faturamento = tabela["Valor Final"].sum()
qtde_produtos = tabela["Quantidade"].sum()
print(faturamento)
print(qtde_produtos)

#Escrever e enviar o E-mail
linkemail = "https://mail.google.com/mail/u/0/#inbox"
destinatario = "carlaalecoi@gmail.com"
assunto = "Teste do Script"
texto = f"""
Mae,

Aqui esta o faturamento do pack do pe do mes de Dezembro: R$ {faturamento:,.2f}
Vendendo: {qtde_produtos:,} unidades

Beijos
"""
pyautogui.hotkey("ctrl", "t")
pyperclip.copy(linkemail)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=85, y=209)
pyautogui.write(destinatario)
pyautogui.hotkey("tab") #seleciona o e-mail
pyautogui.hotkey("tab") #muda para o campo de assunto
pyautogui.write(assunto)
pyautogui.hotkey("tab") #muda para o campo corpo do texto
pyautogui.write(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")