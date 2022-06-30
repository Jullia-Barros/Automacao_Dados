!pip install pyautogui
!pip install pyperclip
!pip install pandas 

import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)
pyautogui.position()

# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=334, y=252, clicks=2)

time.sleep(3)

#Passo 3: Exportar/fazer download da base de dados

pyautogui.click (x=357, y=253, clicks=1) #clicar em vendas
pyautogui.click (x=1769, y=159) #clicar nos 3 pontinhos
pyautogui.click (x=1524, y=561) #clicar para fazer download
time.sleep (5) #esperar para fazer o download

#Passo 4: Importar a base de dados para o python

import pandas as pd
tabela = pd.read_excel(r"C:\Users\Instituto Criar\Downloads\Vendas - Dez.xlsx"

display (tabela) #Exibi (imprimi) a tabela.

#Passo 5: Calcular os indicadores

#faturamento - soma da coluna valor final

faturamento=tabela ["Valor Final"].sum()

#quantidade de produtos - soma da coluna quantidade

quantidade = tabela ["Quantidade"].sum()
print (quantidade)

#Passo 6: Enviar um email para a diretoria com o relatório
pyautogui.hotkey ("Ctrl", "t")
pyperclip.copy ("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey ("Ctrl", "v")
pyautogui.press ("enter")

time.sleep (5)

pyautogui.click(x=80, y=171)

# escreve o email destinatário

pyautogui.write ("testedeprojetos11@hotmail.com")
pyautogui.press ("tab") # selecionando o destinatário
pyautogui.press ("tab") # passando para o assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey ("Ctrl" , "v")

#escreve o corpo do email

pyautogui.press ("tab")
texto= f""" 
 Olá prezados, bom dia
 
 O faturamento de ontem foi de: R$ {faturamento:,.2f} 
 A quantidade de produtos foi de:{quantidade:,}
 
 Abs
 Jullia Barros"""
pyperclip.copy (texto)
pyautogui.hotkey ("Ctrl", "v")
pyautogui.hotkey ("ctrl", "enter")
