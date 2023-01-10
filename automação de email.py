import pyperclip
import pyautogui
import time


#abrindo o chrome
pyautogui.PAUSE = 1
pyautogui.press("winleft")
pyautogui.write("chrome")
pyautogui.press("enter")
pyautogui.hotkey('Ctrl', 'shift', 'n')

# abrindo o google drive

link = 'https://accounts.google.com/v3/signin/identifier?dsh=S-486252896%3A1673360906910703&continue=http%3A%2F%2Fdrive.google.com%2F%3Futm_source%3Den&ltmpl=drive&passive=true&service=wise&usp=gtd&utm_campaign=web&utm_content=gotodrive&utm_medium=button&flowName=GlifWebSignIn&flowEntry=ServiceLogin&ifkv=AeAAQh5-pLmMb7rarWv-A1gJ0vw4AJUKDkgv9b0MlWxI3f9uEvkPN2u-h7qsJCu4Tgwx2wd3Hw2p2A'
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(1)

# login 
pyautogui.write('Victoriadoesantosilva@gmail.com')
pyautogui.press('enter')
time.sleep(1)
pyautogui.write('VICcn@188')
pyautogui.press('enter')
time.sleep(3)

# acessando arquivo e baixando
pyautogui.click(1158, 1725, clicks = 2)
pyautogui.click(1032, 961, button = 'right')
pyautogui.click(1584, 1811, clicks = 2)
time.sleep(5)

#calculando 
import pandas as pd 

df = pd.read_excel(r'C:\Users\Victória\Downloads\Vendas - Dez.xlsx')
print(df)
faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()

# abrindo gmail
pyautogui.hotkey('Ctrl', 't')
pyautogui.write('mail.google.com')
pyautogui.press('enter')
time.sleep(5)

#escrevendo email
pyautogui.click(299, 676 )
pyautogui.write('khaleesitar@protinmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
assunto = 'Relátorio de Vendas'
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
corpo = f'''Prezados, Bom dia!
O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {qtde_produtos:,} 

abs
Victória'''
pyperclip.copy(corpo)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')






