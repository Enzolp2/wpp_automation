import pyautogui as pg
import webbrowser as web
import pandas as pd
import time

try:
    data = pd.read_excel('contatos.xlsx')
except Exception as e:
    print(e)
    
message = ""
data_dict = data.to_dict('list')
numeros = data_dict['WhatsApp']
nomes = data_dict['firstname']
combo = zip(numeros, nomes)
first = True

count = 0
for numero, nome in combo:
    time.sleep(4)
    message = f"Olá {nome}, testando automação!\n Nº: {numero}"
    web.open(f"https://web.whatsapp.com/send?phone={numero}&text={message}")
    time.sleep(10)
    pg.press('enter')
    time.sleep(1)
    pg.hotkey('ctrl', 'w')
    time.sleep(1)
    pg.press('enter')
    count+=1
    
    
