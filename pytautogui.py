import pyautogui
import pyperclip
import requests
import pandas as pd
import os
import re
import mysql.connector
import dotenv
import pyperclip
from pyscreeze import ImageNotFoundException
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from bs4 import BeautifulSoup

# #Conexão com o banco de dados
dotenv.load_dotenv()
host = os.getenv("DB_HOST_LOCAL")
user = os.getenv("DB_USERNAME_LOCAL")
database = os.getenv("DB_DATABASE_LOCAL")

pyautogui.PAUSE = 1.5
pg = 1
links = []
flag = True
for pagina in range(1, 30):
    while flag:
        try:
            falha = pyautogui.locateOnScreen('vivareal\\falha.png', confidence=0.9)
            if falha:
                print("Imagem de falha detectada. Encerrando o loop.")
                flag = False
                break
        except pyautogui.ImageNotFoundException:
            try:
                sleep(3)
                contatar = list(pyautogui.locateAllOnScreen('vivareal\\vivareal.png', confidence=0.9))
                if contatar:
                    print(f"Encontrado(s) {len(contatar)} contato(s) na tela.")
                    for i in contatar:
                        try:
                            pyautogui.moveTo(i[0] + 10, i[1] + 10)
                            pyautogui.rightClick()
                            copiar_btn = pyautogui.locateOnScreen('vivareal\\copiar_link.png', confidence=0.9)
                            if copiar_btn:
                                pyautogui.click(pyautogui.center(copiar_btn))
                                copied_link = pyperclip.paste()
                                links.append(copied_link)
                                print(f"Link copiado: {copied_link}")
                                print('Lista de links:', links)
                            else:
                                print("Botão 'copiar link' não encontrado.")
                        except Exception as e:
                            print(f"Erro ao tentar copiar link: {e}")
                            continue
                    pyautogui.scroll(-1000)
            except ImageNotFoundException:
                try:
                    links__ = list(pyautogui.locateAllOnScreen('vivareal\\links_imoveis.png', confidence=0.9))
                    if links__:
                        for i in links__:
                            pyautogui.moveTo(i[0], i[1])
                            pyautogui.click()
                            sleep(2)
                            pegar_link = pyautogui.locateOnScreen('vivareal\\link_imoveis.png', confidence=0.9)
                            if pegar_link:
                                pyautogui.rightClick(pyautogui.center(pegar_link))
                                copiar_btn = pyautogui.locateOnScreen('vivareal\\copiar_link.png', confidence=0.9)
                                if copiar_btn:
                                    pyautogui.click(pyautogui.center(copiar_btn))
                                    copied_link = pyperclip.paste()
                                    links.append(copied_link)
                                    print(f"Link copiado: {copied_link}")
                                    pyautogui.press('esc')
                        pyautogui.scroll(-1000) 
                except ImageNotFoundException:
                    sleep(3)
                    pg += 1
                    print("Nenhum contato encontrado nesta rolagem.")
                    pagina = pyautogui.locateOnScreen('vivareal\\pagina.png', confidence=0.9)
                    if pagina:
                        pyautogui.click(pyautogui.center(pagina))
                        sleep(3)
                        pyautogui.press('right')
                        if pg > 9:
                            pyautogui.press('backspace', presses=2)
                        else:
                            pyautogui.press('backspace')
                        pyautogui.write(str(pg))
                        pyautogui.press('enter')
                        continue
# Salvando os links em um arquivo Excel
df_links = pd.DataFrame(links, columns=["Links"])
df_links.to_excel("links_vivareal.xlsx", index=False)
print("Arquivo 'links_vivareal.xlsx' salvo com sucesso.")
