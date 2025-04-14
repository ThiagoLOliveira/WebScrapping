import pyautogui
import pyperclip
import requests
import pandas as pd
import os
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from bs4 import BeautifulSoup

if not os.path.exists("pages"):
    os.makedirs("pages")
#Abre o Chrome
chrome_icon = pyautogui.locateOnScreen('img\\chrome_icon.png', confidence=0.8)

if chrome_icon:
    pyautogui.click(pyautogui.center(chrome_icon))
    print("Chrome clicado!")
else:
    print("Ícone do Chrome não encontrado.")

sleep(1.5)


#Abre uma pagina anonima
modo_visitante = pyautogui.locateOnScreen('img\\modo_visitante.png', confidence=0.8)

if modo_visitante:
    print(modo_visitante)
    pyautogui.click(pyautogui.center(modo_visitante))
    print("Modo Visitante clicado!")
else:
    print("Ícone do Chrome não encontrado.")

sleep(0.7)

# nova_pagina = pyautogui.locateOnScreen('img\\nova_pagina.png', confidence=0.8)

#Acessa o html da pagina
pyautogui.write(r'https://www.zapimoveis.com.br/venda', interval=0.1)
pyautogui.press('enter')
for i in range(17):
    pyautogui.scroll(-2100)
    print(i)
    sleep(6)


#Seleciona tudo e copia o html
pyautogui.hotkey('ctrl', 'u')
sleep(10)
pyautogui.hotkey('ctrl', 'a')
sleep(1)
pyautogui.hotkey('ctrl', 'c')


#Armazena o que foi copiado
html = pyperclip.paste()


#Escreve em um arquivo o html copiado
with open(f"pages\\pagina_zapimoveis_page1.html", "w", encoding="utf-8") as f:
    f.write(html)

pages = 10

for i in range(1, pages):
    link = pyautogui.locateOnScreen('img\\link.png', confidence=0.8)

    if link:
        pyautogui.click(pyautogui.center(link))
    
    if i <= 9:
        pyautogui.press('right')
        pyautogui.press('backspace')
        pyautogui.write(str(i + 1), interval=0.1)
    else:
        pyautogui.press('right')
        pyautogui.press('backspace', presses=2)
        pyautogui.write(str(i + 1), interval=0.1)

    pyautogui.press('enter')

    sleep(5)
    pyautogui.hotkey('ctrl', 'a')
    sleep(1)
    pyautogui.hotkey('ctrl', 'c')

    html = pyperclip.paste()

    with open(f"pages\\pagina_zapimoveis_page{i}.html", "w", encoding="utf-8") as f:
        f.write(html)


df = pd.DataFrame()

all_links_extends = []
for i in range(len(os.listdir("pages"))):
    with open(f"pages\\pagina_zapimoveis_page{i + 1}.html", "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, 'html.parser')
        links_tag_a = [a['href'] for a in soup.find_all('a', href=True)
                if a['href'].startswith("https://www.zapimoveis.com.br/imovel/")]

        script_links = []
        for script in soup.find_all("script"):
            if script.string:
                found = re.findall(r"https://www\.zapimoveis\.com\.br/imovel/[^\s\"']+", script.string)
                script_links.extend(found)

        all_links = list(set(links_tag_a + script_links))
        all_links_extends.extend(all_links)
        

all_links_extends = list(set(all_links_extends))

for url in all_links_extends:
    try:
        driver = webdriver.Chrome()
        driver.get(url)
        sleep(5)
        
        elements = driver.find_elements(By.CLASS_NAME, "amenities-item-text")
        property_data = {"Link": url}
        
        for element in elements:
            text = element.text.lower().strip()
            
            if "m²" in text:
                property_data["Área"] = re.search(r'\d+', text).group() if re.search(r'\d+', text) else None
            elif "quartos" in text:
                property_data["Quartos"] = re.search(r'\d+', text).group() if re.search(r'\d+', text) else None
            elif "banheiros" in text:
                property_data["Banheiros"] = re.search(r'\d+', text).group() if re.search(r'\d+', text) else None
            elif "vagas" in text:
                property_data["Vagas"] = re.search(r'\d+', text).group() if re.search(r'\d+', text) else None
            elif "suítes" in text or "suites" in text:
                property_data["Suítes"] = re.search(r'\d+', text).group() if re.search(r'\d+', text) else None
        
        df = pd.concat([df, pd.DataFrame([property_data])], ignore_index=True)
        print(df)
    except Exception as e:
        print(f"Error processing {url}: {str(e)}")
    finally:
        driver.quit()