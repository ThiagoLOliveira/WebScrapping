import pyautogui
import pyperclip
import requests
import pandas as pd
import os
import re
import mysql.connector
import dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from bs4 import BeautifulSoup


# #Verifica se a pasta existe
# if not os.path.exists("pages"):
#     os.makedirs("pages")
    
# #Conexão com o banco de dados
dotenv.load_dotenv()
host = os.getenv("DB_HOST_LOCAL")
user = os.getenv("DB_USERNAME_LOCAL")
database = os.getenv("DB_DATABASE_LOCAL")

connection = mysql.connector.connect(
    host = host,
    user = user,
    database = database,
)

cursor = connection.cursor()

# urls = [
#     'https://www.zapimoveis.com.br/venda/casas/sp+sao-bernardo-do-campo/',
#     'https://www.zapimoveis.com.br/venda/casas/sp+santo-andre/',
#     'https://www.zapimoveis.com.br/venda/casas/sp+carapicuiba++carapicuiba/',
#     'https://www.zapimoveis.com.br/venda/casas/sp+sao-paulo+zona-leste/',
#     'https://www.zapimoveis.com.br/venda/casas/sp+sao-paulo+centro/',
#     'https://www.zapimoveis.com.br/venda/casas/sp+sao-paulo+zona-norte/',
#     'https://www.zapimoveis.com.br/venda/casas/sp+sao-paulo+zona-oeste/',
#     'https://www.zapimoveis.com.br/venda/casas/sp+sao-paulo+zona-sul/',
#     'https://www.zapimoveis.com.br/venda/casas/sp+sao-caetano-do-sul/',
#     'https://www.zapimoveis.com.br/venda/casas/sp+sao-paulo/',
#     'https://www.zapimoveis.com.br/venda/apartamentos/sp+sao-paulo/',
#     'https://www.zapimoveis.com.br/venda/casas/sp+sao-paulo/',
#     ]

# for index, url in enumerate(urls):
#     #Abre o navegador
#     chrome_icon = pyautogui.locateOnScreen('img\\chrome_icon.png', confidence=0.8)
#     if chrome_icon:
#         pyautogui.click(pyautogui.center(chrome_icon))
#         print("Chrome clicado!")
#     else:
#         print("Ícone do Chrome não encontrado.")
#     sleep(1.5)

#     # Abre uma pagina anonima
#     # modo_visitante = pyautogui.locateOnScreen('img\\modo_visitante.png', confidence=0.8)
#     # if modo_visitante:
#     #     print(modo_visitante)
#     #     pyautogui.click(pyautogui.center(modo_visitante))
#     #     print("Modo Visitante clicado!")
#     # else:
#     #     print("Ícone do Chrome não encontrado.")

#     sleep(0.7)

#     #Acessa o html da pagina
#     pyautogui.write(url, interval=0.1)
#     pyautogui.press('enter')

#     screen_width, screen_height = pyautogui.size()
#     pyautogui.moveTo(screen_width / 2, screen_height / 2)

#     for i in range(17):
#         pyautogui.scroll(-2100)
#         print(i)
#         sleep(6)

#     #Seleciona tudo e copia o html
#     pyautogui.hotkey('ctrl', 'u')
#     sleep(10)
#     pyautogui.hotkey('ctrl', 'a')
#     sleep(1)
#     pyautogui.hotkey('ctrl', 'c')


#     #Armazena o que foi copiado
#     html = pyperclip.paste()

#     #Escreve em um arquivo o html copiado
#     with open(f"pages\\pagina_zapimoveis_page1.html", "w", encoding="utf-8") as f:
#         f.write(html)

#     pages = 99

#     for i in range(1, pages + 1):
#         sleep(5)
#         link = pyautogui.locateOnScreen('img\\link.png', confidence=0.8)
#         if link:
#             pyautogui.click(pyautogui.center(link))
#         if i <= 9:
#             pyautogui.press('right')
#             pyautogui.press('backspace')
#             pyautogui.write(str(i + 1), interval=0.1)
#         else:
#             pyautogui.press('right')
#             pyautogui.press('backspace', presses=2)
#             pyautogui.write(str(i + 1), interval=0.1)
#         pyautogui.press('enter')
#         sleep(5)
#         pyautogui.hotkey('ctrl', 'a')
#         sleep(1)
#         pyautogui.hotkey('ctrl', 'c')
#         html = pyperclip.paste()
#         with open(f"pages\\pagina_zapimoveis{index}_page{i}.html", "w", encoding="utf-8") as f:
#             f.write(html)

#     pyautogui.hotkey('alt', 'f4')
    
df = pd.DataFrame()

all_links_extends = []

def extrair_links_do_arquivo(caminho_arquivo):
    try:
        with open(caminho_arquivo, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f, 'html.parser')
            links_tag_a = [a['href'] for a in soup.find_all('a', href=True)
                        if a['href'].startswith("https://www.zapimoveis.com.br/imovel/")]

            script_links = []
            for script in soup.find_all("script"):
                if script.string:
                    found = re.findall(r"https://www\.zapimoveis\.com\.br/imovel/[^\s\"']+", script.string)
                    script_links.extend(found)

            return list(set(links_tag_a + script_links))
    except Exception as e:
        print(f"Erro ao processar {caminho_arquivo}: {e}")
        return []

# Listar todos os arquivos na pasta "pages"
arquivos = os.listdir("pages")

for i in range(1, 200):
    # Tenta o nome padrão
    nome1 = f"pagina_zapimoveis_page{i}.html"
    caminho1 = os.path.join("pages", nome1)

    if nome1 in arquivos:
        all_links_extends.extend(extrair_links_do_arquivo(caminho1))
        continue

    # Tenta o nome alternativo
    nome2 = f"pagina_zapimoveis{i - 1}_page{i}.html"
    caminho2 = os.path.join("pages", nome2)

    if nome2 in arquivos:
        all_links_extends.extend(extrair_links_do_arquivo(caminho2))
    else:
        print(f"Nenhum arquivo encontrado para página {i}, pulando...")

# Resultado final em all_links_extends
print(f"Total de links extraídos: {len(all_links_extends)}")


all_links_extends = list(set(all_links_extends))

links_limpos = [link.rstrip("\\/") for link in all_links_extends]

all_links_extends_limpo = [(link,) for link in links_limpos]

query = """
INSERT INTO amostras (link)
VALUES (%s)
ON DUPLICATE KEY UPDATE link = VALUES(link)
"""

print(len(all_links_extends_limpo))

cursor.executemany(query, all_links_extends_limpo)
connection.commit()

query = """
SELECT link FROM amostras
"""

cursor.execute(query)
resultados = cursor.fetchall()

for link in resultados:
    print(link)
print(len(resultados))