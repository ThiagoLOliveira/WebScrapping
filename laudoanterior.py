import pyautogui
import pandas as pd
import os
from pyautogui import ImageNotFoundException
import time

def wait_for_image(image_path, timeout=30, confidence=0.9):
    """
    Espera até a imagem aparecer na tela ou até o timeout acabar.
    """
    start_time = time.time()
    while True:
        location = pyautogui.locateOnScreen(image_path, confidence=confidence)
        if location:
            return location  # Encontrou a imagem
        if time.time() - start_time > timeout:
            raise TimeoutError(f"Imagem '{image_path}' não encontrada após {timeout} segundos.")
        time.sleep(0.5)  # Pequena pausa para não sobrecarregar o CPU
        

db = pd.read_excel('databases//download.xlsx')
# print(db)

lista = []
for _, row  in db.iterrows():
    lista.append(row['DocID'])

for docid in lista:
    pyautogui.hotkey('esc')
    pesquisa = wait_for_image('cetip\\pesquisar.png')
    
    pyautogui.click(pyautogui.center(pesquisa))

    pyautogui.write(str(docid))
    
    buscar = wait_for_image('cetip\\buscar.png')
    pyautogui.click(pyautogui.center(buscar))
    time.sleep(9)
    try:
        docid_pos = pyautogui.locateOnScreen('cetip\\docid.png', confidence=0.9)
        if docid_pos:
            x_click = docid_pos.left + 20 
            y_click = docid_pos.top + 58
            pyautogui.moveTo(x_click, y_click)
            pyautogui.click()
            print('cliquei no docid')
        pyautogui.hotkey('f5')
        time.sleep(9)
        laudo = wait_for_image('cetip\\laudo_mais_recente.png')
        pyautogui.click(pyautogui.center(laudo))
        print('cliquei no laudo mais recente')
        try:
            time.sleep(10)
            download = pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('cetip\\download.png', confidence=0.9)))
            print('cliquei no download')
        except ImageNotFoundException as e:
            time.sleep(3)
            download = pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('cetip\\download2.png', confidence=0.9)))
            print('cliquei no download')
        time.sleep(4)
        pyautogui.hotkey('enter')
        print('salvei o laudo')
        time.sleep(4)
        pyautogui.hotkey('ctrl', 'w')
        print('fechei a pagina')
        time.sleep(4)
        pyautogui.hotkey('esc')
        print('cliquei no esc')
        time.sleep(4)
        pyautogui.hotkey('ctrl', 'w')
        print('fechei a aba')
        time.sleep(5)
        pyautogui.hotkey('ctrl', 'tab')
        time.sleep(3)
    except ImageNotFoundException as e:
        print('Erro ao localizar o laudo:', e)
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(5)
        continue
