import pyautogui
import pandas as pd
import os
from time import sleep
# pyautogui.moveTo(130, 439)
db = pd.read_excel('databases//download.xlsx')
# print(db)

lista = []
for _, row  in db.iterrows():
    lista.append(row['DocID'])

sleep(2)
for docid in lista:
    pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('cetip\\pesquisar.png', confidence=0.8)))

    pyautogui.write(docid)

    pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('cetip\\buscar.png', confidence=0.8)))

    sleep(7)
    
    try:
        docid_pos = pyautogui.locateOnScreen('cetip\\docid.png', confidence=0.9)
        if docid_pos:
            x_click = docid_pos.left + 20 
            y_click = docid_pos.top + 58

            pyautogui.moveTo(x_click, y_click)
            pyautogui.click()
            sleep(7)
            
            try:
                arquivos = pyautogui.locateOnScreen('cetip\\arquivos.png', confidence=0.8)
            except Exception as e:
                x_atual, y_atual = pyautogui.position()
                novo_y = y_atual + 18
                pyautogui.moveTo(x_atual, novo_y)
                pyautogui.click()
                sleep(7)
                
    except Exception as e:
        print('Erro ao localizar o DocID:', e)
        if pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('cetip\\arquivos.png', confidence=0.8))):
            x_atual, y_atual = pyautogui.position()

            novo_y = y_atual + 12
            pyautogui.moveTo(x_atual, novo_y)
            pyautogui.click()

    # pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('cetip\\arquivos.png', confidence=0.8)))

    sleep(7)

    try:
        laudo = pyautogui.locateOnScreen('cetip\\laudo_mais_recente.png', confidence=0.9)
        # laudos = list(laudo)
        # laudos_data = []
        # laudos_data.append({
        #     'proposta': docid,
        #     'qtd': len(laudos),
        # })

        # df = pd.DataFrame(laudos_data)
        # filename = 'databases/laudos.xlsx'

        # if os.path.exists(filename):
        #     df_existing = pd.read_excel(filename)
        #     df_total = pd.concat([df_existing, df], ignore_index=True)
        # else:
        #     df_total = df

        # df_total.to_excel(filename, index=False)
        sleep(3)
        pyautogui.click(pyautogui.center(laudo))
        sleep(5)
        pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('cetip\\download.png', confidence=0.9)))
        sleep(4)
        pyautogui.hotkey('enter')
        sleep(4)
        pyautogui.hotkey('ctrl', 'w')
        sleep(4)
        pyautogui.hotkey('esc')
        sleep(4)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except Exception as e:
        print('Erro ao localizar o laudo:', e)
        sleep(5)
        continue