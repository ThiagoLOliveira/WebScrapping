import pytesseract
import pyautogui
from PIL import Image
import datetime

# 1. Tirar screenshot da tela
screenshot = pyautogui.screenshot()
data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)

# 2. Procurar por todas as ocorrências de "Laudo"
laudos = []
for i, word in enumerate(data['text']):
    if word.lower() == "laudo":
        # Verifica se a palavra "Laudo" está associada a uma data próxima
        for j in range(i, min(i+20, len(data['text']))):  # verifica nas próximas palavras
            try:
                # Tenta converter uma data no formato dd/mm/yyyy
                possible_date = data['text'][j]
                if "/" in possible_date:
                    date_obj = datetime.datetime.strptime(possible_date, "%d/%m/%Y")
                    laudos.append({
                        "date": date_obj,
                        "x": data['left'][j],
                        "y": data['top'][j]
                    })
                    break
            except:
                continue

# 3. Encontrar o laudo com data mais recente
if laudos:
    laudo_recente = max(laudos, key=lambda x: x['date'])
    
    # 4. Ajustar posição do clique (por exemplo: botão de download está uns 600px à direita da data)
    pyautogui.moveTo(laudo_recente['x'] + 600, laudo_recente['y'] + 10)
    pyautogui.click()
else:
    print("Nenhum laudo encontrado.")
