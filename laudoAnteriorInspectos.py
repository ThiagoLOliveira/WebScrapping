import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

import pyautogui as pg
import dotenv
import os
import time

dotenv.load_dotenv()

login_usuario_inspectos = os.getenv('INSPECTOS_EMAIL')
login_senha_inspectos = os.getenv('INSPECTOS_SENHA')

login_usuario_uono_relatorio = os.getenv('WEBMAIL_RELATORIOS_LOGIN')
login_senha_uono_relatorio = os.getenv('WEBMAIL_RELATORIOS_SENHA')

# options = webdriver.ChromeOptions()
# options.add_argument('--headless')  # Executa o Chrome em segundo plano (sem interface gráfica)
db_inspectos = pd.read_excel(r'databases\INSPECTOS.xls')

download_dir = r"M:\Thiago\Laudo_Inspectos"

# Configura opções do Chrome
chrome_options = Options()
chrome_prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", chrome_prefs)

# print(db_inspectos)

def acessa_email():
    """_Captura o codigo da B3 no email 'webmail'_

    Args:
        driver (_str_): _Driver do navegador_

    Returns:
        _num_: _retorna o capturado do email_
    """
    try:
        driver = webdriver.Chrome()
        driver.get('https://webmail.uonosanchez.com.br/?_task=mail&_mbox=INBOX')
        wait = WebDriverWait(driver, 100)

        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="rcmloginuser"]'))).send_keys(login_usuario_uono_relatorio)
        driver.find_element(By.XPATH, '//*[@id="rcmloginpwd"]').send_keys(login_senha_uono_relatorio)
        driver.find_element(By.XPATH, '//*[@id="rcmloginsubmit"]').click()

        time.sleep(14)
        # Atualiza a lista de emails
        driver.find_element(By.XPATH, '//*[@id="rcmbtn108"]').click()
        driver.refresh()

        # Acessa o primeiro email
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[4]/table[2]/tbody/tr[1]/td[2]/span[4]'))).click()
        time.sleep(4.5)

        # Switch para o iframe e captura o código
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="messagecontframe"]')))
        elements = driver.find_elements(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div/center/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td/b')
        if elements:
            codigo = elements[0].text
        else:
            codigo = None

        driver.switch_to.default_content()
        driver.quit()

        return codigo
    
    except:
        print(f"Não foi possivel baixar os Laudos INSPECTOS")

def realiza_login(driver, wait, usuario, senha):
    """Realiza o login completo no site da Inspectos, incluindo o envio do código recebido por e-mail."""
    try:
        driver.get('https://inspectos.com/sistema/index.html#/home')
        driver.maximize_window()
        
        # Preenche usuário e senha
        wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[3]/form/div[1]/div/div[1]/div/div[1]/div[2]/input'))).send_keys(usuario)
        wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[3]/form/div[1]/div/div[1]/div/div[2]/div[2]/input'))).send_keys(senha)
        
        # Clica em "Entrar"
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/form/div[1]/div/div[2]/div/button').click()
        
        # Envia código para e-mail
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[4]/div[2]/div/div/div/div[2]/div/a[1]'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[4]/div[2]/div/div/div/div[2]/button'))).click()
        
        # Acessa e insere o código
        codigo = acessa_email()
        if len(codigo) < 4:
            raise ValueError("Código recebido é muito curto.")
        
        campos_codigo = [
            '//*[@id="first"]',
            '//*[@id="second"]',
            '//*[@id="third"]',
            '//*[@id="fourth"]'
        ]
        
        for xpath, digito in zip(campos_codigo, codigo[:4]):
            campo = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            campo.clear()
            campo.send_keys(digito)
        
        # Envia o formulário de verificação
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[4]/div[2]/div/div/div/div[2]/form/button'))).click()
    
    except Exception as e:
        print(f"Erro ao realizar login: {e}")
        raise


def busca_inspectos():
    """_Buscador de arquivos no site do Santander_

    Sem necessidade de argumentos.'

    Entra no site, faz todos os acessos sozinho e conclui baixando os arquivos de Excel disponivel
    """
    i = 0  # Índice inicial

    while i < len(db_inspectos):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 10)
            realiza_login(driver, wait, login_usuario_inspectos, login_senha_inspectos)

            while i < len(db_inspectos):
                row = db_inspectos.iloc[i]
                identificador = row['Identificador']
                print(f"Processando: {identificador}")

                try:
                    input_codigo = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/form/div/div[1]/div[1]/div[2]/input')))
                    input_codigo.clear()
                    input_codigo.send_keys(row['Identificador'])
                    time.sleep(3)
                    submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/form/div/div[2]/div[1]/div/div[2]/button[2]')))
                    submit_button.click()
                    time.sleep(3)
                    grid = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div[2]/div/div/div/div/div[3]/div/div/ul/li[2]/a')))
                    grid.click()
                    time.sleep(3)
                    proposta = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idDivRaizGrid"]/div/table/tbody/tr')))
                    proposta.click()
                    time.sleep(3)
                    proposta_laudo = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div[1]/scrollable-tabset/div/div[1]/div/ul/li[7]/a')))
                    proposta_laudo.click()
                    time.sleep(3)
                    proposta_laudo_i = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div[2]/div/div[2]/div/div/table/tbody/tr[1]/td[5]/i')))
                    proposta_laudo_i.click()
                    time.sleep(3)
                    proposta_laudo_download = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div[2]/div/div[2]/div/div/table/tbody/tr[1]/td[5]/div[1]/div[2]/div/div/div[2]')))
                    proposta_laudo_download.click()
                    time.sleep(3)
                    print('Irei Seleionar o download')
                    button_location = pg.locateOnScreen('img/inspectos.png', confidence=0.8)
                    if button_location:
                        pg.click(pg.center(button_location))
                    else:
                        print("Botão não encontrado na tela.")
                    time.sleep(3)
                    bt_download = pg.locateOnScreen('img/download_inspectos.png', confidence=0.9)
                    if bt_download:
                        pg.click(pg.center(bt_download))
                    else:
                        print("Botão de download não encontrado na tela.")
                    time.sleep(3)
                    proposta_exit = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div[1]/i')))
                    proposta_exit.click()
                    time.sleep(3)
                    i += 1
                except Exception as e:
                    driver.quit()
                    break

        except Exception as e:
            print(f"Ocorreu um erro {e}")

        finally:
            driver.quit()

if __name__ == "__main__":
    busca_inspectos()