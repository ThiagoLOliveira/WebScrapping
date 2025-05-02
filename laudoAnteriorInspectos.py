import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import dotenv
import os
import time

dotenv.load_dotenv()

login_usuario_inspectos = os.getenv('INSPECTOS_EMAIL')
login_senha_inspectos = os.getenv('INSPECTOS_SENHA')

login_usuario_uono_relatorio = os.getenv('WEBMAIL_RELATORIOS_LOGIN')
login_senha_uono_relatorio = os.getenv('WEBMAIL_RELATORIOS_SENHA')

options = webdriver.ChromeOptions()

db_inspectos = pd.read_excel(r'databases\INSPECTOS.xls')

# print(db_inspectos)

def acessa_email():
    """_Captura o codigo da B4 no email 'webmail'_

    Args:
        driver (_str_): _Driver do navegador_

    Returns:
        _num_: _retorna o capturado do email_
    """
    try:
        driver = webdriver.Chrome(options=options)
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


def busca_inspectos():
    """_Buscador de arquivos no site do Santander_

    Sem necessidade de argumentos.'

    Entra no site, faz todos os acessos sozinho e conclui baixando os arquivos de Excel disponivel
    """
    try:
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 10) 
        driver.get('https://inspectos.com/sistema/index.html#/home')

        elemento_input_login = wait.until(EC.visibility_of_element_located((By.NAME, 'email'))).send_keys(login_usuario_inspectos)
        elemento_input_senha = wait.until(EC.visibility_of_element_located((By.NAME, 'senha'))).send_keys(login_senha_inspectos)
        elemento_botao_enviar = driver.find_element(By.ID, 'enter').click()
        enviar_codigo = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[4]/div[2]/div/div/div/div[2]/div/a[1]'))).click()
        button_avanc = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[4]/div[2]/div/div/div/div[2]/button'))).click()
        codigo = acessa_email()
        print(codigo)
        if len(codigo) >= 4:
            codigos_ = {
                '//*[@id="first"]': codigo[0],
                '//*[@id="second"]': codigo[1],
                '//*[@id="third"]': codigo[2],
                '//*[@id="fourth"]': codigo[3]
            }
        else:
            raise ValueError("Código recebido é muito curto.")
        
        for xpath, value in codigos_.items():
            try:
                elemento_input_codigo = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
                elemento_input_codigo.clear()  # Limpa o campo antes de inserir o valor
                elemento_input_codigo.send_keys(value)
            except Exception as e:
                print(f"Erro ao processar o elemento {xpath}: {e}")
        enviar = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[4]/div[2]/div/div/div/div[2]/form/button')))
        enviar.click()
        
        for i, row in db_inspectos.iterrows():
            input_codigo = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/form/div/div[1]/div[1]/div[2]/input')))
            input_codigo.clear()
            input_codigo.send_keys(row['Identificador'])
            time.sleep(1)
            submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/form/div/div[2]/div[1]/div/div[2]/button[2]')))
            submit_button.click()
            time.sleep(1)
            grid = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div[2]/div/div/div/div/div[3]/div/div/ul/li[2]/a')))
            grid.click()
            time.sleep(1)
            
    except Exception as e:
        print(f"Ocorreu um erro {e}")

    finally:
        driver.quit()


if __name__ == "__main__":
    busca_inspectos()