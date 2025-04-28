import shutil
import os
import pandas as pd
import dotenv
# import holidays.countries
import numpy as np
from pandas.tseries.offsets import CustomBusinessDay
from glob import glob
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from datetime import date, timedelta
from os import listdir
from os.path import isfile, join, basename
from pathlib import Path
from time import sleep
from glob import glob
import re
import requests
from pandas.tseries.offsets import BDay
from mysql import connector
from dateutil.parser import parse

dotenv.load_dotenv()

options = webdriver.ChromeOptions()
options.add_argument("--headless")

login_usuario_bradesco = os.getenv('LOGIN_BRADESCO')
login_senha_bradesco = os.getenv('SENHA_BRADESCO')

login_usuario_uono_relatorio = os.getenv('WEBMAIL_RELATORIOS_LOGIN')
login_senha_uono_relatorio = os.getenv('WEBMAIL_RELATORIOS_SENHA')

login_usuario_inspectos = os.getenv('LOGIN_INSPECTOS')
login_senha_inspectos = os.getenv('SENHA_INSPECTOS')

login_usuario_uono = os.getenv('WEBMAIL_UONO_LOGIN')
login_senha_uono = os.getenv('WEBMAIL_UONO_SENHA')

login_usuario_viva = os.getenv('LOGIN_VIVA')
login_senha_viva = os.getenv('SENHA_VIVA')

host_db = os.getenv('host')
user_db = os.getenv('user')
password_db = os.getenv('password')
database_db = os.getenv('database')

# Conexão com o banco de dados
connection = connector.connect(
    host=host_db,
    user=user_db,
    password=password_db,
    database=database_db
)
cursor = connection.cursor()

# Lista de feriados fixos
FERIADOS = [
    '2025-01-01', '2025-03-03', '2025-03-04', '2025-04-18', '2025-04-21', '2025-05-01',
    '2025-06-19', '2025-09-07', '2025-10-12', '2025-11-02', '2025-11-15', '2025-11-20', '2025-12-25'
]

FERIADOS = [datetime.strptime(f, "%Y-%m-%d").date() for f in FERIADOS]
user_dir = os.path.expanduser('~')
usuario = os.path.join(user_dir, 'Downloads')


def baixar_planilha(url, max_tentativas=3):
    for tentativa in range(max_tentativas):
        try:
            resposta = requests.get(url, stream=True, timeout=20)
            resposta.raise_for_status()  # Gera erro se a resposta não for 200
            with open("planilha.xlsx", "wb") as f:
                for chunk in resposta.iter_content(chunk_size=8192):
                    f.write(chunk)
            print("Download concluído!")
            return "planilha.xlsx"
        except requests.exceptions.RequestException as e:
            print(f"Erro ao baixar ({tentativa+1}/{max_tentativas}): {e}")
            sleep(2)  # Espera 2 segundos antes de tentar novamente
    raise Exception("Falha ao baixar a planilha após várias tentativas")


def limpar_coluna_data_apr(coluna):
    """Versão aprimorada com tratamento especial para casos do Bradesco"""
    
    def parse_flexivel(data_str):
        try:
            # Primeiro tenta o parser universal
            return parse(data_str, dayfirst=True)
        except:
            # Fallback para tratamento manual
            data_str = re.sub(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', r'\1/\2/\3', data_str)  # Normaliza separadores
            data_str = re.sub(r'[;.,](\d{2})$', r':\1', data_str)  # Horas com separadores errados
            data_str = re.sub(r'\s+', ' ', data_str).strip()  # Espaços múltiplos
            
            # Tenta formatos comuns BR
            formatos = [
                '%d/%m/%Y %H:%M',
                '%d/%m/%Y',
                '%d/%m/%y %H:%M',
                '%d/%m/%y',
                '%Y-%m-%d %H:%M:%S',
                '%Y-%m-%d'
            ]
            
            for fmt in formatos:
                try:
                    return datetime.strptime(data_str, fmt)
                except:
                    continue
            
            return pd.NaT
        
    return coluna.apply(lambda x: parse_flexivel(str(x)) if pd.notna(x) else pd.NaT)


def mover_arquivos(caminho, nome_do_arquivo):
    try:
        if caminho:
            converte_em_excel(caminho, f'{nome_do_arquivo}')
            os.unlink(caminho)
            move(usuario, 'P:\\Planilhas_dash')
    except Exception as e:
        print('Ocorreu um erro ao mover os arquivos:', e)

#Funções de Busca de Laudos
def busca_bradesco():
    """_Buscador de arquivos no site do Bradesco_

    Sem necessidade de argumentos.

    Entra no site, faz todos os acessos sozinho e conclui baixando os arquivos de Excel disponivel
    """
    try:

        driver = webdriver.Chrome(options=options)
        driver.get('https://avaliacaobra.com.br/')
        # driver.minimize_window()
        
        elemento_botao = driver.find_element(By.ID, 'btnFornec')
        elemento_botao.click()

        elemento_input_login = driver.find_element(By.ID, 'txtUsuario')
        print(login_usuario_bradesco)
        elemento_input_login.send_keys(login_usuario_bradesco)

        elemento_input_senha = driver.find_element(By.ID, 'txtSenha')
        print(login_senha_bradesco)
        elemento_input_senha.send_keys(login_senha_bradesco)

        elemento_botao_enviar = driver.find_element(By.ID, 'btnEnviar')
        elemento_botao_enviar.click()


        elemento_hamb = driver.find_element(By.XPATH, '//*[@id="Panel1"]/ul/li[3]/a')
        elemento_hamb.click()

        sleep(1)

        elemento_laudo = driver.find_element(By.XPATH, '//*[@id="mm-0"]/ul/li[1]')
        elemento_laudo.click()

        sleep(15)

        hamb_rev = driver.find_element(By.XPATH, '//*[@id="my-icon"]')
        hamb_rev.click()
        

        revisao = driver.find_element(By.XPATH, '//*[@id="mm-0"]/ul/li[5]')
        revisao.click()
        
        sleep(2)
        
        frame = driver.find_element(By.XPATH, '//*[@id="ifrm"]')
        driver.switch_to.frame(frame)
        
        data_inicio = date.today() - timedelta(days=40)
        
        input_data_inicio = driver.find_element(By.XPATH, '//*[@id="Txtinicio"]')
        input_data_inicio.send_keys(data_inicio.strftime('%d/%m/%Y'))
        
        input_data_fim = driver.find_element(By.XPATH, '//*[@id="txttermino"]')
        input_data_fim.send_keys(date.today().strftime('%d/%m/%Y'))
        
        bt_buscar = driver.find_element(By.XPATH, '//*[@id="BlnProcessar"]').click()
        sleep(2.5)
        
        baixar_xlsx = driver.find_element(By.XPATH, '//*[@id="btnExcel"]').click()
        sleep(4)
        
    except:
        print("Não foi possivel baixar os arquivos do BRADESCO")


def busca_bradesco_concl():
    """_Buscador de arquivos no site do Bradesco_

    Sem necessidade de argumentos.

    Entra no site, faz todos os acessos sozinho e conclui baixando os arquivos de Excel disponivel
    """
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        driver = webdriver.Chrome(options=options)
        driver.get('https://avaliacaobra.com.br/')
        # driver.minimize_window()
        
        elemento_botao = driver.find_element(By.ID, 'btnFornec')
        elemento_botao.click()

        elemento_input_login = driver.find_element(By.ID, 'txtUsuario')
        print(login_usuario_bradesco)
        elemento_input_login.send_keys(login_usuario_bradesco)
        

        elemento_input_senha = driver.find_element(By.ID, 'txtSenha')
        print(login_senha_bradesco)
        elemento_input_senha.send_keys(login_senha_bradesco)

        elemento_botao_enviar = driver.find_element(By.ID, 'btnEnviar')
        elemento_botao_enviar.click()


        elemento_hamb = driver.find_element(By.XPATH, '//*[@id="Panel1"]/ul/li[3]/a')
        elemento_hamb.click()

        sleep(1)

        elemento_hamb_concluidos = driver.find_element(By.XPATH, '//*[@id="mm-0"]/ul/li[2]')
        elemento_hamb_concluidos.click()

        sleep(5)

        iframe = driver.find_element(By.XPATH, '//*[@id="Cont"]')
        driver.switch_to.frame(iframe)
        element = driver.find_element(By.XPATH, '//*[@id="TextBox4"]')
        element.send_keys(Keys.CONTROL, 'a')
        element.send_keys(Keys.ENTER)

        sleep(3)
        # Obtém o diretório de downloads do usuário atual
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")

        # Nome do arquivo que queremos localizar
        nome_arquivo = "Concluidos"

        # Percorre a pasta Downloads para encontrar o arquivo desejado
        encontrado = False
        for arquivo in os.listdir(downloads_dir):
            caminho_completo = os.path.join(downloads_dir, arquivo)
            print(f"Arquivo encontrado: {caminho_completo}")
            encontrado = True
            move(caminho_completo, 'P:\\Planilhas_dash')
            break
                
        if not encontrado:
            print(f"Arquivo '{nome_arquivo}' não encontrado na pasta Downloads.")

        sleep(15)
    
    except:
        print("Não foi possivel baixar os arquivos do BRADESCO concluidos")


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

        sleep(14)
        # Atualiza a lista de emails
        driver.find_element(By.XPATH, '//*[@id="rcmbtn108"]').click()
        driver.refresh()

        # Acessa o primeiro email
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[4]/table[2]/tbody/tr[1]/td[2]/span[4]'))).click()
        sleep(4.5)

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

    Sem necessidade de argumentos.

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
        elemento_hamb = wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='menu-principal']/button"))).click()
        elemento_botao = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/uib-accordion/div/div/div[1]/h4/a/span/div'))).click()
        elemento_botao_prox = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/uib-accordion/div/div/div[2]/div/div/div/div[1]/h4/a'))).click()
        elemento_botao_analitico = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/uib-accordion/div/div/div[2]/div/div/div/div[2]/div/div/div[1]/div[1]/h4/a/span/div'))).click()
        data_hoje = date.today()

        data_anterior = data_hoje - timedelta(days=40)
        data_mes_anterior = data_hoje - timedelta(days=40)

        data_formatada = data_anterior.strftime('%d/%m/%Y')
        elemento_input_data_inicio = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="idFiltroDataPeriodoInicio"]/datepicker[1]/input')))

        sleep(4.5)
        elemento_input_data_inicio.send_keys(Keys.CONTROL, 'a')
        sleep(0.5)
        elemento_input_data_inicio.send_keys('01/01/2025')
        sleep(5)

        elemento_input_botao = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div/div[2]/div[2]/div[2]/div[5]/button[2]').click()
        sleep(5.5)

        elemento_input_botao_excel = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div/div[2]/div[2]/div[2]/div[5]/button[1]').click()
        sleep(7)

    except Exception as e:
        print(f"Ocorreu um erro {e}")

    finally:
        driver.quit()


def busca_vivaintra(option):
    """_Buscador de Laudos no site 'VIVAINTRA'
    Busca os 4 tipos de Laudos, e busca cada parte individualmente
    """
    try:
        driver = webdriver.Chrome(options=options)
        
        data_hoje = date.today() + timedelta(days=1)
        data_inicio_mes = data_hoje.replace(day=1)
        data_anterior = data_inicio_mes - timedelta(days=3)

        wait = WebDriverWait(driver, 400)
        driver.get("https://uonosanchez.vivaintra.com/admin-blog")
        
        # Entra na conta
        elem_email_vivaintra = wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='username']")))
        elem_email_vivaintra.send_keys('contato.thiagolima12@gmail.com')
        sleep(1)
        elem_email_vivaintra.send_keys(Keys.ENTER)
        
        elem_senha = driver.find_element(By.NAME, "password")
        elem_senha.send_keys('27310701')
        sleep(1)
        elem_senha.send_keys(Keys.ENTER)

        
        elem_seleciona = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="intranet-login"]/form/input'
        )))
        elem_seleciona.send_keys('27310701')
        sleep(1)
        elem_seleciona.send_keys(Keys.ENTER)

        elem_seleciona_produtividade = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/nav/div[1]/div[1]/a[2]/i'
        ))).click()
        sleep(1)
        elem_seleciona_requisicoes = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/div[1]/div[2]/div[2]/div/div[2]/div/div[10]/a' # Aqui ele entra nas requisições
        ))).click()
        sleep(1)
        elem_seleciona_toda_as_requisicoes = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[1]/div[2]/ul/li[1]/a'
        ))).click()
        sleep(1)

        seleciona_bradesco = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            f'//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[2]/div/select/option[{option}]'
        ))).click()

        seleciona_bradesco_data = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[11]/div[1]/input'
        ))).send_keys(data_anterior.strftime('%d/%m/%Y'))

        seleciona_bradesco_data_segundo_input = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[11]/div[3]/div/input'
        ))).send_keys(data_hoje.strftime('%d/%m/%Y'))
        sleep(1)
        seleciona_bradesco_data_segundo_input_botao = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[13]/div/button'
        ))).send_keys(Keys.ENTER)
        sleep(1)
        seleciona_bradesco_data_segundo_input_botao_cria_excel = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/nav/div[3]/a'
        ))).click()

        seleciona_bradesco_data_segundo_input_botao_baixa_excel = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[2]/div[1]/a' 
        ))).click()

        sleep(5)

    except Exception as e:
        print(f'Ocorreu o seguinte erro {e}')


    driver.quit()


def busca_vivaintra_conferencia_laudo(esteira):
    """_Buscador de Laudos no site 'VIVAINTRA'
    Busca os 4 tipos de Laudos, e busca cada parte individualmente
    """
    if esteira == 4:
        select = 2
    elif esteira == 1:
        select = 6
    else:
        select = 3
        
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        driver = webdriver.Chrome(options=options)
        
        
        data_hoje = date.today() + timedelta(days=1)
        data_inicio_mes = data_hoje.replace(day=1)
        data_anterior = data_inicio_mes - timedelta(days=40)

        wait = WebDriverWait(driver, 400)
        driver.get("https://uonosanchez.vivaintra.com/admin-blog")
        
        # Entra na conta
        elem_email_vivaintra = wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='username']")))
        elem_email_vivaintra.send_keys('contato.thiagolima12@gmail.com')
        sleep(1)
        elem_email_vivaintra.send_keys(Keys.ENTER)
        sleep(1)
        elem_senha = driver.find_element(By.NAME, "password")
        elem_senha.send_keys('27310701')
        sleep(1)
        elem_senha.send_keys(Keys.ENTER)

        elem_seleciona = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="intranet-login"]/form/input'
        )))
        elem_seleciona.send_keys('27310701')
        sleep(1)
        elem_seleciona.send_keys(Keys.ENTER)

        elem_seleciona_produtividade = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/nav/div[1]/div[1]/a[2]/i'
        ))).click()
        sleep(1)
        elem_seleciona_requisicoes = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/div[1]/div[2]/div[2]/div/div[2]/div/div[10]/a' # Aqui ele entra nas requisições
        ))).click()
        sleep(1)
        selec_esteira = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            f'//*[@id="body-content"]/div[2]/div[2]/div/div[2]/div[5]/table/tbody/tr[{esteira}]/td[4]/div/button'
        ))).click()
        sleep(1)
        relatorio_etapas = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            f'//*[@id="body-content"]/div[2]/div[2]/div/div[2]/div[5]/table/tbody/tr[{esteira}]/td[4]/div/ul/li[5]'
        ))).click()
        
        conferencia_laudo = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            f'//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[3]/div/select/option[{select}]'
        ))).click()
        

        hoje = datetime.today()
        primeiro_dia = hoje.replace(day=1) - timedelta(days=30)
        sleep(1)
        amanha = date.today() + timedelta(days=1)
        amanha_formatado = amanha.strftime('%d/%m/%Y')
        sleep(1)
        data_inicio = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[6]/div[1]/input'
        ))).send_keys(primeiro_dia.strftime('%d/%m/%Y'))
        
        amanha = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[6]/div[3]/div/input'
        ))).send_keys(amanha_formatado)
        sleep(1)
        btt_busca = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/div[1]/div[2]/div[1]/div[4]/div/div[2]/form/div[9]/div/button'
        ))).send_keys(Keys.ENTER)
        
        sleep(5)
        
        btt_excel = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/nav/div[3]/a'
        ))).click()
        
        btt_excel_baixar = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[2]/div[1]/a'
        ))).click()
        
        sleep(5)
        

    except Exception as e:
        print(f'Ocorreu o seguinte erro {e}')

#Funções de manipulação de arquivo
def move(path_origem, path_destino):
    """Move o arquivo da pasta de origem (path_origem) para a pasta de destino (path_destino)

    Args:
        path_origem (_str_): _Pasta onde o arquivo está localizado_
        path_destino (_str_): _Pasta onde o arquivo será encaminhado_
    """
    try:
        for item in [join(path_origem, f) for f in listdir(path_origem) if isfile(join(path_origem, f))]:
            #print(item)
            shutil.move(item, join(path_destino, basename(item)))
            print(f'Arquivo(s) movido de "{item}" para --> "{join(path_destino, basename(item))}"')
    except Exception as e:
        print(f'Ocorreu o seguinte erro ao tentar mover o arquivo {e}')


def converte_em_excel(caminho, nome_saida=None):
    """
    Converte um arquivo CSV ou Excel (presumindo XLSX para Excel) para o formato Excel.
    
    Argumentos:
        caminho (str ou Path): Caminho para o arquivo de entrada.
        
    Retorno:
        Nenhum
    """
    try:
        caminho = str(caminho)
        if caminho.endswith('.csv'):
            # Se for um arquivo CSV
            read_file = pd.read_csv(caminho, sep=';') 
        elif caminho.endswith('.xls') or caminho.endswith('.xlsx'):
            # Se for um arquivo Excel (XLS ou XLSX)
            read_file = pd.read_excel(caminho, engine='xlrd' if caminho.endswith('.xls') else 'openpyxl')
        else:
            print(f'Tipo de arquivo não suportado: {caminho}')
            return
    except Exception as e:
        print(f'Erro ao ler o arquivo {caminho}: {e}')
        return

    if nome_saida:
        caminho_saida = os.path.join(os.path.dirname(caminho), nome_saida + ".xlsx")
    else:
        if caminho.endswith('.csv'):
            caminho_saida = caminho.replace(".csv", ".xlsx")
        else:
            caminho_saida = caminho.replace(".xls", ".xlsx")

    try:
        read_file.to_excel(caminho_saida, index=None, header=True)
        print(f'Arquivo convertido com sucesso: {caminho_saida}')
    except Exception as e:
        print(f'Erro ao converter o arquivo {caminho} para Excel: {e}')


def encontrar_arquivo_mais_recente(nome_inicio_arquivo, pasta, extensao):
    """_Na pasta (pasta) buscara o arquivo mais recente_

    Args:
        nome_inicio_arquivo (_str_): _Nome do arquivo_
        pasta (_str_): _Local onde o arquivo está alocado_
        extensao (_str_): _Extensão do arquivo, Ex: .csv, .xlsx, .pdb, etc_

    Returns:
        _str_: _Retorna o caminho do arquivo mais recente_
    """
    try:
        lista_arquivos = Path(pasta).glob(f"{nome_inicio_arquivo}*.{extensao}")
        arquivo_mais_recente = max(lista_arquivos, key=os.path.getmtime, default=None)
        return arquivo_mais_recente
    except Exception as e:
        print(f"Ocorreu o seguinte erro {e}")

#Ajusta todos as planilhas para deixar no formato correto
def ajustar_horario(data):
    """Ajusta a data para respeitar o intervalo de trabalho das 8h às 18h."""
    if data.hour >= 18:
        return (data + BDay(1)).replace(hour=8, minute=0, second=0)
    elif data.hour < 8:
        return data.replace(hour=8, minute=0, second=0)
    return data


def ajustar_para_dia_util(data):
    """Se a data cair em um fim de semana ou feriado, ajusta para o próximo dia útil."""
    while data.weekday() >= 5 or data.date() in FERIADOS:  # 5=Sábado, 6=Domingo
        data += timedelta(days=1)
    return data


def calcular_horas_uteis(data_inicio, data_fim):
    """Calcula horas úteis (8h às 18h) entre duas datas, pulando fins de semana e feriados."""
    if pd.isna(data_inicio) or pd.isna(data_fim):
        return (0.0,0.0)

    data_inicio = ajustar_para_dia_util(data_inicio)
    data_fim = ajustar_para_dia_util(data_fim)

    # Caso estejam no mesmo dia útil
    if data_inicio.date() == data_fim.date():
        inicio = max(data_inicio, data_inicio.replace(hour=8))
        fim = min(data_fim, data_fim.replace(hour=18))
        total_horas = max((fim - inicio).total_seconds() / 3600, 0)
        return total_horas, total_horas / 10

    # Para dias diferentes
    total_horas = 0

    # Horas úteis do primeiro dia
    inicio = max(data_inicio, data_inicio.replace(hour=8))
    fim = data_inicio.replace(hour=18)
    if data_inicio.date() not in FERIADOS and data_inicio.weekday() < 5:
        total_horas += max((fim - inicio).total_seconds() / 3600, 0)

    # Horas úteis do último dia
    inicio = data_fim.replace(hour=8)
    fim = min(data_fim, data_fim.replace(hour=18))
    if data_fim.date() not in FERIADOS and data_fim.weekday() < 5:
        total_horas += max((fim - inicio).total_seconds() / 3600, 0)

    # Horas úteis dos dias intermediários
    atual = data_inicio + timedelta(days=1)
    while atual.date() < data_fim.date():
        if atual.weekday() < 5 and atual.date() not in FERIADOS:  # Apenas dias úteis e sem feriados
            total_horas += 10
        atual += timedelta(days=1)

    total_dias = total_horas / 10
    return total_horas, total_dias


def calcular_data_limite(row):
    feriados = pd.to_datetime([
    '2025-01-01', '2025-03-03', '2025-03-04', '2025-04-18', '2025-04-21', '2025-05-01',
    '2025-06-19', '2025-09-07', '2025-10-12', '2025-11-02', '2025-11-15', '2025-11-20', '2025-12-25'
    ])

    # Definição do calendário de dias úteis
    dias_uteis = CustomBusinessDay(holidays=feriados)

    # Definição de horário comercial
    HORA_INICIO = 8  # 08:00 da manhã
    HORA_FIM = 18  # 18:00 da tarde
    HORAS_UTEIS_ADICIONAR = 6
    data_recebimento = row["Dt. Recebimento Prest."]

    # Se o valor for nulo, retorna a própria Data Limite
    if pd.isna(data_recebimento):
        return row["Data Limite"]

    # Se recebeu em um feriado ou fim de semana, avança para o próximo dia útil às 08:00
    while data_recebimento.weekday() >= 5 or data_recebimento.normalize() in feriados:
        data_recebimento = (data_recebimento + dias_uteis).replace(hour=HORA_INICIO, minute=0, second=0)

    # Se recebido após as 18h, joga para o próximo dia útil às 08:00
    if data_recebimento.hour >= HORA_FIM:
        data_recebimento = (data_recebimento + dias_uteis).replace(hour=HORA_INICIO, minute=0, second=0)

    horas_restantes = HORAS_UTEIS_ADICIONAR
    while horas_restantes > 0:
        # Cálculo das horas disponíveis no dia atual
        hora_atual = data_recebimento.hour
        horas_disponiveis = HORA_FIM - hora_atual

        if horas_restantes <= horas_disponiveis:
            # Se as horas cabem no mesmo dia, apenas somamos
            data_recebimento += pd.Timedelta(hours=horas_restantes)
            break
        else:
            # Se as horas não cabem, avançamos para o próximo dia útil às 08:00
            horas_restantes -= horas_disponiveis
            data_recebimento = (data_recebimento + dias_uteis).replace(hour=HORA_INICIO, minute=0, second=0)

    return data_recebimento


def calcular_use_casa(row):
    data_vistoria = pd.to_datetime(row["Data Vistoria"], errors='coerce', dayfirst=True)
    

    # Se o valor for nulo, retorna a própria Data Limite
    if pd.isna(data_vistoria):
        return row["Data Limite"]

    data_vistoria = data_vistoria + BDay(1)

    return data_vistoria


def ajustar_inspectos():
    try:
        # Localiza o arquivo no diretório de Downloads
        user_dir = os.path.expanduser('~')
        arquivos = glob(os.path.join(user_dir, 'Downloads', '*.xls'))
        if not arquivos:
            print("Nenhum arquivo encontrado na pasta Downloads.")
            return

        print(arquivos)
        # Processa os arquivos encontrados
        for arquivo in arquivos:
            print(f'Processando o arquivo: {arquivo}')
            try:
                xls = pd.ExcelFile(arquivo)
                print(f'Abas encontradas: {xls.sheet_names}')

                # Processar as abas
                aba_dados = {}
                for aba in ['Crédito imobiliário', 'Renegociação', 'Garantias']:
                    if aba in xls.sheet_names:
                        print(f'Processando a aba: {aba}')
                        db = pd.read_excel(arquivo, sheet_name=aba)
                        colunas_principais = [
                            'Identificador', 'Tipo Inspeção', 'Nro. Proposta', 'Município',
                            'Data Limite', 'UF', 'Status', 'Tipo Imovel',
                            'Data Vistoria', 'Dt. Recebimento Prest.', 'Data Entrega 1° Laudo (B)','Data Agendamento','Horário Agendamento',
                            'Produto Comercial', 'Data Vistoria'
                        ]
                        # Mantém apenas as colunas necessárias
                        aba_dados[aba] = db[[col for col in colunas_principais if col in db.columns]]

            except Exception as e:
                print(f"Erro ao processar o arquivo '{arquivo}': {e}")
                continue

            # Combina todas as abas ajustadas em um único DataFrame
            db_concatenado = pd.concat(aba_dados.values(), ignore_index=True)
            os.makedirs('archives', exist_ok=True)
            db_concatenado.to_excel('archives\\inspectos_ajustada.xlsx', index=False)
            # Exemplo para teste
            db = pd.read_excel('archives\\inspectos_ajustada.xlsx')
            db['Dt. Recebimento Prest.'] = pd.to_datetime(db['Dt. Recebimento Prest.'], errors='coerce', dayfirst=True)
            db['Data Entrega 1° Laudo (B)'] = pd.to_datetime(db['Data Entrega 1° Laudo (B)'], errors='coerce', dayfirst=True)

            db["Horas Diferença"], db["Dias Diferença"] = zip(*db.apply(
                lambda row: calcular_horas_uteis(row["Dt. Recebimento Prest."], row["Data Entrega 1° Laudo (B)"]),
                axis=1
            ))


            # Definição de horário comercial
            HORA_INICIO = 8  # 08:00 da manhã
            HORA_FIM = 18  # 18:00 da tarde
            HORAS_UTEIS_ADICIONAR = 6

            # Aplicando a função ao DataFrame
            db["Data Limite"] = db.apply(
                lambda row: calcular_data_limite(row) if row["Tipo Inspeção"] == "AVM" else row["Data Limite"],
                axis=1
            )
            
            db["Data Limite"] = db.apply(
                lambda row: calcular_use_casa(row) if row["Produto Comercial"] == "Cred. Imob. PF - Crédito Pessoal (use casa)" else row["Data Limite"],
                axis=1
            )

            db["Nro. Proposta"] = db.apply(
                lambda row: row['Identificador'] if pd.isna(row["Nro. Proposta"]) else row["Nro. Proposta"],
                axis=1
            )

            # os.makedirs('archives', exist_ok=True)
            db.to_excel('P:\\Planilhas_dash\\inspectos_ajustada.xlsx', index=False)
            print("Arquivo processado e movido com sucesso.")

    except Exception as e:
        print(f"Erro ao ajustar a planilha: {e}")


def juntar_planilhas():
    planilhas = ['bradesco_viva.xlsx','presenciais.xlsx', 'itau_viva.xlsx', 'avm-viva.xlsx']
    caminho = 'P:\\Planilhas_dash\\'
    lista_db = []

    for planilha in planilhas:
        caminho_arquivo = join(caminho, planilha)
        db = pd.read_excel(caminho_arquivo)
        lista_db.append(db)

    db_concat = pd.concat(lista_db, ignore_index = True)

    db_concat.to_excel(join(caminho,'Planilha_mesclada.xlsx'), index=False)


def adicionar_horas_uteis(data_criacao, horas_a_adicionar, feriados):
    """Função para adicionar horas úteis a partir de uma data inicial"""
    hora_inicio_comercial = 8
    hora_fim_comercial = 18
    horas_por_dia = hora_fim_comercial - hora_inicio_comercial

    # Ajustar para o próximo dia útil às 8h se a data de criação for sábado, domingo ou feriado
    if data_criacao.weekday() >= 5 or data_criacao.date() in feriados:  # Sábado (5) ou domingo (6) ou feriado
        data_criacao = (data_criacao + pd.offsets.CustomBusinessDay(n=1, holidays=feriados)).normalize() + pd.Timedelta(hours=hora_inicio_comercial)

    # Ajustar para o próximo dia útil às 8h, se após as 18h
    elif data_criacao.hour >= hora_fim_comercial:
        data_criacao = (data_criacao + pd.offsets.CustomBusinessDay(n=1, holidays=feriados)).normalize() + pd.Timedelta(hours=hora_inicio_comercial)

    # Ajustar para dentro do horário comercial, se antes das 8h
    elif data_criacao.hour < hora_inicio_comercial:
        data_criacao = data_criacao.normalize() + pd.Timedelta(hours=hora_inicio_comercial)

    # Calcular o tempo restante no dia de criação
    horas_restantes_hoje = max(0, hora_fim_comercial - data_criacao.hour)

    # Se as horas a adicionar cabem no mesmo dia
    if horas_a_adicionar <= horas_restantes_hoje:
        return data_criacao + pd.Timedelta(hours=horas_a_adicionar)

    # Reduzir as horas e passar para o próximo dia útil
    horas_a_adicionar -= horas_restantes_hoje
    data_criacao = (data_criacao + pd.offsets.CustomBusinessDay(n=1, holidays=feriados)).normalize() + pd.Timedelta(hours=hora_inicio_comercial)

    # Adicionar horas úteis nos dias seguintes
    while horas_a_adicionar > 0:
        if horas_a_adicionar <= horas_por_dia:
            return data_criacao + pd.Timedelta(hours=horas_a_adicionar)

        # Subtrair um dia útil e passar para o próximo
        horas_a_adicionar -= horas_por_dia
        data_criacao = (data_criacao + pd.offsets.CustomBusinessDay(n=1, holidays=feriados)).normalize() + pd.Timedelta(hours=hora_inicio_comercial)

    return data_criacao


def vencimentos_itau():
    user_dir = os.path.expanduser('~')
    feriados_brasil = list(holidays.countries.Brazil(years=[2024, 2025]).keys())

    # Caminho do arquivo
    file_path = 'P:\\Planilhas_dash\\download.xlsx'
    db = pd.read_excel(file_path)
    db['Data de criação'] = pd.to_datetime(db['Data de criação'], dayfirst=True)
    db['Data da 1a entrega do laudo'] = pd.to_datetime(db['Data da 1a entrega do laudo'], dayfirst=True)
    db['Data de finalização'] = pd.to_datetime(db['Data de finalização'], dayfirst=True)
    
    # Somar 44 horas úteis
    db['Vencimento_Ajustado_Final'] = db['Data de criação'].apply(
        lambda x: adicionar_horas_uteis(x, 24, FERIADOS)
    )

    # Ajustar status
    status_de_para = {
        'Finalizado': 'Concluído',
        'Finalizado (Com Restrições)': 'Concluído',
        'Laudo Recebido': 'Concluído',
        'Laudo em Preenchimento': 'EmAberto',
        'Solicitação Aceita': 'EmAberto',
    }
    db.replace(status_de_para, inplace=True)

    db['Data de criação'] = pd.to_datetime(db['Data de criação'], errors='coerce', dayfirst=True)
    db['Data da 1a entrega do laudo'] = pd.to_datetime(db['Data da 1a entrega do laudo'], errors='coerce', dayfirst=True)
    db["Horas Diferença"], db["Dias Diferença"] = zip(*db.apply(
        lambda row: calcular_horas_uteis(row["Data de criação"], row["Data da 1a entrega do laudo"]),
        axis=1
    ))

    # Salvar arquivo atualizado
    output_path = os.path.join('P:\\Planilhas_dash\\download_atualizado.xlsx')
    db.to_excel(output_path, index=False)
    return f'Arquivo salvo em: {output_path}'


def vencimentos_itau_mensal():
    user_dir = os.path.expanduser('~')
    feriados_brasil = list(holidays.countries.Brazil(years=[2024, 2025]).keys())

    # Caminho do arquivo
    file_path = 'P:\\Planilhas_dash\\download_mensal.xlsx'
    db = pd.read_excel(file_path)
    db['Data de criação'] = pd.to_datetime(db['Data de criação'], dayfirst=True)
    db['Data da 1a entrega do laudo'] = pd.to_datetime(db['Data da 1a entrega do laudo'], dayfirst=True)
    db['Data de finalização'] = pd.to_datetime(db['Data de finalização'], dayfirst=True)

    # Somar 44 horas úteis
    db['Vencimento_Ajustado_Final'] = db['Data de criação'].apply(
        lambda x: adicionar_horas_uteis(x, 24, FERIADOS)
    )

    # Ajustar status
    status_de_para = {
        'Finalizado': 'Concluído',
        'Finalizado (Com Restrições)': 'Concluído',
        'Laudo Recebido': 'Concluído',
        'Laudo em Preenchimento': 'EmAberto',
        'Solicitação Aceita': 'EmAberto',
    }
    db.replace(status_de_para, inplace=True)

    db['Data de criação'] = pd.to_datetime(db['Data de criação'], errors='coerce', dayfirst=True)
    db['Data da 1a entrega do laudo'] = pd.to_datetime(db['Data da 1a entrega do laudo'], errors='coerce', dayfirst=True)
    db["Horas Diferença"], db["Dias Diferença"] = zip(*db.apply(
        lambda row: calcular_horas_uteis(row["Data de criação"], row["Data da 1a entrega do laudo"]),
        axis=1
    ))

    # Salvar arquivo atualizado
    output_path = os.path.join('P:\\Planilhas_dash\\download_atualizado_mensal.xlsx')
    db.to_excel(output_path, index=False)
    return f'Arquivo salvo em: {output_path}'


def planilha_preco_imovel(cpf, senha):
    #Abertura do Google Chrome
    try:
        options = webdriver.ChromeOptions()
        chrome_options = webdriver.ChromeOptions()
        download_dir = os.path.join(os.path.expanduser("~"), "Downloads") 
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }

        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--headless")
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 100)
        site_bradesco = driver.get('https://avaliacaobra.com.br')
        
        #----------------------------------------

        elemento_botao = driver.find_element(By.ID, 'btnFornec')
        elemento_botao.click()

        elemento_input_login = driver.find_element(By.ID, 'txtUsuario')
        elemento_input_login.send_keys(cpf)

        elemento_input_senha = driver.find_element(By.XPATH, '//*[@id="txtSenha"]')
        elemento_input_senha.send_keys(senha)

        elemento_botao_enviar = driver.find_element(By.ID, 'btnEnviar')
        elemento_botao_enviar.click()

        # Diretório do usuário para a pasta Downloads
        user_dir = os.path.expanduser('~')
        downloads_dir = 'P:\\Planilhas_dash'

        # Função para obter o arquivo mais recente com o nome "EmAndamento"
        def get_latest_em_andamento():
            files = [f for f in os.listdir(downloads_dir) if f.startswith("EmAndamento") and f.endswith(".xlsx")]
            if not files:
                raise FileNotFoundError("Nenhum arquivo 'EmAndamento.xlsx' encontrado na pasta Downloads.")
            
            # Obter o caminho completo e a data de modificação de cada arquivo
            files_with_time = [(os.path.join(downloads_dir, f), os.path.getmtime(os.path.join(downloads_dir, f))) for f in files]
            
            # Selecionar o arquivo com a data de modificação mais recente
            latest_file = max(files_with_time, key=lambda x: x[1])[0]
            return latest_file

        # Chamar a função para obter o arquivo mais recente
        file_path = get_latest_em_andamento()
        sleep(4)

        # Ler o arquivo Excel pulando a primeira linha
        file_path_plan = pd.read_excel(file_path, skiprows=1)
        if not isinstance(file_path_plan, pd.DataFrame):
            raise ValueError("file_path_plan is not a DataFrame. Please check the input file.")
        print(file_path_plan)

        #Inicio da iteração em cada linha para buscar as informações devidas
        for index, value in file_path_plan.iterrows():
            prop = value['Solicitação']
            print(prop)
            qtd = file_path_plan
            print(f'{qtd.shape[0]}/{index + 1}')

            solicitacao_avaliacao = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="id_0"]'))).click()

            #Transformar o Iframe em um frame clicavel
            iframe = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div/iframe')))
            driver.switch_to.frame(iframe)
            #----------------------------------------

            #Adiciona a proposta na busca e procura por ela
            inserir_proposta = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/form/div[5]/div/table/tbody/tr/td[1]/div[1]/table/tbody/tr/td[2]/div/div/input'))).send_keys(prop)
            procurar_proposta = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/form/div[5]/div/table/tbody/tr/td[1]/div[1]/table/tbody/tr/td[2]/div/div/span/a/i'))).click()
            #----------------------------------------
            sleep(3)
            try:
                tipo_imov = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox27"]')))
                sleep(0.5)
                file_path_plan.at[index, 'Tipo Imovel'] = tipo_imov.get_attribute('value') if tipo_imov else 'Não encontrado'
            except Exception as e:
                file_path_plan.at[index, 'Tipo Imovel'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'Tipo Imovel': {e}")

            try:
                telefone = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox31"]')))
                sleep(0.5)
                file_path_plan.at[index, 'Telefone1'] = telefone.get_attribute('value') if telefone else 'Não encontrado'
            except Exception as e:
                file_path_plan.at[index, 'Telefone1'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'Telefone1': {e}")

            try:
                telefone2 = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox32"]')))
                sleep(0.5)
                file_path_plan.at[index, 'Telefone2'] = telefone2.get_attribute('value') if telefone2 else 'Não encontrado'
            except Exception as e:
                file_path_plan.at[index, 'Telefone2'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'Telefone2': {e}")

            try:
                telefone3 = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox33"]')))
                sleep(0.5)
                file_path_plan.at[index, 'Telefone3'] = telefone3.get_attribute('value') if telefone3 else 'Não encontrado'
            except Exception as e:
                file_path_plan.at[index, 'Telefone3'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'Telefone3': {e}")

            try:
                contato_element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox48"]')))
                sleep(0.5)
                file_path_plan.at[index, 'Contato'] = contato_element.get_attribute('value') if contato_element else 'Não encontrado'
            except Exception as e:
                file_path_plan.at[index, 'Contato'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'Contato': {e}")

            try:
                cep_element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox39"]')))
                sleep(0.5)
                file_path_plan.at[index, 'CEP'] = cep_element.get_attribute('value') if cep_element else 'Não encontrado'
            except Exception as e:
                file_path_plan.at[index, 'CEP'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'CEP': {e}")
            
            try:
                bairro_element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox40"]')))
                sleep(0.5)
                file_path_plan.at[index, 'Bairro'] = bairro_element.get_attribute('value') if bairro_element else 'Não encontrado'
            except Exception as e:
                file_path_plan.at[index, 'Bairro'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'Bairro': {e}")

            try:
                matricula = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="txtMatricula"]')))
                sleep(0.5)
                file_path_plan.at[index, 'Matricula'] = matricula.get_attribute('value') if matricula else 'Não encontrado'
            except Exception as e:
                file_path_plan.at[index, 'Matricula'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'Matricula': {e}")

            try:
                valor_imovel = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox54"]')))
                sleep(0.5)
                file_path_plan.at[index, 'Valor Imóvel'] = valor_imovel.get_attribute('value') if valor_imovel else 'Não encontrado'
            except Exception as e:
                file_path_plan.at[index, 'Valor Imóvel'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'Valor Imóvel': {e}")

            try:
                endereco = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox41"]')))
                sleep(0.5)
                file_path_plan.at[index, 'Endereço do Imóvel'] = endereco.get_attribute('value') if endereco else 'Não encontrado'
            except Exception as e:
                file_path_plan.at[index, 'Endereço do Imóvel'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'Endereço do Imóvel': {e}")

            try:
                bt_horario = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="btnAgendvisita"]')))
                bt_horario.click()
                data_visita = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox9"]')))
                hora_visita = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="TextBox11"]')))
                data = data_visita.get_attribute('value') if data_visita else 'Não encontrado'
                hora = hora_visita.get_attribute('value') if hora_visita else 'Não encontrado'
                horario = f'{data} {hora}'
                file_path_plan.at[index, 'Data Visita Sistema'] = horario
                fechar_bloco = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="form1"]/div[7]/div[1]/button/span[1]'))).click()
            except Exception as e:
                file_path_plan.at[index, 'Data Visita Sistema'] = 'Erro ao buscar'
                print(f"Erro ao buscar 'Data Visita Sistema': {e}")

            sleep(1)
            driver.refresh()
            sleep(1)
            print(file_path_plan)
        #Fim da Iteração ----------------------------------------
        #Separar "Endereço" e "Numero_Complemento" com a vírgula como delimitador
        file_path_plan[['Endereço', 'Numero_Complemento']] = file_path_plan['Endereço do Imóvel'].str.split(',', expand=True, n=1)
        #Usar " - " como delimitador para separar "Numero" e "Complemento" e preencher os valores ausentes
        numero_complemento_split = file_path_plan['Numero_Complemento'].str.split(' - ', expand=True, n=1)

        #Verificar se o split retornou duas colunas e ajustar conforme necessário
        if numero_complemento_split.shape[1] == 2:
            #Atribuir os valores para as colunas "Numero" e "Complemento"
            file_path_plan['Numero'] = numero_complemento_split[0]
            file_path_plan['Complemento'] = numero_complemento_split[1].fillna('')
        else:
            #Caso contrário, preencher "Complemento" com strings vazias
            file_path_plan['Numero'] = numero_complemento_split[0]
            file_path_plan['Complemento'] = ''

        #Remover colunas intermediárias
        file_path_plan.drop(columns=['Numero_Complemento', 'Endereço do Imóvel'], inplace=True)

        output_path = 'P:\\Planilhas_dash\\Bradesco_vencimentos.xlsx'
        file_path_plan.to_excel(output_path, index=False)
        #----------------------------------------

        print(f'DataFrame salvo em: {output_path}')

        return output_path
    except Exception as e:
        print(f"Ocorreu o seguinte erro {e}")


def ajustar_bradesco_producao():
    """Cria a coluna de vencimentos e ajusta para que, se o vencimento cair em um feriado,
    seja adicionado um dia útil adicional.
    
    Returns:
        str: Retorna uma mensagem de sucesso ou erro.
    """
    try:
        # Lista de feriados
        feriados = pd.to_datetime([
            '2025-01-01', '2025-03-03', '2025-03-04', '2025-04-18', '2025-04-21', '2025-05-01',
            '2025-06-19', '2025-09-07', '2025-10-12', '2025-11-02', '2025-11-15', '2025-11-20', '2025-12-25'
        ]).normalize()

        # Carregar a base de dados
        db_base = pd.read_excel('P:\\Planilhas_dash\\Bradesco_vencimentos.xlsx')
        print('Arquivo carregado com sucesso.')

        # Verificar se a coluna "Data Envio Solicitação" existe
        if "Data Envio Solicitação" not in db_base.columns:
            return "Erro: A coluna 'Data Envio Solicitação' não foi encontrada no arquivo."

        # Identificar valores inválidos antes da conversão
        print("Verificando valores nulos ou inválidos antes da conversão...")
        if db_base["Data Envio Solicitação"].isna().sum() > 0:
            print("⚠️ Aviso: Existem valores nulos na coluna 'Data Envio Solicitação'.")

        # Forçar conversão para datetime, ignorando erros
        db_base["Data Envio Solicitação"] = pd.to_datetime(
            db_base["Data Envio Solicitação"], 
            errors='coerce',  # Converte erros para NaT (em vez de quebrar o código)
            format='mixed',
            dayfirst=True
        )

        # Verificar se há datas inválidas que viraram NaT
        if db_base["Data Envio Solicitação"].isna().sum() > 0:
            print("⚠️ Existem datas inválidas que foram convertidas para NaT. Linhas problemáticas:")
            print(db_base[db_base["Data Envio Solicitação"].isna()])  # Mostra as linhas com erro

        print('Coluna "Data Envio Solicitação" convertida com sucesso.')

        # Função para calcular o vencimento ignorando feriados
        def calcular_vencimento(data_inicio, horas_uteis, feriados):
            if pd.isna(data_inicio):
                return pd.NaT  # Retorna NaT para evitar erro ao tentar somar datas inválidas

            vencimento = data_inicio

            while horas_uteis > 0:
                vencimento += BDay(1)  # Avança um dia útil
                if vencimento.normalize() in feriados:
                    continue  # Se for feriado, não conta esse dia
                horas_uteis -= 10  # Assume que cada dia útil tem 10 horas

            return vencimento

        # Aplicando a função de vencimento para cada linha
        db_base["Vencimentos"] = db_base.apply(
            lambda row: calcular_vencimento(
                row["Data Envio Solicitação"],
                30 if row["Tipo Imovel"] in ['CASA', 'APARTAMENTO', 'PRÉDIO RESIDENCIAL', 'CONDOMÍNIO RESIDENCIAL'] else 60,
                feriados
            ),
            axis=1
        )

        # Salvando no Excel
        db_base.to_excel('P:\\Planilhas_dash\\Bradesco_vencimentos_imovel.xlsx', index=False, engine='openpyxl')
        print('Vencimentos ajustados com sucesso.')
        return "Processo concluído com sucesso."

    except FileNotFoundError:
        return 'Arquivo Bradesco não encontrado'
    except Exception as e:
        return f'Ocorreu um erro: {e}'


def vencimento_avm():
    db_avm_path = 'P:\\Planilhas_dash\\inspectos_ajustada.xlsx'
    db_avm = pd.read_excel(db_avm_path)
    db_avm_filtrado = db_avm[db_avm['Tipo Inspeção'] == 'AVM']
    print(db_avm_filtrado)


def bradesco_controle():
    url = 'https://docs.google.com/spreadsheets/d/1ejuWam5JSkmSFi0ggtkZwH_8Iig9CM2zhtBUArICsXY/export?format=xlsx'
    df = baixar_planilha(url)
    excel_file = pd.ExcelFile(df)
    df = pd.read_excel(excel_file, sheet_name='BRADESCO - Abril25')

    # Debug inicial
    print("\n👉 Dados CRUS - Data Solicitação:")
    print(df['Data Solicitação'].head(3).to_markdown())
    print("\n👉 Dados CRUS - Data Agendamento:")
    print(df['Data Agendamento Sistema'].head(3).to_markdown())

    # Função de conversão segura para datas BR
    def converter_data_br(data):
        if pd.isna(data):
            return pd.NaT
            
        try:
            # Primeiro tenta o parser com dayfirst=True
            data_conv = pd.to_datetime(data, dayfirst=True, format='mixed')
            
            # Verificação adicional para evitar mm/dd/yyyy
            if data_conv.month > 12:  # Se mês > 12, está invertido
                data_conv = pd.to_datetime(data, dayfirst=False, format='mixed')
            
            return data_conv
        except:
            return pd.NaT

    # Converter colunas de data
    df['Data Solicitação'] = df['Data Solicitação'].apply(converter_data_br)
    df['Data Agendamento Sistema'] = df['Data Agendamento Sistema'].apply(converter_data_br)

    # Debug pós-conversão
    print("\n🔧 Datas CONVERTIDAS - Data Solicitação:")
    print(df['Data Solicitação'].head(3).to_markdown())
    print("\n🔧 Datas CONVERTIDAS - Data Agendamento:")
    print(df['Data Agendamento Sistema'].head(3).to_markdown())

    # Cálculo das horas úteis
    def calcular_diff(row):
        if pd.isna(row['Data Solicitação']) or pd.isna(row['Data Agendamento Sistema']):
            return (0.0, 0.0)
        return calcular_horas_uteis(row['Data Solicitação'], row['Data Agendamento Sistema'])

    df[['Tempo marcacao sistema', 'Dias marcacao sistema']] = df.apply(calcular_diff, axis=1, result_type='expand')
    
    # Ajustes finais
    df['Dias marcacao sistema'] = np.floor(df['Dias marcacao sistema'])
    df['Tempo marcacao sistema'] = np.round(df['Tempo marcacao sistema'], decimals=1)
    
    df.to_excel('df_bradesco.xlsx', index=False)
    return df


def itau_controle():
    url = 'https://docs.google.com/spreadsheets/d/1ejuWam5JSkmSFi0ggtkZwH_8Iig9CM2zhtBUArICsXY/export?format=xlsx'
    df = baixar_planilha(url)
    excel_file = pd.ExcelFile(df)
    df = pd.read_excel(excel_file, sheet_name='ITAÚ - ABRIL25')

    # Função de conversão segura
    def converter_data_br(data):
        try:
            data_conv = pd.to_datetime(data, dayfirst=True, format='mixed')
            if data_conv.month > 12:  # Correção para datas invertidas
                data_conv = pd.to_datetime(data, dayfirst=False, format='mixed')
            return data_conv
        except:
            return pd.NaT

    # Converter colunas de data
    df['Data Criação'] = df['Data Criação'].apply(converter_data_br)
    df['Data Eng,'] = df['Data Eng,'].apply(converter_data_br)
    df['DATA DA VISTORIA'] = pd.to_datetime(df['DATA DA VISTORIA'], dayfirst=True, errors='coerce')

    # Cálculo das horas úteis
    def calcular_diff(row):
        if pd.isna(row['Data Criação']) or pd.isna(row['Data Eng,']):
            return (0.0, 0.0)
        return calcular_horas_uteis(row['Data Criação'], row['Data Eng,'])

    df[['Tempo marcacao sistema', 'Dias marcacao sistema']] = df.apply(calcular_diff, axis=1, result_type='expand')

    # Tratamento da vistoria
    df['data_hora_vistoria'] = df.apply(
        lambda row: pd.Timestamp.combine(
            row['DATA DA VISTORIA'],
            pd.to_datetime(row['HORARIO VISTORIA'], format='%H:%M:%S').time()
        ) if pd.notna(row['DATA DA VISTORIA']) and pd.notna(row['HORARIO VISTORIA']) else pd.NaT,
        axis=1
    )

    # Ajustes finais
    df['Dias marcacao sistema'] = np.floor(df['Dias marcacao sistema'])
    df['Tempo marcacao sistema'] = np.round(df['Tempo marcacao sistema'], decimals=1)
    
    df.to_excel('df_itau.xlsx', index=False)
    return df


def inspectos_controle():
    url = 'https://docs.google.com/spreadsheets/d/1ejuWam5JSkmSFi0ggtkZwH_8Iig9CM2zhtBUArICsXY/export?format=xlsx'
    df = baixar_planilha(url)
    excel_file = pd.ExcelFile(df)
    df = pd.read_excel(excel_file, sheet_name='SANTANDER Inspectos - Abril25')

    # Função de conversão robusta
    def converter_data_br(data):
        try:
            data_conv = pd.to_datetime(data, dayfirst=True, format='mixed')
            if data_conv.month > 12:  # Verificação de formato invertido
                data_conv = pd.to_datetime(data, dayfirst=False, format='mixed')
            return data_conv
        except:
            return pd.NaT

    # Converter colunas de data
    df['DATA'] = df['DATA'].apply(converter_data_br)
    df['DATA SISTEMA'] = df['DATA SISTEMA'].apply(converter_data_br)

    # Cálculo das horas úteis
    def calcular_diff(row):
        if pd.isna(row['DATA']) or pd.isna(row['DATA SISTEMA']):
            return (0.0, 0.0)
        return calcular_horas_uteis(row['DATA'], row['DATA SISTEMA'])

    df[['Tempo marcacao sistema', 'Dias marcacao sistema']] = df.apply(calcular_diff, axis=1, result_type='expand')

    # Ajustes finais
    df['Dias marcacao sistema'] = np.floor(df['Dias marcacao sistema'])
    df['Tempo marcacao sistema'] = np.round(df['Tempo marcacao sistema'], decimals=1)
    
    df.to_excel('df_inspectos.xlsx', index=False)
    return df


def busca_vivaintra_assinaturas():
    """_Buscador de Laudos no site 'VIVAINTRA'
    Busca os 4 tipos de Laudos, e busca cada parte individualmente
    """
        
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        driver = webdriver.Chrome(options=options)
        
        data_hoje = date.today() + timedelta(days=1)
        # data_inicio_mes = data_hoje.replace(day=1)
        data_anterior = data_hoje - timedelta(days=2)

        wait = WebDriverWait(driver, 400)
        driver.get("https://uonosanchez.vivaintra.com/admin-blog")
        driver.maximize_window()
        # Entra na conta
        elem_email_vivaintra = wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='username']")))
        sleep(1)
        elem_email_vivaintra.send_keys('contato.thiagolima12@gmail.com')
        elem_email_vivaintra.send_keys(Keys.ENTER)
        sleep(1 )
        elem_senha = driver.find_element(By.NAME, "password")
        sleep(1)
        elem_senha.send_keys('27310701')
        elem_senha.send_keys(Keys.ENTER)

        elem_seleciona = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="intranet-login"]/form/input'
        )))
        elem_seleciona.send_keys('27310701')
        elem_seleciona.send_keys(Keys.ENTER)

        elem_seleciona_produtividade = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/nav/div[1]/div[1]/a[2]/i'
        ))).click()

        elem_seleciona_requisicoes = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/div[1]/div[2]/div[2]/div/div[2]/div/div[10]/a' # Aqui ele entra nas requisições
        ))).click()
        sleep(1)
        todas_requisições = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[1]/div[2]/ul/li[1]/a'
        ))).click()
        sleep(1)
        
        data_inicio_conclusao = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[10]/div[1]/input'
        )))
        
        data_inicio_conclusao.send_keys(data_anterior.strftime('%d/%m/%Y'))
        
        data_fim_conclusao = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[10]/div[3]/div/input'
        )))
        
        data_fim_conclusao.send_keys(data_hoje.strftime('%d/%m/%Y'))
        
        body = driver.find_element(By.TAG_NAME, 'body')
        body.send_keys(Keys.HOME)
        sleep(1)
        lista_ativo = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[5]/div'
        ))).click()
        
        sleep(1)
        ativo = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[5]/div/select/option[2]'
        ))).click()
        sleep(1)

        lista = [8, 9, 10, 11,12]
        nome_arch = ['Assinaturas-AVM', 'Assinaturas-Bradesco', 'Assinaturas-Itaú', 'Assinaturas-Presenciais', 'Assinaturas-Bancos']
        
        for i, nome in zip(lista, nome_arch):
            print('entrei no for')
            requisicao_list = wait.until(EC.visibility_of_element_located(
            (By.XPATH,
                f'//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[2]/div/select'
            ))).click()
            print('clicou na lista')
            sleep(1)
            requisicao = wait.until(EC.visibility_of_element_located(
            (By.XPATH,
                f'//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[2]/div/select/option[{i}]'
            ))).click()
            print('clicou na opção')  
            
            sleep(1)
            
            botao_bscr = wait.until(EC.visibility_of_element_located(
            (By.NAME,
                'submit'
            ))).send_keys(Keys.ENTER)
            sleep(6)
            
            gerar_excel = wait.until(EC.visibility_of_element_located(
            (By.XPATH,
                f'/html/body/nav/div[3]/a'
            ))).click()
            sleep(1)
            
            baixar_planilha = wait.until(EC.visibility_of_element_located(
            (By.XPATH,
                f'//*[@id="body-content"]/div[2]/div[2]/div[1]/a'
            ))).click()
            sleep(1)
            caminho = encontrar_arquivo_mais_recente('Exportacao', usuario, 'csv')
            
            mover_arquivos(caminho, nome)
            sleep(4)
            
    except Exception as e:
        print(f"Erro ao buscar 'Laudo': {e}")
        return


if __name__ == "__main__":
    # lista = [8, 9, 10, 11]
    # nome_arquivo_conferencia = ['avm_viva-conferencia','bradesco_viva-conferencia', 'itau_viva-conferencia', 'presenciais-conferencia']

    # for i in lista:
    #     busca_vivaintra(i)
    # try:
    #     inspectos_controle()
    # except Exception as e:
    #     print(f"Erro ao buscar 'Laudo': {e}")
    # try:
    #     itau_controle()
    # except Exception as e:
    #     print(f"Erro ao buscar 'Laudo': {e}")
    # try:
    #     bradesco_controle()
    # except Exception as e:
    #     print(f"Erro ao buscar 'Laudo': {e}")
    # itau_controle()
    # busca_inspectos()
    # busca_bradesco()
    # planilha_preco_imovel(login_usuario_bradesco, login_senha_bradesco)
    ajustar_inspectos()
    # ajustar_bradesco_producao()
    # vencimentos_itau_mensal()
    # vencimentos_itau()
    # busca_vivaintra_assinaturas()
    # bradesco_controle()
    # inspectos_data_agendamento()
    # busca_vivaintra(9)
    # ajustar_bradesco_producao()
