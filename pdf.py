import pandas as pd
import fitz
import os
import cv2
import pytesseract
import re
import mysql.connector
import dotenv
from pathlib import Path
from pdf2image import convert_from_path
import datetime
import openpyxl
import locale
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import pdfplumber

# Carregar variáveis de ambiente
dotenv.load_dotenv()
host = os.getenv("DB_HOST")
user = os.getenv("DB_USERNAME")
password = os.getenv("DB_PASSWORD")
database = os.getenv("DB_DATABASE")

# Conectar ao banco
connection = mysql.connector.connect(
    host=host,
    user=user,
    password=password,
    database=database,
    connection_timeout=60
)
cursor = connection.cursor()
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
poppler_path = r'C:\poppler-24.08.0\Library\bin'

# Função para converter coordenadas DMS (graus, minutos, segundos) para decimal
def dms_para_decimal(coordenadas_str):
    pattern = r"(\d+)°(\d+)'([\d.]+)\"?([NSWE])"
    matches = re.findall(pattern, coordenadas_str)

    if len(matches) != 2:
        raise ValueError("Formato de coordenadas inválido.")

    def converter(match):
        graus, minutos, segundos, direcao = match
        decimal = float(graus) + float(minutos)/60 + float(segundos)/3600
        if direcao in ['S', 'W']:
            decimal = -decimal
        return decimal

    lat = converter(matches[0])
    lon = converter(matches[1])
    return lat, lon

# Função para sanitizar nomes de arquivos
def limpar_nome(nome):
    return re.sub(r'[\\/*?:"<>|]', "_", nome)


def extrair_trecho(texto, inicio, fim):
    """
    Extrai o texto entre dois marcadores: 'inicio' e 'fim'.
    """
    padrao = rf"{re.escape(inicio)}(.*?){re.escape(fim)}"
    match = re.search(padrao, texto, flags=re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return "Trecho não encontrado"

# Função para buscar o valor após uma palavra-chave
def encontrar_valor(texto, chave, linha_abaixo=False, linha_acima=False):
    linhas = texto.split("\n")
    for i, linha in enumerate(linhas):
        if chave in linha:
            if linha_abaixo and i + 1 < len(linhas):
                return linhas[i + 1].strip()
            elif linha_acima and i - 1 >= 0:
                return linhas[i - 1].strip()
            else:
                return linha.split(chave)[-1].strip()
    return "Não encontrado"

# Função para ajustar os nomes dos arquivos PDF
def ajustar_nomes():
    diretorio = 'M:\\nova_leva_laudos\\'
    lista_arquivos = os.listdir(diretorio)

    for arquivo in lista_arquivos:
        if not arquivo.lower().endswith(".pdf"):
            continue  # Ignorar arquivos que não são PDF

        caminho_completo = os.path.join(diretorio, arquivo)
        print(f"Processando: {caminho_completo}")

        try:
            with fitz.open(caminho_completo) as doc:
                texto_pg1 = doc[0].get_text()
                uf = encontrar_valor(texto_pg1, "UF")
                cidade = encontrar_valor(texto_pg1, "Cidade")
                matricula = encontrar_valor(texto_pg1, "Matrícula")
                bairro = encontrar_valor(texto_pg1, "Bairro/Setor")
                num_controle = encontrar_valor(texto_pg1, "Nº Controle Interno / Ordem de Serviço")

                texto_pg2 = doc[1].get_text()
                data_laudo = encontrar_valor(texto_pg2, "Data Elaboração Laudo", linha_abaixo=True)
                if data_laudo == "Não encontrado":
                    texto_pg3 = doc[2].get_text()
                    data_laudo = encontrar_valor(texto_pg3, "Data Elaboração Laudo", linha_abaixo=True)
                elif data_laudo == "Não encontrado":
                    texto_pg4 = doc[3].get_text()
                    data_laudo = encontrar_valor(texto_pg4, "Data Elaboração Laudo", linha_abaixo=True)

            uf = limpar_nome(uf)
            cidade = limpar_nome(cidade)
            bairro = limpar_nome(bairro)
            num_controle = limpar_nome(num_controle)
            data_laudo = limpar_nome(data_laudo)

            # Renomear o arquivo
            path = Path(caminho_completo)
            novo_nome = f'{uf} - {cidade} - {bairro} - {matricula} - {num_controle} - {data_laudo}.pdf'
            novo_caminho = path.with_name(novo_nome)

            path.rename(novo_caminho)
            print(f"✅ Arquivo renomeado para: {novo_caminho}")

        except PermissionError as e:
            print(f"🚫 Erro de permissão ao renomear '{arquivo}': {e}")
        except Exception as e:
            print(f"⚠️ Erro ao processar '{arquivo}': {e}")


def ajustar_nomes_inspectos():

    diretorio = 'M:\\Thiago\\WebScrapping\\pdf\\'
    lista_arquivos = os.listdir(diretorio)
    for arquivo in lista_arquivos:
        if not arquivo.lower().endswith(".pdf"):
            continue  # Ignorar arquivos que não são PDF

        caminho_completo = os.path.join(diretorio, arquivo)
        print(f"Processando: {caminho_completo}")

        try:
            with fitz.open(caminho_completo) as doc:
                texto_pg1 = doc[0].get_text()
                num_controle = encontrar_valor(texto_pg1, "N° do Pedido", linha_abaixo=True)
            num_controle = limpar_nome(num_controle)

            # Renomear o arquivo
            path = Path(caminho_completo)
            novo_nome = f'{num_controle}.pdf'
            novo_caminho = path.with_name(novo_nome)

            path.rename(novo_caminho)
            print(f"✅ Arquivo renomeado para: {novo_caminho}")

        except PermissionError as e:
            print(f"🚫 Erro de permissão ao renomear '{arquivo}': {e}")
        except Exception as e:
            print(f"⚠️ Erro ao processar '{arquivo}': {e}")

# Função para encontrar o número da residência
def encontrar_numero_residencia(texto):
    # Divide o texto em linhas
    linhas = texto.split('\n')
    for linha in linhas:
        # Verifica se tem "N°" ou "Nº" sozinho seguido de número
        if re.search(r'\bN[º°]\b', linha):
            if not any(palavra in linha.lower() for palavra in ['pavimento', 'cartório', 'elevador', 'unidade', 'condomínio', 'ofício']):
                match = re.search(r'\bN[º°] ?[:\-]? ?(\d{1,5})\b', linha)
                if match:
                    return match.group(1)
    return "Não encontrado"


def encontrar_numero_apos_texto(texto, chave):
    """
    Procura um número imediatamente após uma chave específica no texto.
    Retorna o número encontrado ou 'Não encontrado'.
    """
    # Escapa caracteres especiais na chave para evitar conflitos na regex
    chave_escapada = re.escape(chave)

    # Padrão: chave + possível separador + número (com até 5 dígitos)
    padrao = rf"{chave_escapada}\s*[:\-]?\s*(\d{{1,5}})"
    
    match = re.search(padrao, texto, flags=re.IGNORECASE)
    if match:
        return match.group(1)
    
    return "Não encontrado"

# Função para encontrar o valor de outras paginas
def outras_linhas(valor, doc, texto):
    if valor == "Não encontrado":
        texto_pg2 = doc[1].get_text()
        valor = encontrar_valor(texto_pg2, f"{texto}", linha_abaixo=True)
    elif valor == "Não encontrado":
        texto_pg3 = doc[2].get_text()
        valor = encontrar_valor(texto_pg3, f"{texto}", linha_abaixo=True)
    
    return valor

# Função para fazer o scrapping dos PDFs
def scrapping_pdf():
    """
    Função para pegar dados dentro do PDF padrão Itaú.
    """
    # Defina o caminho do diretório onde os arquivos PDF estão localizados
    diretorio = 'M:\\Laudos Anteriores Itau\\'
    # Obtenha a lista de arquivos no diretório
    lista_arquivos = os.listdir(diretorio)

    dados = []
    # Itera sobre cada arquivo no diretório
    for arquivo in lista_arquivos:
        # Verifica se o arquivo é um PDF (ignorando maiúsculas/minúsculas)
        if not arquivo.lower().endswith(".pdf"):
            continue  # Ignora arquivos que não são PDF

        # Cria o caminho completo do arquivo
        caminho_completo = os.path.join(diretorio, arquivo)

        try:
            # Abre o arquivo PDF usando PyMuPDF
            with fitz.open(caminho_completo) as doc:

                # Informações do Imóvel Avaliando
                texto_pg1 = doc[0].get_text()
                texto_pg2 = doc[1].get_text()
                texto_pg3 = doc[2].get_text()
                num_controle = encontrar_valor(texto_pg1, "Nº Controle Interno / Ordem de Serviço")
                valor_compra_e_venda = encontrar_valor(texto_pg1, "Valor Compra Venda")
                matricula = encontrar_valor(texto_pg1, "Matrícula")
                logradouro = encontrar_valor(texto_pg1, "Logradouro")
                numero = encontrar_numero_residencia(texto_pg1)
                andar = encontrar_valor(texto_pg1, "Andar")
                complemento = encontrar_valor(texto_pg1, "Complemento")
                bairro = encontrar_valor(texto_pg1, "Bairro/Setor")
                cidade = encontrar_valor(texto_pg1, "Cidade")
                uf = encontrar_valor(texto_pg1, "UF")
                latitude = encontrar_valor(texto_pg1, "Latitude").split(' ')[0].strip()
                longitude = encontrar_valor(texto_pg1, "Longitude")
                idade = encontrar_valor(texto_pg1, "Anos")
                coordenadas = encontrar_valor(texto_pg1, "Coordenadas do imóvel:")

                if coordenadas == "Não encontrado":
                    coordenadas = encontrar_valor(texto_pg2, "Coordenadas do imóvel:")

                # Usa regex para extrair dois números com ponto decimal (latitude e longitude)
                match = re.search(r"(-?\d+\.\d+)[, ]+\s*(-?\d+\.\d+)", coordenadas)

                if match:
                    coordenadas_lat_long = [match.group(1), match.group(2)]

                    if latitude == "Não encontrado" or latitude == "":
                        latitude = coordenadas_lat_long[0].strip()

                    if longitude == "Não encontrado" or longitude == "":
                        longitude = coordenadas_lat_long[1].strip()

                
                # Dimensões do Imóvel Avaliando
                areatotal = encontrar_valor(texto_pg1, "Área Total:")
                testada = encontrar_valor(texto_pg1, "Testada (Frente):")
                fracao = encontrar_valor(texto_pg1, "Fração Ideal:")
                area_privativa = encontrar_valor(texto_pg1, "Área Privativa:", linha_abaixo=True)
                area_comum = encontrar_valor(texto_pg1, "Área Comum (m²):", linha_abaixo=True)
                area_total = encontrar_valor(texto_pg1, "Área Total (m²):", linha_abaixo=True)
                area_averbada = encontrar_valor(texto_pg1, "Área Averbada:")
                area_nao_averbada = encontrar_valor(texto_pg1, "Área não Averbada:")

                # Edifício ao qual pertence o Imóvel Avaliando
                pavimentos = encontrar_valor(texto_pg1, "N° de Pavimentos:")
                unidades_por_andar = encontrar_valor(texto_pg1, "N° Unidades Por Andar:")
                total_unindades_condominio = encontrar_numero_apos_texto(texto_pg1, "condomínio:")
                qtd_elevadores = encontrar_valor(texto_pg1, "N° de Elevadores:")
                descricao_andares_pavimentos = encontrar_valor(texto_pg1, "Andares/Pavimentos")
                uso_edificio = encontrar_valor(texto_pg1, "Uso do Edifício:")
                uso_imovel_avaliando = encontrar_valor(texto_pg1, "Uso do Imóvel Avaliando:")
                vagas_cobertas = encontrar_valor(texto_pg1, "Cobertas").split(' ')[0].strip()
                vagas_descobertas = encontrar_valor(texto_pg1, "Descobertas").split(' ')[2].strip()
                vagas_privativas = encontrar_valor(texto_pg1, "Privativas").strip()
                fechamento_paredes = encontrar_valor(texto_pg1, "Fechamento das Paredes:")
                total_banheiros = encontrar_valor(texto_pg1, "Total de Banheiros:")
                fachada_principal = encontrar_valor(texto_pg1, "Fachada Principal")
                esquadrias = encontrar_valor(texto_pg1, "Esquadrias")
                num_pavimentos_unidade = encontrar_valor(texto_pg1, "N° Pavimentos da Unidade")
                num_dormitorios = encontrar_valor(texto_pg1, "Nº Dormitórios")
                
                obs_finais = extrair_trecho(texto_pg1, "Observações Finais", "Amostras")
                if obs_finais == "Trecho não encontrado":
                    obs_finais = extrair_trecho(texto_pg2, "Observações Finais", "Amostras")
                # Valores
                valor_mercado_total = encontrar_valor(texto_pg1, "Valor de Mercado Total do Imóvel:", linha_abaixo=True)
                if valor_mercado_total == "Não encontrado":
                    valor_mercado_total = encontrar_valor(texto_pg2, "Valor de Mercado Total do Imóvel:", linha_abaixo=True)
                
                valor_liquidez = encontrar_valor(texto_pg1, "Valor de Liquidez:")
                if valor_liquidez == "Não encontrado":
                    valor_liquidez = encontrar_valor(texto_pg2, "Valor de Liquidez:")
                
                data_elaboracao = encontrar_valor(texto_pg2, "Data Elaboração Laudo", linha_abaixo=True)
                if data_elaboracao == "Não encontrado":
                    data_elaboracao = encontrar_valor(texto_pg3, "Data Elaboração Laudo", linha_abaixo=True)
                    
                valor_terreno = encontrar_valor(texto_pg2, "Valor Terreno:")
                if valor_terreno == "Não encontrado":
                    valor_terreno = encontrar_valor(texto_pg3, "Valor Terreno:")
                    
                valor_edificacao = encontrar_valor(texto_pg2, "Valor Edificação:")
                if valor_edificacao == "Não encontrado":
                    valor_edificacao = encontrar_valor(texto_pg3, "Valor Edificação::")

                # print('data_elaboracao', data_elaboracao)
                print('-' * 80)
                # print('Coordenadas: ', coordenadas)
                
                if latitude == "Não encontrado":
                    latitude = coordenadas.split(' ')[0].strip()
                    print('Latitude coordenadas', latitude)
                else:
                    print('Latitude encontrada', latitude)

                if longitude == "Não encontrado":
                    longitude = coordenadas.split(' ')[1].strip()
                    print('Longitude coordenadas', longitude)
                else:
                    print('Longitude encontrada', longitude)


                
                origem = "Desconhecido"
                for pagina in doc:
                    # print(f"Página {pagina.number + 1}:")
                    # print(pagina.get_text())
                    # print("\n" + "-" * 80 + "\n")
                    # texto_pagina = pagina.get_text().lower().strip()
                    texto_pagina = pagina.get_text().lower().strip()
                    if "banco santander" in texto_pagina:
                        origem = "VS"
                        break
                    elif "itaú" in texto_pagina or "itau" in texto_pagina:
                        origem = "VI"
                        break

                    
                # Adiciona os dados a um DataFrame
                dados.append({
                    "Proposta": num_controle,
                    'Valor Compra e Venda': valor_compra_e_venda,
                    'Matrícula': matricula,
                    'Logradouro': logradouro,
                    'Número': numero,
                    'Andar': andar,
                    'Complemento': complemento,
                    'Bairro': bairro,
                    'Cidade': cidade,
                    'UF': uf,
                    'Latitude': latitude,
                    'Longitude': longitude,
                    'Coordenadas': coordenadas,
                    'Área Total': areatotal,
                    'Testada (Frente)': testada,
                    'Fração Ideal': fracao,
                    'N° de Pavimentos': pavimentos,
                    'N° Unidades Por Andar': unidades_por_andar,
                    'Nº Total de Unidades no condomínio': total_unindades_condominio,
                    'N° de Elevadores': qtd_elevadores,
                    'Andares/Pavimentos': descricao_andares_pavimentos,
                    'Uso do Edifício': uso_edificio,
                    'Uso do Imóvel Avaliando': uso_imovel_avaliando,
                    'Cobertas': vagas_cobertas,
                    'Descobertas': vagas_descobertas,
                    'Privativas': vagas_privativas,
                    'Fechamento das Paredes': fechamento_paredes,
                    'Total de Banheiros': total_banheiros,
                    'Fachada Principal': fachada_principal,
                    'Esquadrias': esquadrias,
                    'N° Pavimentos da Unidade': num_pavimentos_unidade,
                    'Nº Dormitórios': num_dormitorios,
                    'Área Privativa': area_privativa,
                    'Área Comum': area_comum,
                    'Área Total': area_total,
                    'Área Averbada': area_averbada,
                    'Área não Averbada': area_nao_averbada,
                    'Valor de Mercado Total do Imóvel': valor_mercado_total,
                    'Valor de Liquidez': valor_liquidez,
                    'Valor Terreno': valor_terreno,
                    'Valor Edificação': valor_edificacao,
                    'Data Elaboração Laudo': data_elaboracao,
                    'Observações Finais': obs_finais
                    })
                data_elaboracao_formatada = None
                if data_elaboracao:
                    data_elaboracao = data_elaboracao.strip()
                    try:
                        data_elaboracao_formatada = datetime.datetime.strptime(data_elaboracao, "%d/%m/%Y").date()
                    except ValueError:
                        print(f"Formato inválido: {data_elaboracao}")
                        data_elaboracao_formatada = None
                # print('Data Laudo: ', data_elaboracao_formatada)
                # print('Origem: ', origem)
                print('Numero Controle: ', num_controle)
                print('Coordenadas: ', coordenadas)
                print('Latitude: ', latitude)
                print('Longitude: ', longitude)
                
                
                data_conferencia = [(
                    num_controle,
                    matricula,
                    latitude,
                    longitude,
                    cidade,
                    uf,
                    bairro,
                    valor_compra_e_venda,
                    data_elaboracao_formatada,
                    origem
                )]
                
                query = """
                    INSERT INTO laudos_anteriores 
                    (num_proposta, matricula, latitude, longitude, cidade, uf, bairro, valor_compra_venda, data_laudo, origem)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s )
                    ON DUPLICATE KEY UPDATE
                    matricula = VALUES(matricula),
                    latitude = VALUES(latitude),
                    longitude = VALUES(longitude),
                    cidade = VALUES(cidade),
                    uf = VALUES(uf),
                    bairro = VALUES(bairro),
                    valor_compra_venda = VALUES(valor_compra_venda),
                    data_laudo = VALUES(data_laudo),
                    origem = VALUES(origem)
                """
                
                cursor.execute(query, data_conferencia[0])
                connection.commit()
                
        except PermissionError as e:
            print(f"🚫 Erro de permissão '{arquivo}': {e}")
        except Exception as e:
            print(f"⚠️ Erro ao processar '{arquivo}': {e}")
    df = pd.DataFrame(dados)
    df.to_excel('M:\\Thiago\\WebScrapping\\teste_pdf\\resultado-cetip.xlsx', index=False)
    print("✅ Dados extraídos e salvos em resultado.xlsx")

# Função para ler as páginas do PDF
def ler_paginas():
    diretorio = 'M:\\Thiago\\WebScrapping\\teste_pdf\\'
    lista_arquivos = os.listdir(diretorio)

    for arquivo in lista_arquivos:
        if not arquivo.lower().endswith(".pdf"):
            continue

        caminho_completo = os.path.join(diretorio, arquivo)
        print(f"Processando: {caminho_completo}")

        try:
            with fitz.open(caminho_completo) as doc:
                for i, pagina in enumerate(doc):
                    print(f"Página {i + 1}:\n")
                    print(pagina.get_text())
                    print("\n" + "-" * 80 + "\n")
        except Exception as e:
            print(f"⚠️ Erro ao processar '{arquivo}': {e}"); 


def pegar_checkbox():
    # 1. PDF -> Imagem
    pages = convert_from_path("teste_pdf/MG - UBERLANDIA - SHOPPING PARK - 136622 - 09842077 - 13_11_2024.pdf", dpi=300)
    image_path = "pagina1.png"
    pages[0].save(image_path, "PNG")

    # 2. OpenCV: carregamento e pré-processamento
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

    # 3. Encontrar contornos
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    marcados = []

    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)

        # 3.1 Verifica se o contorno é parecido com um quadrado pequeno
        if 10 < w < 25 and 10 < h < 25:
            aspect_ratio = w / float(h)
            if 0.8 < aspect_ratio < 1.2:  # formato quadrado
                roi = thresh[y:y+h, x:x+w]
                filled = cv2.countNonZero(roi)

                if filled > (w * h) * 0.5:  # se estiver "marcado"
                    # Aumenta área de OCR horizontalmente
                    text_roi = img[y-5:y+h+5, x+w+5:x+w+300]
                    
                    # Melhora OCR
                    text_gray = cv2.cvtColor(text_roi, cv2.COLOR_BGR2GRAY)
                    text_thresh = cv2.threshold(text_gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

                    # Usa OCR com config melhorada
                    config = "--psm 6 -l por"
                    text = pytesseract.image_to_string(text_thresh, config=config).strip()
                    
                    if text:  # Evita vazios
                        marcados.append(text)

    print("Checkboxes marcados:")
    for item in marcados:
        print("-", item)


def laudos_bradesco():
    diretorio = 'P:\\BRADESCO ISOLADOS\\'
    lista_arquivos = os.listdir(diretorio)
    dados_formulario = [arq for arq in lista_arquivos if 'formulario' in arq.lower() or 'formulário' in arq.lower()]
    dados_xlsx = [xlsx for xlsx in dados_formulario if 'xlsx' in xlsx.lower() or 'xls' in xlsx.lower()]
    laudos_com_erros = []
    
    for arquivo in dados_xlsx:
        try:
            wb = openpyxl.load_workbook(os.path.join(diretorio, arquivo), data_only=True)
            aba = wb['FORMULARIO']
            # print('Aba ativa: ', aba.title)
            value_proposta = aba['F8'].value
            value_matricula = aba['J34'].value
            if value_matricula == None:
                value_matricula = aba['F34'].value
            values_coordenadas = aba['E101'].value
            lat, long = dms_para_decimal(values_coordenadas)
            values_bairro = aba['F15'].value
            values_cidade = aba['F16'].value
            values_uf = aba['F18'].value
            values_valor = aba['J133'].value
            # aba = wb['PAG.1']
            values_data = aba['F61'].value
            values_imovel = aba['E106'].value
            

            # Define o locale para o Brasil
            locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

            # Supondo que values_valor foi lido da célula como número:
            valor_formatado = locale.currency(values_valor, grouping=True, symbol=False)

            # print(valor_formatado)
            # print('Proposta: ', value_proposta)
            # print('Matrícula: ', value_matricula)
            # print('Bairro: ', values_bairro)
            # print('Cidade: ', values_cidade)
            # print('UF: ', values_uf)
            # print('Valor: R$', valor_formatado)
            # print('Imóvel: ', values_imovel)
            # print('Data: ', values_data)
            # print('Coordenadas: ', values_coordenadas)
            # print('Latitude: ', lat)
            # print('Longitude: ', long)
            
            valor_final = 'R$ ' + str(valor_formatado)
            
            data_conferencia = [(
                value_proposta,
                value_matricula,
                lat,
                long,
                values_cidade,
                values_uf,
                values_bairro,
                valor_final,
                values_data,
                'VB',
                values_imovel
            )]
            
            query = """
                INSERT INTO laudos_anteriores 
                (num_proposta, matricula, latitude, longitude, cidade, uf, bairro, valor_compra_venda, data_laudo, origem, tipo_imovel)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                matricula = VALUES(matricula),
                latitude = VALUES(latitude),
                longitude = VALUES(longitude),
                cidade = VALUES(cidade),
                uf = VALUES(uf),
                bairro = VALUES(bairro),
                valor_compra_venda = VALUES(valor_compra_venda),
                data_laudo = VALUES(data_laudo),
                origem = VALUES(origem),
                tipo_imovel = VALUES(tipo_imovel)
            """
            
            cursor.execute(query, data_conferencia[0])
            connection.commit()
            print('Laudo Salvo no Banco de Dados')
        except Exception as e:
            # print(f"⚠️ Erro ao abrir o arquivo '{arquivo}': aba não encontrada.")
            laudos_com_erros.append(arquivo)
            pd.DataFrame(laudos_com_erros).to_excel('P:\\BRADESCO ISOLADOS\\laudos_com_erros.xlsx', index=False)


def atualizar_tipo_imovel_com_excel():
    # Caminho do arquivo Excel
    caminho_excel = 'databases\\download.xlsx'
    
    # Carregar a planilha Excel
    df_excel = pd.read_excel(caminho_excel)
    
    # Iterar sobre as linhas da planilha
    for _, row in df_excel.iterrows():
        num_proposta = row['Proposta']
        tipo_imovel = row['Tipo do Imóvel']
        print('Atualizando: ', num_proposta, tipo_imovel)
        # Atualizar o banco de dados com base no num_proposta
        query = """
            UPDATE laudos_anteriores
            SET tipo_imovel = %s
            WHERE num_proposta = %s
        """
        cursor.execute(query, (tipo_imovel, num_proposta))
    
    # Confirmar as alterações no banco de dados
    connection.commit()
    print("Tipo de imóvel atualizado com base na planilha Excel.")


def extrair_primeira_coordenada(pdf_path):
    # Expressão regular para o padrão de coordenadas
    padrao = r'\d{1,2}°\d{1,2}\'\d{1,2}"[NS] ?\/ ?\d{1,2}°\d{1,2}\'\d{1,2}"[WE]'

    with fitz.open(pdf_path) as doc:
        # Itera pelas páginas do PDF começando da última
        for page in reversed(doc):
            texto = page.get_text()
            coordenadas = re.findall(padrao, texto)
            if coordenadas:
                return coordenadas[0] # Retorna apenas a primeira coordenada encontrada
    return None


def extrair_data(pdf_path):
    # Expressão regular para o padrão de coordenadas
    padrao = r'São Paulo, [A-Za-zçãéêíú\-]+, \d{1,2} de [A-Za-zçãéêíú]+ de \d{4}'

    with fitz.open(pdf_path) as doc:
        # Itera pelas páginas do PDF começando da última
        for page in reversed(doc):
            texto = page.get_text()
            coordenadas = re.findall(padrao, texto)
            if coordenadas:
                return coordenadas[0] # Retorna apenas a primeira coordenada encontrada
    return None


def verificar_se_avm(pdf_path):
    # Expressão regular para o padrão de coordenadas
    padrao = r'AVM'

    with fitz.open(pdf_path) as doc:
        for page in doc:
            texto = page.get_text()
            coordenadas = re.findall(padrao, texto)
            if coordenadas:
                return coordenadas[0]  # Retorna apenas a primeira coordenada encontrada
    return None


def scrapping_pdf_inspectos():
    """
    Função para pegar dados dentro do PDF padrão Itaú.
    """
    # Defina o caminho do diretório onde os arquivos PDF estão localizados
    diretorio = 'M:\\Thiago\\WebScrapping\\pdf\\'
    # Obtenha a lista de arquivos no diretório
    lista_arquivos = os.listdir(diretorio)

    dados = []
    meses = {'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06',
            'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
    # Itera sobre cada arquivo no diretório
    for arquivo in lista_arquivos:
        # Verifica se o arquivo é um PDF (ignorando maiúsculas/minúsculas)
        if not arquivo.lower().endswith(".pdf"):
            continue  # Ignora arquivos que não são PDF

        # Cria o caminho completo do arquivo
        caminho_completo = os.path.join(diretorio, arquivo)

        try:
            # Abre o arquivo PDF usando PyMuPDF
            with fitz.open(caminho_completo) as doc:

                # Informações do Imóvel Avaliando
                texto_pg1 = doc[0].get_text()
                # print(texto_pg1)
                # print('/n' '-' * 80)
                texto_pg2 = doc[1].get_text()
                texto_pg3 = doc[2].get_text()
                avm = verificar_se_avm(doc)
                if avm: 
                    print('AVM encontrado')
                else:
                    texto_pg7 = doc[6].get_text()
                    num_controle = encontrar_valor(texto_pg1, "N° do Pedido", linha_abaixo=True)
                    
                    if num_controle == "Não encontrado":
                        num_controle = encontrar_valor(texto_pg1, "Nº da Proposta", linha_abaixo=True)
                        
                    valor_compra_e_venda = encontrar_valor(texto_pg7, "Valor de avaliação para efeito de garantia", linha_abaixo=True)
                    
                    matricula = encontrar_valor(texto_pg1, "Matrícula", linha_abaixo=True)
                    logradouro = encontrar_valor(texto_pg1, "Endereço", linha_abaixo=True)
                    numero = encontrar_valor(texto_pg1, "Número", linha_abaixo=True)
                    complemento = encontrar_valor(texto_pg1, "Complemento", linha_abaixo=True)
                    bairro = encontrar_valor(texto_pg1, "Bairro", linha_abaixo=True)
                    tipo_imovel = encontrar_valor(texto_pg1, "Tipo do imóvel", linha_abaixo=True)
                    cidade = encontrar_valor(texto_pg1, "Municipio", linha_abaixo=True)
                    metodologia = encontrar_valor(texto_pg1, "METODOLOGIA APLICADA", linha_abaixo=True)
                    uf = encontrar_valor(texto_pg1, "UF", linha_abaixo=True)
                    nome_arquivo = arquivo.split('.pdf')[0]
                    
                    coordenadas = extrair_primeira_coordenada(doc)
                    if coordenadas:
                        print(coordenadas)
                        coordenadas.replace('/', '')
                        lat, long = dms_para_decimal(coordenadas)
                    else:
                        print('Coordenadas não encontradas')
                        lat, long = None, None
                        
                    data_elaboracao = extrair_data(doc)

                    # Adiciona os dados a um DataFrame
                    data_elaboracao_formatada = None
                    if data_elaboracao:
                        data_elaboracao_formatada = data_elaboracao.strip().split(',')
                        data = data_elaboracao_formatada[2].strip()
                        data = data.split(' ')
                        data = data[0] + '/' + meses[data[2]] + '/' + data[4]
                        data = datetime.datetime.strptime(data, "%d/%m/%Y").date()
                        print('Data:', data)
                        
                    print('-' * 80)
                    print('Coordenadas: ', coordenadas)
                    print('Latitude: ', lat)
                    print('Longitude: ', long)
                    # print('Data: ', data_elaboracao)
                    print('Data formatada: ', data)
                    print('Tipo de imóvel: ', tipo_imovel)
                    print('Valor: ', valor_compra_e_venda)
                    print('Número: ', numero)
                    print('Metodologia: ', metodologia)
                    print('Complemento: ', complemento)
                    print('Bairro: ', bairro)
                    print('Cidade: ', cidade)
                    print('UF: ', uf)
                    print('Logradouro: ', logradouro)
                    print('Matrícula: ', matricula)
                    print('Número do controle: ', num_controle)
                    print('Nome do arquivo: ', nome_arquivo)
                    print('-' * 80)
                    
                    dados.append({
                        "Proposta": num_controle,
                        'Valor Compra e Venda': valor_compra_e_venda,
                        'Matrícula': matricula,
                        'Logradouro': logradouro,
                        'Número': numero,
                        'Complemento': complemento,
                        'Bairro': bairro,
                        'Cidade': cidade,
                        'UF': uf,
                        'Latitude': lat,
                        'Longitude': long,
                        'Coordenadas': coordenadas,
                        'Data Elaboração Laudo': data,
                        'Metodologia Aplicada': metodologia,
                        'Tipo do Imóvel': tipo_imovel,
                        'Nome do Arquivo': nome_arquivo
                    })
                
                data_conferencia = [(
                    num_controle,
                    matricula,
                    lat,
                    long,
                    uf,
                    bairro,
                    valor_compra_e_venda,
                    data,
                    'VS',
                    tipo_imovel
                )]
                
                query = """
                    INSERT INTO laudos_anteriores 
                    (num_proposta, matricula, latitude, longitude, uf, bairro, valor_compra_venda, data_laudo, origem, tipo_imovel)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE
                    matricula = VALUES(matricula),
                    latitude = VALUES(latitude),
                    longitude = VALUES(longitude),
                    uf = VALUES(uf),
                    bairro = VALUES(bairro),
                    valor_compra_venda = VALUES(valor_compra_venda),
                    data_laudo = VALUES(data_laudo),
                    tipo_imovel = VALUES(tipo_imovel),
                    origem = VALUES(origem)
                """
                
                cursor.execute(query, data_conferencia[0])
                connection.commit()
                
        except PermissionError as e:
            print(f"🚫 Erro de permissão '{arquivo}': {e}")
        except Exception as e:
            print(f"⚠️ Erro ao processar '{arquivo}': {e}")
    df = pd.DataFrame(dados)
    df.to_excel('M:\\Thiago\\WebScrapping\\pdf\\resultado.xlsx', index=False)
    print("✅ Dados extraídos e salvos em resultado.xlsx")



def extrair_dados_endereco(texto):
    # Quebra o texto em linhas, removendo vazias e espaços extras
    linhas = [linha.strip() for linha in texto.split('\n') if linha.strip()]

    endereco = numero = complemento = bairro = municipio = ''

    idx_endereco = idx_bairro = -1

    # Localiza os índices das palavras-chave
    for i, linha in enumerate(linhas):
        if linha.lower() == "endereço":
            idx_endereco = i
        elif linha.lower() == "bairro":
            idx_bairro = i
            break  # Pode parar aqui, já achou os dois

    # Garante que encontramos ambos
    if idx_endereco != -1 and idx_bairro != -1:
        trecho = linhas[idx_endereco + 1:idx_bairro]

        if len(trecho) >= 2:
            endereco = trecho[0]
            numero = trecho[1]
            complemento = trecho[2] if len(trecho) >= 3 else ''
        
        # Agora os valores após "Bairro"
        if idx_bairro + 2 < len(linhas):
            bairro = linhas[idx_bairro + 1]
            municipio = linhas[idx_bairro + 2]

    return endereco, numero, complemento, bairro, municipio


def scrapping_pdf_avm():
    """
    Função para pegar dados dentro do PDF padrão AVM.
    """
    diretorio = 'M:\\Thiago\\WebScrapping\\pdf'
    lista_arquivos = os.listdir(diretorio)

    dados = []
    meses = {'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06',
            'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
    for arquivo in lista_arquivos:
        if not arquivo.lower().endswith(".pdf"):
            continue  # Ignora arquivos que não são PDF

        caminho_completo = os.path.join(diretorio, arquivo)

        try:
            with fitz.open(caminho_completo) as doc:

                texto_pg1 = doc[0].get_text()
                texto_pg2 = doc[1].get_text()
                texto_pg3 = doc[2].get_text()
                avm = verificar_se_avm(doc)
                if avm: 
                    print('AVM encontrado')
                    num_controle = encontrar_valor(texto_pg1, "N° do Pedido", linha_abaixo=True)
                    matricula = encontrar_valor(texto_pg1, "Matrícula", linha_abaixo=True)
                    valor_compra_e_venda = encontrar_valor(texto_pg2, "Valor de avaliação", linha_abaixo=True)
                    logradouro, numero, complemento, bairro, cidade = extrair_dados_endereco(texto_pg1)
                    data_elaboracao = extrair_data(doc)
                    if data_elaboracao:
                        data_elaboracao_formatada = data_elaboracao.strip().split(',')
                        data = data_elaboracao_formatada[2].strip()
                        data = data.split(' ')
                        data = data[0] + '/' + meses[data[2]] + '/' + data[4]
                        data = datetime.datetime.strptime(data, "%d/%m/%Y").date()
                        print('Data:', data)
                    uf = encontrar_valor(texto_pg1, "UF", linha_abaixo=True)
                    print('Logradouro: ', logradouro)
                    print('Número: ', numero)
                    print('Complemento: ', complemento)
                    print('Bairro: ', bairro)
                    print('Cidade: ', cidade)
                    print('UF: ', uf)
                    
                    endereco = f"{logradouro}, {numero}, {cidade}, {uf}"
                    lat, lng = get_lat_long(endereco)
                    if lat is None or lng is None:
                        lat, lng = 0, 0
                    print('Endereço: ', endereco)
                    print("Latitude:", lat, "Longitude:", lng)
                    print('Número do controle: ', num_controle)
                    print('Matrícula: ', matricula)
                    print('Valor: ', valor_compra_e_venda)
                    print('Data: ', data)
                    print('\n', '-' * 80, '\n')
                else:
                    print('AVM não encontrado')
                
                data_conferencia = [(
                    num_controle,
                    matricula,
                    lat,
                    lng,
                    uf,
                    bairro,
                    valor_compra_e_venda,
                    data,
                    'AVM',
                    'Apartamento'
                )]
                
                query = """
                    INSERT INTO laudos_anteriores 
                    (num_proposta, matricula, latitude, longitude, uf, bairro, valor_compra_venda, data_laudo, origem, tipo_imovel)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE
                    matricula = VALUES(matricula),
                    latitude = VALUES(latitude),
                    longitude = VALUES(longitude),
                    uf = VALUES(uf),
                    bairro = VALUES(bairro),
                    valor_compra_venda = VALUES(valor_compra_venda),
                    data_laudo = VALUES(data_laudo),
                    tipo_imovel = VALUES(tipo_imovel),
                    origem = VALUES(origem)
                """
                
                cursor.execute(query, data_conferencia[0])
                connection.commit()
                
        except PermissionError as e:
            print(f"🚫 Erro de permissão '{arquivo}': {e}")
        except Exception as e:
            print(f"⚠️ Erro ao processar '{arquivo}': {e}")
    df = pd.DataFrame(dados)
    df.to_excel('M:\\Thiago\\WebScrapping\\pdf\\resultado.xlsx', index=False)
    print("✅ Dados extraídos e salvos em resultado.xlsx")


def get_lat_long(endereco):
    geolocator = Nominatim(user_agent="latlng", timeout=10)
    try:
        location = geolocator.geocode(endereco)
        if location:
            return location.latitude, location.longitude
        else:
            return None, None
    except GeocoderTimedOut:
        return None, None


if __name__ == "__main__":
    # scrapping_pdf_inspectos()
    # pdf_to_text()
    scrapping_pdf_avm()
    # endereco = "RUA LALITA COSTA, 217, Salvador, BA"
    # lat, lng = get_lat_long(endereco)
    # print("Latitude:", lat, "Longitude:", lng)
    # ajustar_nomes_inspectos()
    # print("Primeira coordenada encontrada:", primeira_coordenada)
    # ajustar_nomes()
    # scrapping_pdf()
    # laudos_bradesco()
    # atualizar_tipo_imovel_com_excel()
    # print("Funções disponíveis:")
    # pegar_checkbox()
    # ler_paginas()