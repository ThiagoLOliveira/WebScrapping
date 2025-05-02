import fitz
import os
import cv2
import pytesseract
import re
import pandas as pd
from pathlib import Path
from pdf2image import convert_from_path

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
poppler_path = r'C:\poppler-24.08.0\Library\bin'

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
def encontrar_valor(texto, chave, linha_abaixo=False):
    linhas = texto.split("\n")
    for i, linha in enumerate(linhas):
        if chave in linha:
            if linha_abaixo and i + 1 < len(linhas):
                return linhas[i + 1].strip()
            else:
                return linha.split(chave)[-1].strip()
    return "Não encontrado"

# Função para ajustar os nomes dos arquivos PDF
def ajustar_nomes():
    diretorio = 'M:\\Laudos Anteriores Itau\\'
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
    diretorio = 'M:\\Thiago\\WebScrapping\\teste_pdf\\'
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
                coordenadas = encontrar_valor(texto_pg1, "Coordenadas do imóvel:")
                idade = encontrar_valor(texto_pg1, "Anos")
                if coordenadas == "Não encontrado":
                    texto_pg3 = doc[1].get_text()
                    coordenadas = encontrar_valor(texto_pg3, "Coordenadas do imóvel:")

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
                
                data_elaboração = encontrar_valor(texto_pg2, "Data Elaboração Laudo", linha_abaixo=True)
                if data_elaboração == "Não encontrado":
                    data_elaboração = encontrar_valor(texto_pg3, "Data Elaboração Laudo", linha_abaixo=True)
                    
                valor_terreno = encontrar_valor(texto_pg2, "Valor Terreno:")
                if valor_terreno == "Não encontrado":
                    valor_terreno = encontrar_valor(texto_pg3, "Valor Terreno:")
                    
                valor_edificacao = encontrar_valor(texto_pg2, "Valor Edificação:")
                if valor_edificacao == "Não encontrado":
                    valor_edificacao = encontrar_valor(texto_pg3, "Valor Edificação::")


                # print("-"*80)
                print(f"Processando: {caminho_completo}")
                print("Proposta: ", num_controle)
                print("Valor Compra e Venda: ", valor_compra_e_venda)
                print("Matrícula: ", matricula)
                print("Logradouro: ", logradouro)
                print("Número: ", numero)
                print("Andar: ", andar)
                print("Complemento: ", complemento)
                print("Bairro: ", bairro)
                print("Cidade: ", cidade)
                print("UF: ", uf)
                print("Latitude: ", latitude)
                print("Longitude: ", longitude)
                print("Coordenadas: ", coordenadas)
                print("Área Total: ", areatotal)
                print("Testada (Frente): ", testada)
                print("Fração Ideal: ", fracao)
                print("N° de Pavimentos: ", pavimentos)
                print("N° Unidades Por Andar: ", unidades_por_andar)
                print("Nº Total de Unidades no condomínio: ", total_unindades_condominio)
                print("N° de Elevadores: ", qtd_elevadores)
                print("Andares/Pavimentos: ", descricao_andares_pavimentos)
                print("Uso do Edifício: ", uso_edificio)
                print("Uso do Imóvel Avaliando: ", uso_imovel_avaliando)
                print("Cobertas: ", vagas_cobertas)
                print("Descobertas: ", vagas_descobertas)
                print("Privativas: ", vagas_privativas)
                print("Fechamento das Paredes: ", fechamento_paredes)
                print("Total de Banheiros: ", total_banheiros)
                print("Fachada Principal: ", fachada_principal)
                print("Esquadrias: ", esquadrias)
                print("N° Pavimentos da Unidade: ", num_pavimentos_unidade)
                print("Nº Dormitórios: ", num_dormitorios)
                print("Área Privativa: ", area_privativa)
                print("Área Comum: ", area_comum)
                print("Área Total: ", area_total)
                print("Área Averbada: ", area_averbada)
                print("Área não Averbada: ", area_nao_averbada)
                print("Valor de Mercado Total do Imóvel: ", valor_mercado_total)
                print("Valor de Liquidez: ", valor_liquidez)
                print("Data Elaboração Laudo: ", data_elaboração)
                print("Valor Terreno: ", valor_terreno)
                print("Valor Edificação: ", valor_edificacao)
                print("Observações Finais: ", obs_finais)
                print("-"*80)

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
                    'Data Elaboração Laudo': data_elaboração,
                    'Observações Finais': obs_finais
                    })
                
        except PermissionError as e:
            print(f"🚫 Erro de permissão '{arquivo}': {e}")
        except Exception as e:
            print(f"⚠️ Erro ao processar '{arquivo}': {e}")
    df = pd.DataFrame(dados)
    df.to_excel('M:\\Thiago\\WebScrapping\\teste_pdf\\resultado.xlsx', index=False)
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

coords = []

def click_event(event, x, y, flags, param):
    if event == cv2.EVENT_LBUTTONDOWN:
        coords.append((x, y))
        print(f"Coordenada capturada: {x}, {y}")

# Carrega a imagem
img = cv2.imread("pagina1.png")
# Redimensionar para caber na tela (ex: 50% do tamanho original)
scale_percent = 50
width = int(img.shape[1] * scale_percent / 100)
height = int(img.shape[0] * scale_percent / 100)
resized = cv2.resize(img, (width, height))

cv2.imshow("Clique nas caixas", resized)
cv2.setMouseCallback("Clique nas caixas", click_event)

cv2.waitKey(0)
cv2.destroyAllWindows()

# Depois que fechar a janela, você terá as coordenadas em `coords`
print("Coordenadas finais:", coords)

if __name__ == "__main__":
    # ajustar_nomes()
    # scrapping_pdf()
    print("Funções disponíveis:")
    # pegar_checkbox()
    # ler_paginas()