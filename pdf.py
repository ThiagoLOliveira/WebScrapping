import fitz  # PyMuPDF
import os
from pathlib import Path
import re

# Função para sanitizar nomes de arquivos
def limpar_nome(nome):
    return re.sub(r'[\\/*?:"<>|]', "_", nome)

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
