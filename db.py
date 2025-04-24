import pandas as pd
import os
import mysql.connector
import dotenv
from time import sleep

# Carregar variáveis de ambiente
dotenv.load_dotenv()
host = os.getenv("DB_HOST_LOCAL")
user = os.getenv("DB_USERNAME_LOCAL")
password = os.getenv("DB_PASSWORD_LOCAL")
database = os.getenv("DB_DATABASE_LOCAL")

# Conectar ao banco
connection = mysql.connector.connect(
    host=host,
    user=user,
    password=password,
    database=database,
    connection_timeout=60
)

cursor = connection.cursor()

# Ler planilha
df = pd.read_excel('imoveisOLX.xlsx')

# Montar tuplas
data_conferencia = [(
    row['Link'] if pd.notna(row['Link']) else "x",
    row['Endereço'] if pd.notna(row['Endereço']) else "x",
    row['CEP'] if pd.notna(row['CEP']) else "x",
    row['Valor Imóvel'] if pd.notna(row['Valor Imóvel']) else "1",
    row['Quartos'] if pd.notna(row['Quartos']) else "0",
    row['Banheiros'] if pd.notna(row['Banheiros']) else "0",
    row['Vagas na garagem'] if pd.notna(row['Vagas na garagem']) else "0",
    row['Área construída'] if pd.notna(row['Área construída']) else "0",
    row['Acabamentos'] if pd.notna(row['Acabamentos']) else "x",
    row['Tipo Imóvel'] if pd.notna(row['Tipo Imóvel']) else "x",
    row['Descrição'] if pd.notna(row['Descrição']) else "x",
) for _, row in df.iterrows()]

# Query (sem o ID se for auto_increment)
query = """
    INSERT INTO amostras 
    (link, endereco, cep, valor, quartos, banheiros, vagas, area, acabamentos, tipo_imovel, descricao)
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""

# Inserção por lotes
batch_size = 500  # pode ajustar entre 500–2000 dependendo da performance
for i in range(0, len(data_conferencia), batch_size):
    batch = data_conferencia[i:i+batch_size]
    try:
        cursor.executemany(query, batch)
        connection.commit()
        print(f"Lote {i // batch_size + 1} inserido com sucesso.")
        sleep(0.05)
    except mysql.connector.Error as err:
        print(f"Erro no lote {i // batch_size + 1}: {err}")
        connection.rollback()
        break

cursor.close()
connection.close()
