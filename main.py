import cloudscraper
from bs4 import BeautifulSoup
import pandas as pd

url = "https://www.zapimoveis.com.br/imovel/venda-cobertura-3-quartos-com-churrasqueira-tijuca-zona-norte-rio-de-janeiro-rj-118m2-id-2767013415/"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

scraper = cloudscraper.create_scraper()
response = scraper.get(url, headers=headers)
df = pd.DataFrame(columns=["Link","Área", "Banheiro", "Estacionamento", "Suíte"])

df["Link"] = [url]

campos = ["Área", "Banheiro", "Estacionamento", "Suíte"]
if response.status_code == 200:
    soup = BeautifulSoup(response.text, "html.parser")
    
    # Encontrar a div com a classe "details"
    area = soup.find_all("li", {"itemprop": "floorSize"})
    banheiros = soup.find_all("li", {"itemprop": "numberOfBathroomsTotal"})
    estacionamentos = soup.find_all("li", {"itemprop": "numberOfParkingSpaces"})
    suites = soup.find_all("li", {"itemprop": "numberOfSuites"})
    
    for values in zip(area, banheiros, estacionamentos, suites):
        # print(values)
        text = [value.get_text(strip=True) for value in values]
        print(text)
        for i, (line, campo) in enumerate(zip(text, campos)):
            # print(line.strip())
            parts = line.split(' ', 1)
            valor = parts[0]
            print(valor)
            df[campo] = [valor]
            df[campo] = [valor]
            df[campo] = [valor]
            df[campo] = [valor]
        print("\n" + "-"*50 + "\n")  # Separador visual entre blocos

else:
    print(f"Erro: {response.status_code}")

print(df)  # Imprime o DataFrame com os valores e características
df.to_excel("dados.xlsx", index=False)
