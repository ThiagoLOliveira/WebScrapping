import requests

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)...",
    "X-Requested-With": "XMLHttpRequest",  # Pode ser necessário
}

url_api = "https://www.zapimoveis.com.br/api/search/paginated?..."
response = requests.get(url_api, headers=headers)
dados_imoveis = response.json()  # Extraia preços, endereços, etc.