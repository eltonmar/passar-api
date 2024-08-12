import requests
from bs4 import BeautifulSoup
import pandas as pd

# URL da página
url = "https://www.bagaggio.com.br/malas?order=OrderByReleaseDateDESC"

# Faz a requisição para a página
response = requests.get(url)
soup = BeautifulSoup(response.content, 'html.parser')

# Lista para armazenar informações dos produtos
products = []

# Verifica a estrutura da página para encontrar os produtos
product_elements = soup.find_all('div', class_='vtex-product-summary-2-x-container')

# Extrai as informações dos produtos
for product in product_elements:
    try:
        title = product.find('span', class_='vtex-store-components-3-x-productName').text.strip()
        price = product.find('span', class_='vtex-product-showcase-2-x-currencyContainer').text.strip()
        link = product.find('a', class_='vtex-product-summary-2-x-clearLink')['href']
        products.append({'Title': title, 'Price': price, 'Link': link})
    except AttributeError as e:
        print(f"Erro ao processar um produto: {e}")

# Converte para DataFrame e salva em um arquivo Excel
df = pd.DataFrame(products)
df.to_excel(r'C:\Users\BG-PROVISORIO\Desktop\teste.xlsx', index=False)

print("Dados salvos em 'C:\\Users\\BG-PROVISORIO\\Desktop\\teste.xlsx'")
