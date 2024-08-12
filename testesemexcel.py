import requests
from bs4 import BeautifulSoup

# URL da página
url = "https://www.bagaggio.com.br/malas?order=OrderByReleaseDateDESC"

# Faz a requisição para a página
response = requests.get(url)
soup = BeautifulSoup(response.content, 'html.parser')

# Verifica a estrutura da página para encontrar os produtos
product_elements = soup.find_all('div', class_='vtex-product-summary-2-x-container')

# Verifica quantos elementos de produtos foram encontrados
print(f"Total de produtos encontrados: {len(product_elements)}")

# Lista para armazenar informações dos produtos
products = []

# Extrai as informações dos produtos
for product in product_elements:
    try:
        title = product.find('container', class_='vtex-store-components-3-x-productName').text.strip()
        price = product.find('container', class_='vtex-product-showcase-2-x-currencyContainer').text.strip()
        link = product.find('container', class_='vtex-product-summary-2-x-clearLink')['href']
        products.append({'Title': title, 'Price': price, 'Link': link})
    except AttributeError as e:
        print(f"Erro ao processar um produto: {e}")

# Imprime os produtos no terminal
for product in products:
    print(f"Title: {product['Title']}, Price: {product['Price']}, Link: {product['Link']}")

# Verifica se a lista de produtos está vazia
if not products:
    print("Nenhum produto foi encontrado. Verifique os seletores usados.")
