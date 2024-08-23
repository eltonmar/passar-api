import requests
from bs4 import BeautifulSoup
import pandas as pd

# URL base do site
base_url = 'https://www.lepostiche.com.br/malas#/pagina-{}'
all_products = []

# Loop para percorrer as páginas de 1 a 12
for page in range(1, 13):
    url = base_url.format(page)
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Seleciona todos os itens do produto na página
    products = soup.find_all('li', class_='product-ganhe-seguro-viagem')

    # Loop para coletar informações de cada produto
    for product in products:
        product_name = product.get('data-name')
        product_id = product.get('data-product-id')
        product_category = product.get('data-category')
        product_price = product.find('span', class_='preco').text.strip() if product.find('span', class_='preco') else 'Preço não disponível'
        product_discount = product.find('span', class_='porcentagem').text.strip() if product.find('span', class_='porcentagem') else 'Sem desconto'

        # Exibe os detalhes do produto no terminal
        print(f"Produto: {product_name}, ID: {product_id}, Categoria: {product_category}, Preço: {product_price}, Desconto: {product_discount}")

        all_products.append({
            'Nome': product_name,
            'ID': product_id,
            'Categoria': product_category,
            'Preço': product_price,
            'Desconto': product_discount
        })

    print(f"Página {page} concluída")

# Convertendo a lista de produtos em um DataFrame do pandas
df = pd.DataFrame(all_products)

# Salvando os dados em um arquivo Excel
df.to_excel('lepostiche_products.xlsx', index=False)

print("Scraping concluído e dados salvos em lepostiche_products.xlsx")
