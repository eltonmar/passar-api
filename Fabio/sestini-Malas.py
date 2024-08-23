import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configurar o Selenium para usar o ChromeDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Iniciar a contagem de tempo
start_time = time.time()

# Contador de requisições
request_count = 0

# Lista para armazenar as informações dos produtos
products_data = []

# Função para extrair informações de produtos em uma página
def extract_product_info():
    products = driver.find_elements(By.CLASS_NAME, 'x-shelf__item')
    for product in products:
        try:
            # Extrair o título
            title_element = product.find_element(By.CLASS_NAME, 'x-shelf__title')
            product_title = title_element.text.strip()

            # Verificar se o título não está vazio
            if not product_title:
                continue

            # Extrair a URL da imagem frontal
            try:
                product_front_image_url = product.find_element(By.CLASS_NAME, 'x-shelf__img-front').get_attribute('src')
            except:
                product_front_image_url = 'Imagem frontal não encontrada'

            # Extrair o preço original
            try:
                original_price_element = product.find_element(By.CLASS_NAME, 'x-shelf__best-price')
                product_original_price = original_price_element.text
            except:
                product_original_price = 'Preço original não encontrado'

            # Extrair as tags (se existirem)
            product_tags = []
            try:
                if product.find_element(By.CLASS_NAME, 'night-sale'):
                    product_tags.append('Night Sale')
            except:
                pass
            try:
                if product.find_element(By.CLASS_NAME, 'sestini-lancamentos'):
                    product_tags.append('Lançamento')
            except:
                pass

            # Armazenar as informações do produto
            products_data.append({
                'Título': product_title,
                'URL da Imagem Frontal': product_front_image_url,
                'Preço Original': product_original_price,
                'Tags': ', '.join(product_tags) if product_tags else 'Nenhuma tag encontrada'
            })

        except Exception as e:
            print(f'Erro ao extrair informações de um produto: {e}')

# Iterar sobre as páginas
for page in range(1, 6):
    url = f'https://www.sestini.com.br/mala?map=c#{page}'
    driver.get(url)
    request_count += 1

    # Aguardar carregar a página
    wait = WebDriverWait(driver, 6)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

    # Extrair informações dos produtos
    extract_product_info()

# Fechar o navegador
driver.quit()

# Criar um DataFrame a partir dos dados dos produtos
df = pd.DataFrame(products_data)

# Salvar o DataFrame em um arquivo Excel
df.to_excel('informacoes_produtos.xlsx', index=False)

# Calcular o tempo total
end_time = time.time()
total_time = end_time - start_time

# Aviso sobre o número de requisições e o tempo total
print(f'Número total de requisições feitas: {request_count}')
print(f'Tempo total gasto: {total_time:.2f} segundos')

# Informar quantos produtos diferentes foram encontrados
unique_product_titles = set(df['Título'])
print(f'Número total de produtos diferentes encontrados: {len(unique_product_titles)}')

