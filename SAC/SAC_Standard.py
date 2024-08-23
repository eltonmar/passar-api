import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import time
import re

# Configurações da API e autenticação
url = "https://bagaggio.zendesk.com/api/v2/tickets.json?page[size]=100"
headers = {
    "Content-Type": "application/json",
}
email_address = 'wpp.sac@bagaggio.com.br'
api_token = 'CUn5NJ5enSkBQF0vbUIsNCl9NFUCgy7aYsSJqUFG'
auth = HTTPBasicAuth(f'{email_address}/token', api_token)

# Função para buscar todos os tickets da API
def fetch_all_tickets(url, auth, headers, retries=5, backoff_factor=1):
    all_tickets = []
    while url:
        for attempt in range(retries):
            try:
                response = requests.get(url, auth=auth, headers=headers)
                response.raise_for_status()  # Verifica se houve erro na requisição
                data = response.json()
                all_tickets.extend(data.get('tickets', []))
                url = data.get('links', {}).get('next')  # Pega a próxima página, se houver
                print(f"Próxima página: {url}")
                break  # Se a requisição foi bem-sucedida, sai do loop de tentativas
            except (requests.exceptions.RequestException, ConnectionResetError) as e:
                wait_time = backoff_factor * (2 ** attempt)  # Tempo de espera exponencial
                print(f"Erro na conexão: {e}. Tentando novamente em {wait_time} segundos...")
                time.sleep(wait_time)
        else:
            print(f"Falha após {retries} tentativas. Parando a execução.")
            break
    return all_tickets

# Função para remover caracteres ilegais
def remove_illegal_characters(text):
    if isinstance(text, str):
        # Remove caracteres não suportados pelo Excel
        return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', text)
    return text

# Buscar dados dos tickets
tickets_data = fetch_all_tickets(url, auth, headers)

# Definição dos campos padrão de interesse e seus nomes na API do Zendesk
standard_field_ids = {
    'Assunto': 'subject',
    'Descrição': 'description',
    'Grupo': 'group_id',
    'Atribuído para': 'assignee_id',
    'Status': 'status',
    'Ticket ID': 'id'
}

# IDs dos campos customizados que queremos extrair
custom_field_ids = ['123456789', '987654321']  # Substitua pelos IDs reais dos campos customizados

# Lista para armazenar os dados filtrados
filtered_data = []

# Filtrar os campos padrão e customizados dos tickets e adicionar ao DataFrame
for ticket in tickets_data:
    # Extrair os campos padrão do ticket usando o mapeamento correto
    filtered_fields = {field_id: remove_illegal_characters(ticket.get(field_name, '')) for field_id, field_name in standard_field_ids.items()}

    # Extrair os campos customizados, se presentes
    custom_fields = {field['id']: remove_illegal_characters(field['value']) for field in ticket.get('custom_fields', []) if str(field['id']) in custom_field_ids}

    # Adicionar os campos customizados aos campos padrão
    filtered_fields.update(custom_fields)

    # Adicionar o ticket filtrado à lista de dados
    filtered_data.append(filtered_fields)

# Converter os dados filtrados para um DataFrame do pandas
df = pd.DataFrame(filtered_data)

# Escrever os dados em um arquivo Excel existente
with pd.ExcelWriter("C:\\Users\\BG-PROVISORIO\\Desktop\\Talita\\TesteStandard.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, index=False, sheet_name='API_Tickets')

print("Dados salvos no arquivo.")


