import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import time
import re

def remove_illegal_characters(value):
    if isinstance(value, str):
        # Remove caracteres não imprimíveis
        return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', value)
    return value

url = "https://bagaggio.zendesk.com/api/v2/tickets.json?page[size]=100"
headers = {
    "Content-Type": "application/json",
}
email_address = 'wpp.sac@bagaggio.com.br'
api_token = 'CUn5NJ5enSkBQF0vbUIsNCl9NFUCgy7aYsSJqUFG'
auth = HTTPBasicAuth(f'{email_address}/token', api_token)

def fetch_all_tickets(url, auth, headers, retries=5, backoff_factor=1):
    all_tickets = []
    while url:
        for attempt in range(retries):
            try:
                response = requests.get(url, auth=auth, headers=headers)
                response.raise_for_status()
                data = response.json()
                all_tickets.extend(data.get('tickets', []))
                url = data.get('links', {}).get('next')
                print(f"Próxima página: {url}")
                break
            except (requests.exceptions.RequestException, ConnectionResetError) as e:
                wait_time = backoff_factor * (2 ** attempt)
                print(f"Erro na conexão: {e}. Tentando novamente em {wait_time} segundos...")
                time.sleep(wait_time)
        else:
            print(f"Falha após {retries} tentativas. Parando a execução.")
            break
    return all_tickets

tickets_data = fetch_all_tickets(url, auth, headers)

custom_field_ids = [
    '20481751634964',  # Área retorno
    '23450471389460',  # Data de envio Área responsável
    '23450335909780',  # Previsão de retorno Área responsável
    '7896616478612',  # Assunto do Email
    '360041469032',  # Canal de Entrada
    '360041468692',  # Dúvida
    '360041432051',  # Solicitação
    '360041431951',  # Problema
    '360041432091',  # Outros
    '22541325',  # Transportadora
    '8225162131348',  # Produto
    '360041040172',  # Número do Pedido
    '360030577731',  # SKU dos produtos
    '360040274491',  # Número da NF
    '23507539076884',  # Estorno: valor
    '23465090667540',  # Tipo de estorno
    '24157626991892',  # Atendente
    '360030496932',  # Nome Titular do Pedido
    '23555735385236',  # Estorno: causa raiz
    '23555716189844',  # Estorno: tipo de problema
    '25219880343316',  # Estorno: tipo de pagamento
    '27112346684948',  # Status da coleta
    '25783014985492',  # CD: Troca e Acionamento de Garantia
    '27112338364436',  # Coleta foi solicitada mais de uma vez?
    '27265259806228',  # O caso foi 100% resolvido no atendimento anterior?
    '26678660208916',  # Número da Loja
    '25907732988436',  # CD: Outras demandas
    '27112064079636',  # Réplica?
    '28405635340308',  # Sentimento
    '26241507056916',  # Status de assistência técnica
    '26256563363348',  # Plano de ação OS vencidas
    '26241374621588',  # Prazo 1ª cobrança
    '25808063108756',  # CD: Devolução e Voucher
    '27112048306068',  # Loja Física ou Loja Virtual
    '25780172368020',  # Etapas de coleta
    '27112103294868',  # Avaliação no RA?
    '27112199178132',  # Nota da avaliação
    '25820195084948',  # Demanda
    '26256620215444',  # Plano de ação insatisfação resultado de OS
    '25966692319380',  # Número da OS
    '27265194513556',  # Cliente Reincidente?
    '25427606175380'  # Número da NFD
        'ticket_id' #'Ticket ID'
]
filtered_data = []

# Filtrar os custom fields dos tickets e adicionar o ticket_id
for ticket in tickets_data:
    filtered_fields = {'ticket_id': ticket['id']}  # Inicializa com o ticket_id
    custom_fields = {field['id']: remove_illegal_characters(field['value'])
                     for field in ticket['custom_fields']
                     if str(field['id']) in custom_field_ids}
    filtered_fields.update(custom_fields)
    filtered_data.append(filtered_fields)

# Converta os dados filtrados para um DataFrame do pandas
df = pd.DataFrame(filtered_data)

# Escreva os dados no arquivo Excel existente
with pd.ExcelWriter("C:\\Users\\BG-PROVISORIO\\Desktop\\Talita\\Testando com a Pagination.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, index=False, sheet_name='API_Tickets')

print("Dados salvos no arquivo.")




