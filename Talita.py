import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import pyodbc
import schedule
import time

url = "https://bagaggio.zendesk.com/api/v2/tickets.json?page[size]=100"
headers = {
    "Content-Type": "application/json",
}
email_address = 'wpp.sac@bagaggio.com.br'
api_token = 'CUn5NJ5enSkBQF0vbUIsNCl9NFUCgy7aYsSJqUFG'
auth = HTTPBasicAuth(f'{email_address}/token', api_token)


def fetch_all_tickets(url, auth, headers):
    all_tickets = []
    while url:
        response = requests.get(url, auth=auth, headers=headers)
        if response.status_code != 200:
            print(f"Falha na solicitação: {response.status_code}")
            print(f"Mensagem de erro: {response.text}")
            break
        data = response.json()
        all_tickets.extend(data.get('tickets', []))
        url = data.get('links', {}).get('next')
        print(f"Próxima página: {url}")
    return all_tickets


# IDs dos campos padrão que você mencionou
standard_field_ids = {
    '22333195': 'Assunto',
    '22333205': 'Descrição',
    '22333245': 'Grupo',
    '22333255': 'Atribuido para',
    '9647033755156': 'Status'
}

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
]

tickets_data = fetch_all_tickets(url, auth, headers)

filtered_data = []

for ticket in tickets_data:
    filtered_fields = {field['id']: field['value'] for field in ticket['custom_fields'] if
                       str(field['id']) in custom_field_ids}

    # Adiciona os campos padrão ao dicionário
    for field_id, field_name in standard_field_ids.items():
        filtered_fields[field_name] = ticket.get('fields', {}).get(field_id)

    filtered_data.append(filtered_fields)

df = pd.DataFrame(filtered_data)
'3'

print(df)
