import requests
from requests.auth import HTTPBasicAuth
import pandas as pd

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

tickets_data = fetch_all_tickets(url, auth, headers)

# IDs dos custom fields de interesse
custom_field_ids = [
    '22333195',        # Assunto
    '22333205',        # Descrição
    '22333245',        # Grupo
    '22333255',        # Atribuído para
    '9647033755156',   # Ticket status
    '20451751200',     # Área retorno
    '2345071389460',   # Data de envio Área responsável
    '2345035909780',   # Previsão de retorno Área responsável
    '7896616478612',   # Assunto do Email
    '360041468692',    # Dúvida
    '360041432051',    # Solicitação
    '360041431951',    # Problema
    '360041432091',    # Outros
    '22541325',        # Transportadora
    '8225162131348',   # Produto
    '360041404172',    # Número do Pedido
    '360030577731',    # SKU dos produtos
    '360040274491',    # Número da NF
    '23507539076884',  # Estorno: valor
    '2345060967540',   # Tipo de estorno
    '2415769622903',   # Atendente
    '360030496932',    # Nome Titular do Pedido
    '2355753385236',   # Estorno: causa raiz
    '23557511689844',  # Estorno: tipo de problema
    '25219880343316',  # Estorno: tipo de pagamento
    '27112346689448',  # Status da coleta
    '25783014985492',  # CD: Troca e Acionamento de Garantia
    '27112338364436',  # Coleta foi solicitada mais de uma vez?
    '27265258906228',  # O caso foi 100% resolvido no atendimento anterior?
    '26676862089848',  # Número da Loja
    '25907732988436',  # CD: Outras demandas
    '27112064079636',  # Réplica?
    '26251470659619',  # Status de assistência técnica
    '26256563333498',  # Plano de ação OS vencidas
    '2624137216588',   # Prazo 1ª cobrança
    '25907732988436',  # CD: Devolução e Voucher
    '25907732988436',  # Loja Física ou Loja Virtual
    '25781017832492',  # Etapas de coleta
    '27266858906228',  # Avaliação no RA?
    '25781017832492',  # Nota da avaliação
    '25781017832492',  # Demanda
    '26256563333498',  # Plano de ação insatisfação resultado de OS
    '25969563333498',  # Número da OS
    '25781017832492',  # Cliente Reincidente?
    '25427606157380'   # Número da NFD
]

# Lista para armazenar os dados filtrados
filtered_data = []

# Filtrar os custom fields dos tickets e adicionar o ticket_id
for ticket in tickets_data:
    filtered_fields = {field['id']: field['value'] for field in ticket['custom_fields'] if str(field['id']) in custom_field_ids}
    filtered_fields['ticket_id'] = ticket['id']
    filtered_data.append(filtered_fields)

# Converta os dados filtrados para um DataFrame do pandas
df = pd.DataFrame(filtered_data)

# Renomeie as colunas para nomes amigáveis se necessário
column_mapping = {
    '22333195': 'Assunto',
    '22333205': 'Descrição',
    '22333245': 'Grupo',
    '22333255': 'Atribuído para',
    '9647033755156': 'Ticket status',
    '20451751200': 'Área retorno',
    '2345071389460': 'Data de envio Área responsável',
    '2345035909780': 'Previsão de retorno Área responsável',
    '7896616478612': 'Assunto do Email',
    '360041468692': 'Dúvida',
    '360041432051': 'Solicitação',
    '360041431951': 'Problema',
    '360041432091': 'Outros',
    '22541325': 'Transportadora',
    '8225162131348': 'Produto',
    '360041404172': 'Número do Pedido',
    '360030577731': 'SKU dos produtos',
    '360040274491': 'Número da NF',
    '23507539076884': 'Estorno: valor',
    '2345060967540': 'Tipo de estorno',
    '2415769622903': 'Atendente',
    '360030496932': 'Nome Titular do Pedido',
    '2355753385236': 'Estorno: causa raiz',
    '23557511689844': 'Estorno: tipo de problema',
    '25219880343316': 'Estorno: tipo de pagamento',
    '27112346689448': 'Status da coleta',
    '25783014985492': 'CD: Troca e Acionamento de Garantia',
    '27112338364436': 'Coleta foi solicitada mais de uma vez?',
    '27265258906228': 'O caso foi 100% resolvido no atendimento anterior?',
    '26676862089848': 'Número da Loja',
    '25907732988436': 'CD: Outras demandas',
    '27112064079636': 'Réplica?',
    '26251470659619': 'Status de assistência técnica',
    '26256563333498': 'Plano de ação OS vencidas',
    '2624137216588': 'Prazo 1ª cobrança',
    '25907732988436': 'CD: Devolução e Voucher',
    '25907732988436': 'Loja Física ou Loja Virtual',
    '25781017832492': 'Etapas de coleta',
    '27266858906228': 'Avaliação no RA?',
    '25781017832492': 'Nota da avaliação',
    '25781017832492': 'Demanda',
    '26256563333498': 'Plano de ação insatisfação resultado de OS',
    '25969563333498': 'Número da OS',
    '25781017832492': 'Cliente Reincidente?',
    '25427606157380': 'Número da NFD',
    'ticket_id': 'Ticket ID'
}

df.rename(columns=column_mapping, inplace=True)

# Escreva os dados no arquivo Excel existente
with pd.ExcelWriter("C:\\Users\\BG-PROVISORIO\\Desktop\\Talita\\Testando com a Pagination.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, index=False, sheet_name='API_Tickets')

print("Dados salvos no arquivo.")
