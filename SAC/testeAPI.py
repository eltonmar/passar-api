import requests
from requests.auth import HTTPBasicAuth
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

driver = "ODBC Driver 17 for SQL Server"
server = "187.0.198.167"
user = "victor.oliveira"
password = "@primo01"
database = "DADOS_EXCEL"
port = 41433

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
    '7896616478612',   # Assunto do Email
    '360041469032',    # Canal de Entrada
    '360041468692',    # Dúvida
    '360041432051',    # Solicitação
    '360041431951',    # Problema
    '360041432091',    # Outros
    '22541325',        # Transportadora
    '8225162131348',   # Produto
    '360041040172',    # Número do Pedido
    '360030577731',    # SKU dos produtos
    '360040274491',    # Número da NF
    '23507539076884',  # Estorno: valor
    '23465090667540',  # Tipo de estorno
    '24157626991892',  # Atendente
    '360030496932',    # Nome Titular do Pedido
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
    '25427606175380'   # Número da NFD
]

def create_connection(driver, server, database, user, password, port):
    connection = None
    try:
        connection = pyodbc.connect(
            f'DRIVER={{{driver}}};'
            f'SERVER={server},{port};'
            f'DATABASE={database};'
            f'UID={user};'
            f'PWD={password}'
        )
        print("Connection to SQL Server successful")
    except pyodbc.Error as e:
        print(f"The error '{e}' occurred")
    return connection

def truncate_table(connection):
    try:
        with connection.cursor() as cursor:
            cursor.execute("TRUNCATE TABLE BD_RECEBIVEIS_SEMANAL")
            connection.commit()
            print("Tabela foi truncada com sucesso.")
    except pyodbc.Error as e:
        print(f"O erro foi: {e}")

def insert_data(connection, data_row):
    insert_query = """
    INSERT INTO BD_SAC (
    Area_retorno, Data_de_envio_Area_responsável, Previsao_de_retorno_Area_responsavel, Assunto_do_Email, 
    Canal_de_Entrada, Duvida, Solicitacao, Problema, Outros, Transportadora, Produto, Numero_do_Pedido, 
    SKU_dos_produtos, Numero_da_NF, "Estorno:valor", Tipo_de_estorno, Atendente, Nome_Titular_do_Pedido, 
    "Estorno:causa_raiz", "Estorno:tipo_de_problema", "Estorno:tipo_de_pagamento", Status_da_coleta, 
    "CD:Troca", Coleta_mais_de_uma_vez, Caso_foi_resolvido, Numero_da_Loja, "CD:Outras_demandas", 
    "Replica?", Sentimento, Status_de_assistencia_tecnica, Plano_OS_vencidas, "1_cobranca", "CD:Devolucao", 
    Loja_Fisica_ou_Virtual, Etapas_de_coleta, Avaliacao, Nota_da_avaliacao, Demanda, Plano_ação, 
    Número_da_OS, Cliente_Reincidente, Número_NFD, Assunto, Descricao, Grupo, Atribuido_para, Status
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
     ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    cursor = connection.cursor()
    cursor.execute(insert_query, data_row)
    connection.commit()


def fetch_first_page_tickets(url, auth, headers):
    response = requests.get(url, auth=auth, headers=headers)
    if response.status_code != 200:
        print(f"Falha na solicitação: {response.status_code}")
        print(f"Mensagem de erro: {response.text}")
        return []

    data = response.json()
    tickets = data.get('tickets', [])
    return tickets

def job():
    connection = create_connection(driver, server, database, user, password, port)
    if not connection:
        print("Não foi possível estabelecer a conexão.")
        return

    truncate_table(connection)

    tickets_data = fetch_first_page_tickets(url, auth, headers)

    for ticket in tickets_data:
        filtered_fields = {field['id']: field['value'] for field in ticket['custom_fields'] if
                           str(field['id']) in custom_field_ids}

        for field_id, field_name in standard_field_ids.items():
            filtered_fields[field_name] = ticket.get('fields', {}).get(field_id)

        data_row = (
            filtered_fields.get('Área retorno', ''),
            filtered_fields.get('Data de envio Área responsável', ''),
            filtered_fields.get('Previsão de retorno Área responsável', ''),
            filtered_fields.get('Assunto do Email', ''),
            filtered_fields.get('Canal de Entrada', ''),
            filtered_fields.get('Dúvida', ''),
            filtered_fields.get('Solicitação', ''),
            filtered_fields.get('Problema', ''),
            filtered_fields.get('Outros', ''),
            filtered_fields.get('Transportadora', ''),
            filtered_fields.get('Produto', ''),
            filtered_fields.get('Número do Pedido', ''),
            filtered_fields.get('SKU dos produtos', ''),
            filtered_fields.get('Número da NF', ''),
            filtered_fields.get('Estorno: valor', ''),
            filtered_fields.get('Tipo de estorno', ''),
            filtered_fields.get('Atendente', ''),
            filtered_fields.get('Nome Titular do Pedido', ''),
            filtered_fields.get('Estorno: causa raiz', ''),
            filtered_fields.get('Estorno: tipo de problema', ''),
            filtered_fields.get('Estorno: tipo de pagamento', ''),
            filtered_fields.get('Status da coleta', ''),
            filtered_fields.get('CD: Troca e Acionamento de Garantia', ''),
            filtered_fields.get('Coleta foi solicitada mais de uma vez?', ''),
            filtered_fields.get('O caso foi 100% resolvido no atendimento anterior?', ''),
            filtered_fields.get('Número da Loja', ''),
            filtered_fields.get('CD: Outras demandas', ''),
            filtered_fields.get('Réplica?', ''),
            filtered_fields.get('Sentimento', ''),
            filtered_fields.get('Status de assistência técnica', ''),
            filtered_fields.get('Plano de ação OS vencidas', ''),
            filtered_fields.get('Prazo 1ª cobrança', ''),
            filtered_fields.get('CD: Devolução e Voucher', ''),
            filtered_fields.get('Loja Física ou Loja Virtual', ''),
            filtered_fields.get('Etapas de coleta', ''),
            filtered_fields.get('Avaliação no RA?', ''),
            filtered_fields.get('Nota da avaliação', ''),
            filtered_fields.get('Demanda', ''),
            filtered_fields.get('Plano de ação insatisfação resultado de OS', ''),
            filtered_fields.get('Número da OS', ''),
            filtered_fields.get('Cliente Reincidente?', ''),
            filtered_fields.get('Número da NFD', ''),
            filtered_fields.get('Assunto', ''),
            filtered_fields.get('Descrição', ''),
            filtered_fields.get('Grupo', ''),
            filtered_fields.get('Atribuido para', ''),
            filtered_fields.get('Status', '')
        )

        insert_data(connection, data_row)

    connection.close()

# Agenda o job para rodar diariamente
schedule.every().day.at("16:22").do(job)

# Loop para manter o schedule rodando
while True:
    schedule.run_pending()
    time.sleep(0.5)

if __name__ == "__main__":
    job()
