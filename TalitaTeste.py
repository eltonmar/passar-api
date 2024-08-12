import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import pymysql
import time

# Função para criar a conexão com o banco de dados com tentativa de reconexão
def create_db_connection():
    attempts = 0
    max_attempts = 3
    while attempts < max_attempts:
        try:
            connection = pymysql.connect(
                host='187.0.198.167',
                port=41433,
                user='victor.oliveira',
                password='@primo01',
                database='DADOS_EXCEL'
            )
            return connection
        except pymysql.Error as e:
            print(f"Erro ao conectar ao banco de dados: {e}")
            attempts += 1
            if attempts < max_attempts:
                print(f"Tentando reconectar (tentativa {attempts} de {max_attempts})...")
                time.sleep(5)  # Espere 5 segundos antes de tentar novamente
            else:
                raise  # Se exceder o número máximo de tentativas, levante a exceção

# Função para inserir dados no banco de dados com limitação de caracteres
def insert_data(cursor, data):
    insert_query = """
    INSERT INTO tickets (url, id, external_id, via, created_at, updated_at, generated_timestamp, type, subject, raw_subject,
    description, priority, status, recipient, requester_id, submitter_id, assignee_id, organization_id, group_id,
    collaborator_ids, follower_ids, email_cc_ids, forum_topic_id, problem_id, has_incidents, is_public, due_at, tags,
    custom_fields, satisfaction_rating, sharing_agreement_ids, custom_status_id, fields, followup_ids, ticket_form_id,
    brand_id, allow_channelback, allow_attachments, from_messaging_channel)
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """
    for index, row in data.iterrows():
        # Truncate values to 60 characters if they exceed
        truncated_values = tuple(str(value)[:60] if isinstance(value, str) else value for value in row)

        cursor.execute(insert_query, truncated_values)

# URL da API e credenciais
url = "https://bagaggio.zendesk.com/api/v2/tickets"
email_address = 'wpp.sac@bagaggio.com.br'
api_token = 'CUn5NJ5enSkBQF0vbUIsNCl9NFUCgy7aYsSJqUFG'

# Cabeçalhos da requisição
headers = {
    "Content-Type": "application/json",
}

# Autenticação básica
auth = HTTPBasicAuth(f'{email_address}/token', api_token)

# Fazendo a requisição GET
response = requests.request(
    "GET",
    url,
    auth=auth,
    headers=headers
)

# Verificando se a requisição foi bem sucedida
if response.status_code == 200:
    data = response.json()

    # Verificando se a chave 'tickets' está presente na resposta
    if 'tickets' in data:
        tickets_data = data['tickets']
        df = pd.DataFrame(tickets_data)

        # Conectando ao banco de dados com tentativa de reconexão
        connection = create_db_connection()
        cursor = connection.cursor()

        # Inserindo os dados no banco de dados
        insert_data(cursor, df)

        # Confirmar as alterações
        connection.commit()

        # Fechar a conexão
        cursor.close()
        connection.close()

        print("Dados inseridos no banco de dados com sucesso.")
    else:
        print("A chave 'tickets' não está presente na resposta.")
else:
    print(f"Falha na solicitação: {response.status_code}")
