import requests
from requests.auth import HTTPBasicAuth
import pandas as pd

url = "https://bagaggio.zendesk.com/api/v2/tickets"
headers = {
    "Content-Type": "application/json",
}
email_address = 'wpp.sac@bagaggio.com.br'
api_token = 'CUn5NJ5enSkBQF0vbUIsNCl9NFUCgy7aYsSJqUFG'
# Use basic authentication
auth = HTTPBasicAuth(f'{email_address}/token', api_token)

response = requests.request(
    "GET",
    url,
    auth=auth,
    headers=headers
)

# Verifique se a resposta está correta e se o conteúdo é JSON
if response.status_code == 200:
    data = response.json()

    # Verifique se a chave 'tickets' está presente
    if 'tickets' in data:
        tickets_data = data['tickets']

        # Seleciona apenas as colunas desejadas
        columns_of_interest = [
            'id', 'external_id', 'created_at', 'updated_at', 'generated_timestamp',
            'type', 'priority', 'status', 'recipient', 'requester_id', 'submitter_id',
            'assignee_id', 'organization_id', 'group_id', 'collaborator_ids', 'follower_ids',
            'email_cc_ids', 'forum_topic_id', 'problem_id', 'has_incidents', 'is_public',
            'due_at', 'sharing_agreement_ids', 'custom_status_id',
            'followup_ids', 'ticket_form_id', 'brand_id', 'allow_channelback',
            'allow_attachments', 'from_messaging_channel'
        ]

        # Converta os dados para um DataFrame do pandas, selecionando apenas as colunas desejadas
        df = pd.DataFrame(tickets_data, columns=columns_of_interest)

        # Escreva os dados no arquivo Excel existente
        with pd.ExcelWriter("C:\\Users\\BG-PROVISORIO\\Desktop\\Talita.xlsx", engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
            df.to_excel(writer, index=False, sheet_name='API_Tickets')

        print("Dados salvos no arquivo Talita.xlsx")
    else:
        print("A chave 'tickets' não está presente na resposta.")
else:
    print(f"Falha na solicitação: {response.status_code}")
