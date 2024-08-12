import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
from datetime import datetime, timedelta

base_url = "https://bagaggio.zendesk.com/api/v2/tickets"
headers = {
    "Content-Type": "application/json",
}
email_address = 'wpp.sac@bagaggio.com.br'
api_token = 'CUn5NJ5enSkBQF0vbUIsNCl9NFUCgy7aYsSJqUFG'
# Use basic authentication
auth = HTTPBasicAuth(f'{email_address}/token', api_token)

all_tickets = []

# Função para gerar intervalos de datas
def generate_date_ranges(start_date, end_date, delta_days):
    current_date = start_date
    while current_date < end_date:
        next_date = current_date + timedelta(days=delta_days)
        yield current_date, next_date
        current_date = next_date

start_date = datetime(2023, 1, 1)
end_date = datetime(2024, 1, 1)
delta_days = 30  # Intervalo de 30 dias

for start, end in generate_date_ranges(start_date, end_date, delta_days):
    next_page = f"{base_url}?start_time={int(start.timestamp())}&end_time={int(end.timestamp())}"
    while next_page:
        try:
            response = requests.request(
                "GET",
                next_page,
                auth=auth,
                headers=headers,
                timeout=30  # Ajuste o tempo limite para 30 segundos
            )

            # Verifique se a resposta está correta e se o conteúdo é JSON
            if response.status_code == 200:
                data = response.json()

                # Verifique se a chave 'tickets' está presente
                if 'tickets' in data:
                    all_tickets.extend(data['tickets'])
                    next_page = data['next_page']
                else:
                    print("A chave 'tickets' não está presente na resposta.")
                    next_page = None
            else:
                print(f"Falha na solicitação: {response.status_code}")
                next_page = None
        except requests.exceptions.ReadTimeout:
            print("A solicitação excedeu o tempo limite.")
            next_page = None
        except requests.exceptions.RequestException as e:
            print(f"Ocorreu um erro na solicitação: {e}")
            next_page = None

# Se houver tickets, converta-os para um DataFrame do pandas
if all_tickets:
    # Seleciona apenas as colunas desejadas
    columns_of_interest = [
        'id', 'custom_fields'
    ]

    # Converta os dados para um DataFrame do pandas, selecionando apenas as colunas desejadas
    df = pd.DataFrame(all_tickets)

    # Verifique se as colunas de interesse existem no DataFrame
    df = df[columns_of_interest]

    # Escreva os dados no arquivo Excel existente
    with pd.ExcelWriter("C:\\Users\\BG-PROVISORIO\\Desktop\\Talita\\Talita.xlsx", engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
        df.to_excel(writer, index=False, sheet_name='API_Tickets')

    print("Dados salvos no arquivo Talita.xlsx")
else:
    print("Nenhum ticket encontrado.")
