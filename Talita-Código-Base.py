import requests
from requests.auth import HTTPBasicAuth
import pandas as pd

url = "https://example.zendesk.com/api/v2/ticket_fields"
headers = {
    "Content-Type": "application/json",
}
email_address = 'wpp.sac@bagaggio.com.br'
api_token = 'CUn5NJ5enSkBQF0vbUIsNCl9NFUCgy7aYsSJqUFG'
auth = HTTPBasicAuth(f'{email_address}/token', api_token)

response = requests.request(
	"GET",
	url,
	auth=auth,
	headers=headers
)

print(response.text)

"""
# Escreva os dados no arquivo Excel existente
with pd.ExcelWriter("C:\\Users\\BG-PROVISORIO\\Desktop\\Talita\\TesteBasico.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, index=False, sheet_name='API_Tickets')

print("Dados salvos no arquivo.")
"""
