import requests
from openpyxl import Workbook
import datetime

url = "https://api.userede.com.br/redelabs/oauth/token"
body = {
    "grant_type": "password",
    "username": "elton.marinho@bagaggio.com.br",
    "password": ":5@A9hPr9-Po"
}

headers = {
    "Authorization": "Basic N2I3OWIyNjUtNjFjMi00YmJiLThlNmItZGE2NDNjMDliMThiOjI3bWJXMnpDeFE="
}

response = requests.post(url, data=body, headers=headers)

if response.status_code == 200:
    data = response.json()
    access_token = data.get("access_token", "")
    refresh_token = data.get("refresh_token", "")
    token_type = data.get("token_type", "")
    expires_in = data.get("expires_in", "")
    scope = data.get("scope", "")

url = "https://api.userede.com.br/redelabs/merchant-statement/v2/receivables/summary"

companyNumbers = [
    95367160, 82445265, 94572593, 95008845, 91489520, 32145330
]

headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + access_token
}

data = datetime.datetime.now()

lista_dias = [(data + datetime.timedelta(days=i)).day for i in range(6)]

data_to_write = []

for day in lista_dias:
    startdate = f"{data.year}-{data.month:02d}-{day:02d}"  # Garantindo dois dígitos para mês e dia
    enddate = f"{data.year}-{data.month:02d}-{day:02d}"  # Garantindo dois dígitos para mês e dia

    for companyNumber in companyNumbers:
        params = {
            "startDate": startdate,
            "endDate": enddate,
            "parentCompanyNumber": companyNumber
        }
        response = requests.get(url, params=params, headers=headers)
        if response.status_code == 200:
            content = response.json()['content']
            if content:
                amount = content[0]['amount']
                total = content[0]['total']
                data_to_write.append((startdate, enddate, companyNumber, amount, total))
                print("Valores inseridos para a empresa", companyNumber)
            elif response.status_code == 204:  # Correção da indentação aqui
                data_to_write.append((startdate, enddate, companyNumber, 0, 0))
                print("Inserido vazio para a empresa", companyNumber)
        else:
            print(f"Erro para a empresa {companyNumber}: {response.status_code}")

# Agora que todos os dados foram coletados, vão para o Excel
excel_file_path = r"C:\Users\BG-PROVISORIO\Desktop\Teste-Recebiveis.xlsx"
workbook = Workbook()
sheet = workbook.active

for row_data in data_to_write:
    sheet.append(row_data)
workbook.save(excel_file_path)
print("Dados escritos no arquivo Excel com sucesso!")
