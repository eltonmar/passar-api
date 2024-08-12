import requests
from openpyxl import Workbook
from datetime import datetime, timedelta
import time

def get_tokens():
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

        return access_token, refresh_token, token_type, expires_in, scope
    else:
        print("Falha na obtenção do token. Status code:", response.status_code)
        return None, None, None, None, None

def refresh_access_token(refresh_token):
    url = "https://api.userede.com.br/redelabs/oauth/token"
    body = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token
    }

    headers = {
        "Authorization": "Basic N2I3OWIyNjUtNjFjMi00YmJiLThlNmItZGE2NDNjMDliMThiOjI3bWJXMnpDeFE="
    }

    response = requests.post(url, data=body, headers=headers)

    if response.status_code == 200:
        data = response.json()
        access_token = data.get("access_token", "")
        refresh_token = data.get("refresh_token", "")

        return access_token, refresh_token
    else:
        print("Falha na atualização do token. Status code:", response.status_code)
        return None, None

def fetch_data(url, params, headers, retries=3, delay=5):
    for i in range(retries):
        try:
            response = requests.get(url, params=params, headers=headers, timeout=10)
            response.raise_for_status()
            return response
        except requests.exceptions.RequestException as e:
            print(f"Falha ao obter dados, tentativa {i + 1}/{retries}: {e}")
            time.sleep(delay)
    return None

def validate_params(start_date_str, end_date_str, company_number):
    try:
        datetime.strptime(start_date_str, "%Y-%m-%d")
        datetime.strptime(end_date_str, "%Y-%m-%d")
    except ValueError:
        print("Data no formato incorreto. Use 'YYYY-MM-DD'.")
        return False

    if not isinstance(company_number, int):
        print("Número da empresa deve ser um inteiro.")
        return False

    return True

def main():
    access_token, refresh_token, _, _, _ = get_tokens()

    if not access_token or not refresh_token:
        print("Falha na obtenção do token.")
        return

    url = "https://api.userede.com.br/redelabs/merchant-statement/v1/sales"

    companyNumbers = [
        92117791
    ]

    data_to_write = []

    start_date_str = "2024-07-01"
    end_date_str = "2024-07-02"

    if not validate_params(start_date_str, end_date_str, companyNumbers[0]):
        return

    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d")

    for companyNumber in companyNumbers:
        current_date = start_date

        while current_date <= end_date:
            date_str = current_date.strftime("%Y-%m-%d")
            params = {
                "startDate": date_str,
                "endDate": date_str,
                "parentCompanyNumber": companyNumber,
                "size": 100
            }

            headers_request = {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + access_token
            }

            response = fetch_data(url, params, headers_request)

            if response is not None:
                content = response.json().get('content', {})
                transactions = content.get('transactions', [])
                for transaction in transactions:
                    row_data = {
                        "companyNumber": transaction['merchant'].get("companyNumber"),
                        "documentNumber": transaction['merchant'].get("documentNumber"),
                        "nsu": transaction.get("nsu"),
                        "saleDate": transaction.get("saleDate"),
                        "deviceType": transaction.get("deviceType"),
                        "device": transaction.get("device")
                    }
                    data_to_write.append(row_data)
            else:
                print(f"Falha ao obter dados para o número da empresa {companyNumber} no dia {date_str} após múltiplas tentativas.")

            current_date += timedelta(days=1)

    # Criar workbook e worksheet
    workbook = Workbook()
    sheet = workbook.active

    # Adicionar cabeçalhos
    headers_excel = ["companyNumber", "documentNumber", "nsu", "saleDate", "deviceType", "device"]
    sheet.append(headers_excel)

    # Adicionar dados à planilha
    for row_data in data_to_write:
        row = [
            row_data.get("companyNumber"),
            row_data.get("documentNumber"),
            row_data.get("nsu"),
            row_data.get("saleDate"),
            row_data.get("deviceType"),
            row_data.get("device")
        ]
        sheet.append(row)

    # Especificar o caminho do arquivo Excel
    excel_file_path = r"C:\Users\BG-PROVISORIO\Desktop\Teste-GestaoDeVendas.xlsx"

    # Salvar arquivo Excel
    workbook.save(excel_file_path)
    print("Dados escritos no arquivo Excel com sucesso!")

if __name__ == "__main__":
    main()
