import requests
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

    try:
        response = requests.post(url, data=body, headers=headers, timeout=10)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print("Erro ao obter o token:", e)
        return None, None

    data = response.json()
    access_token = data.get("access_token", "")
    refresh_token = data.get("refresh_token", "")
    return access_token, refresh_token


def refresh_access_token(refresh_token):
    url = "https://api.userede.com.br/redelabs/oauth/token"
    body = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token
    }

    headers = {
        "Authorization": "Basic N2I3OWIyNjUtNjFjMi00YmJiLThlNmItZGE2NDNjMDliMThiOjI3bWJXMnpDeFE="
    }

    try:
        response = requests.post(url, data=body, headers=headers, timeout=10)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print("Erro ao atualizar o token:", e)
        return None, None

    data = response.json()
    access_token = data.get("access_token", "")
    refresh_token = data.get("refresh_token", "")
    return access_token, refresh_token


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


def main():
    access_token, refresh_token = get_tokens()

    if not access_token or not refresh_token:
        print("Falha na obtenção do token.")
        return

    companyNumbers = [94959056]

    url = "https://api.userede.com.br/redelabs/merchant-statement/v1/sales"

    start_date_str = "2024-06-14"
    end_date_str = "2024-06-14"
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d")

    current_date = start_date

    headers_terminal = ["companyNumber", "documentNumber", "nsu", "saleDate", "deviceType", "device"]
    print("\t".join(headers_terminal))

    while current_date <= end_date:
        date_str = current_date.strftime("%Y-%m-%d")
        for companyNumber in companyNumbers:
            params = {
                "startDate": date_str,
                "endDate": date_str,
                "parentCompanyNumber": companyNumber,
                "subsidiaries": companyNumber,
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
                    row = [row_data.get("companyNumber"),
                           row_data.get("documentNumber"),
                           row_data.get("nsu"),
                           row_data.get("saleDate"),
                           row_data.get("deviceType"),
                           row_data.get("device")]
                    print("\t".join(map(str, row)))
            else:
                print(f"Falha ao obter dados para o número da empresa {companyNumber} no dia {date_str} após múltiplas tentativas.")

        current_date += timedelta(days=1)


if __name__ == "__main__":
    main()
