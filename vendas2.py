import requests
from openpyxl import Workbook

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
        return access_token, refresh_token
    else:
        print("Falha na obtenção do token. Status code:", response.status_code)
        print("Response:", response.text)
        return None, None

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
        print("Response:", response.text)
        return None, None

def fetch_transactions(access_token, companyNumber, start_date, end_date):
    url = "https://api.userede.com.br/redelabs/merchant-statement/v1/sales"
    params = {
        "startDate": start_date,
        "endDate": end_date,
        "parentCompanyNumber": companyNumber,
        "subsidiaries": companyNumber
    }

    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + access_token
    }

    response = requests.get(url, params=params, headers=headers)

    if response.status_code == 200:
        try:
            return response.json().get('content', [])
        except ValueError:
            print(f"Erro ao decodificar JSON para a empresa {companyNumber}")
            return []
    else:
        print(f"Erro para a empresa {companyNumber}: {response.status_code}")
        print("Response:", response.text)
        return []

def main():
    access_token, refresh_token = get_tokens()

    if not access_token or not refresh_token:
        print("Falha na obtenção do token.")
        return

    companyNumbers = [3016412, 84232633, 84232668, 32143460, 3111628]
    data_to_write = []

    manual_startdate = "2024-05-01"
    manual_enddate = "2024-05-15"

    for companyNumber in companyNumbers:
        transactions = fetch_transactions(access_token, companyNumber, manual_startdate, manual_enddate)
        if not transactions:
            access_token, refresh_token = refresh_access_token(refresh_token)
            if access_token:
                transactions = fetch_transactions(access_token, companyNumber, manual_startdate, manual_enddate)

        for transaction in transactions:
            if isinstance(transaction, dict):
                data_to_write.append([
                    transaction.get('deviceType', ''),
                    transaction.get('netAmount', 0),
                    transaction.get('flex', False),
                    transaction.get('cardNumber', ''),
                    transaction.get('captureTypeCode', 0),
                    transaction.get('authorizationCode', 0),
                    transaction.get('amount', 0),
                    transaction.get('movementDate', ''),
                    transaction.get('saleHour', ''),
                    transaction.get('mdrAmount', 0),
                    transaction.get('flexFee', 0),
                    transaction.get('brandCode', 0),
                    transaction.get('discountAmount', 0),
                    transaction.get('boardingFeeAmount', 0),
                    transaction.get('saleDate', ''),
                    transaction.get('tracking', [{}])[0].get('date', ''),
                    transaction.get('tracking', [{}])[0].get('amount', 0),
                    transaction.get('tracking', [{}])[0].get('status', ''),
                    transaction.get('saleSummaryNumber', 0),
                    transaction.get('nsu', 0),
                    transaction.get('flexAmount', 0),
                    transaction.get('device', ''),
                    transaction.get('installmentQuantity', 0),
                    transaction.get('captureType', ''),
                    transaction.get('feeTotal', 0),
                    transaction.get('prePaid', False),
                    transaction.get('tokenized', False),
                    transaction.get('status', ''),
                    transaction.get('mdrFee', 0),
                    transaction.get('merchant', {}).get('companyNumber', ''),
                    transaction.get('merchant', {}).get('companyName', ''),
                    transaction.get('merchant', {}).get('documentNumber', ''),
                    transaction.get('merchant', {}).get('tradeName', ''),
                    transaction.get('modality', {}).get('type', ''),
                    transaction.get('modality', {}).get('code', 0),
                    transaction.get('modality', {}).get('product', ''),
                    transaction.get('modality', {}).get('productCode', 0)
                ])
                print(f"Transação adicionada para a empresa {transaction.get('merchant', {}).get('companyNumber', '')}: {transaction}")

    excel_file_path = r"C:\Users\BG-PROVISORIO\Desktop\Teste-GestaoDeVendas.xlsx"
    workbook = Workbook()
    sheet = workbook.active

    sheet.append([ "DeviceType", "NetAmount", "Flex", "CardNumber", "CaptureTypeCode", "AuthorizationCode",
        "Amount", "MovementDate", "SaleHour", "MdrAmount", "FlexFee", "BrandCode",
        "DiscountAmount", "BoardingFeeAmount", "SaleDate", "TrackingDate", "TrackingAmount",
        "TrackingStatus", "SaleSummaryNumber", "Nsu", "FlexAmount", "Device", "InstallmentQuantity",
        "CaptureType", "FeeTotal", "PrePaid", "Tokenized", "Status", "MdrFee", "MerchantCompanyNumber",
        "MerchantCompanyName", "MerchantDocumentNumber", "MerchantTradeName", "ModalityType", "ModalityCode",
        "ModalityProduct", "ModalityProductCode"
    ])

    for row_data in data_to_write:
        sheet.append(row_data)

    workbook.save(excel_file_path)
    print("Dados escritos no arquivo Excel com sucesso!")

if __name__ == "__main__":
    main()
