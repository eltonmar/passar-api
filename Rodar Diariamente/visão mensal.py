import requests
from openpyxl import Workbook
import datetime

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


def main():
    access_token, refresh_token, _, _, _ = get_tokens()

    if not access_token or not refresh_token:
        print("Falha na obtenção do token.")
        return

    url = "https://api.userede.com.br/redelabs/merchant-statement/v2/receivables/summary"

    companyNumbers = [
        92320384, 32145403, 32145454, 32145497, 32145586, 32145667, 23816090, 84417072,
        84232706, 33810311, 32145772, 84232749, 86819747, 84233052, 84233079, 91494117,
        91289548, 84233117, 86006118, 86011553, 84233184, 75538695, 73887145, 86930168,
        75527073, 73986461, 88244296, 75557487, 82370591, 87814773, 73836281, 87815397,
        84232897, 86982109, 84232935, 195618, 86820052, 95102450, 75539330, 91590663,
        75539268, 89351746, 75557584, 84233028, 82370680, 82370702, 88244202, 88244440,
        32145330, 82370583, 82370605, 75539527, 75557614, 82370664, 82370494, 75557525,
        75539535, 83830200, 94959579, 83830030, 95207864, 82370648, 82370672, 94959455,
        82370699, 95006346, 82370737, 95367217, 82370753, 82370788, 82370621, 82370710,
        82370745, 82370923, 95207872, 82370800, 82370850, 82370940, 82370885, 82370907,
        83928731, 83928774, 83928790, 83928812, 83967460, 83885625, 84037857, 84013842,
        88244580, 88645940, 88646203, 88646351, 89162420, 89339622, 89322908, 89912764,
        89910354, 90246225, 91489520, 91808316, 91917697, 92069207, 92761860, 92969992,
        92953298, 93053827, 93530218, 93569580, 93717393, 95008292, 95207880, 94959331,
        95008500, 94959102, 95008845, 94959072, 94959056, 95012095, 94959013, 95207899,
        95370234, 57040117, 85544426, 58460560, 66359732, 85544388, 82445249, 95367225,
        86891740, 73392642, 72982217, 85485993, 82445222, 94764182, 82445265, 95400729,
        85544680, 88450295, 88437795, 88729222, 90104951, 90092147, 90086155, 91581869,
        91861675, 91860849, 92054668, 92175473, 92124410, 92117791, 92223389, 92414443,
        94338442, 94003114, 93889100, 94234906, 94090483, 94212732, 94572593, 94434107,
        94378398, 94852804, 94543828, 95400460, 94765774, 95071512, 95071431, 95367160,
        95425764, 95477756, 95515208, 95515216, 95515224, 95515232, 95515240, 95515267, 95515275, 95515283, 93419244,
        95515291, 95515313, 95515321, 95515330, 95515356, 95515372, 95515402, 95515410, 95515445, 95515496, 95515500,
        95515518, 95515526, 95515534, 95515542, 95515550, 95515569, 95515577, 95587632, 95593446, 95587543
    ]


    data = datetime.datetime.now()

    data_to_write = []

    lista_meses = [(data.month + i) % 12 or 12 for i in range(13)]


    for companyNumber in companyNumbers:
        current_year = data.year
        for month in lista_meses:
            if month == 1:
                current_year += 1
            if month in [1, 3, 5, 7, 8, 10, 12]:
                dayfin = 31
            elif month in [4, 6, 9, 11]:
                dayfin = 30
            else:
                dayfin = 28

            if data.month == month:
                dayinit = 1
            else:
                dayinit = 1

            startdate = f"{current_year:02d}-{month:02d}-{dayinit:02d}"
            enddate = f"{current_year:02d}-{month:02d}-{dayfin:02d}"

            params = {
                "startDate": startdate,
                "endDate": enddate,
                "parentCompanyNumber": companyNumber
            }

            headers = {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + access_token
            }

            response = requests.get(url, params=params, headers=headers)

            if response.status_code == 200:
                content = response.json().get('content')
                if content:
                    amount = content[0]['amount']
                    total = content[0]['total']
                    data_to_write.append((startdate, enddate, companyNumber, amount, total))
                    print("Valores inseridos para a empresa", companyNumber)
                else:
                    data_to_write.append((startdate, enddate, companyNumber, 0, 0))
                    print("Inserido vazio para a empresa", companyNumber)
            elif response.status_code == 401:
                new_access_token, new_refresh_token = refresh_access_token(refresh_token)
                if new_access_token:
                    access_token = new_access_token
                    refresh_token = new_refresh_token
                    headers["Authorization"] = "Bearer " + access_token
                    response = requests.get(url, params=params, headers=headers)
                    if response.status_code == 200:
                        content = response.json().get('content')
                        if content:
                            amount = content[0]['amount']
                            total = content[0]['total']
                            data_to_write.append((startdate, enddate, companyNumber, amount, total))
                            print("Valores inseridos para a empresa", companyNumber)
                        else:
                            data_to_write.append((startdate, enddate, companyNumber, 0, 0))
                            print("Inserido vazio para a empresa", companyNumber)
                    else:
                        print(f"Erro para a empresa {companyNumber}: {response.status_code}")
                else:
                    print("Falha na atualização do token.")
            else:
                print(f"Erro para a empresa {companyNumber}: {response.status_code}")

    # Salvamento do arquivo Excel após o término do loop
    excel_file_path = r"C:\Users\BG-PROVISORIO\Desktop\Teste-Recebiveis.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    for row_data in data_to_write:
        sheet.append(row_data)
    workbook.save(excel_file_path)
    print("Dados escritos no arquivo Excel com sucesso!")

if __name__ == "__main__":
    main()
