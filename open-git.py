import requests
from openpyxl import Workbook

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

    companyNumbers2 = [
        84233028, 84014059, 82370605, 89322908, 84232897, 88645940, 1471481, 88646351, 82370583, 32144261, 195618,
        91917697, 75539268, 75527162, 75539098, 82370850,
        83928812, 75539527, 84037857, 82370672, 90092147, 92175473, 94212732, 9476418, 91808316, 88244296, 3111628,
        95207872, 93530218, 84233079, 90246225, 32144890,
        84233117, 89339622, 32145268, 94959331, 73836281, 82370788, 75539489, 75527065, 82370621, 82370745, 82370648,
        83830030, 90086155, 92124410, 85544426, 32145454,
        84232749, 82370680, 92969992, 94959102, 14740664, 94959579, 33810311, 94959455, 93419244, 14998173, 95207864,
        75557525, 75557487, 84686774, 73938831, 75527073,
        73938807, 84013842, 88450295, 82445249, 92223389, 85544469, 95102450, 94959056, 94959072, 87815397, 89912764,
        84232668, 89910354, 32143460, 89162420, 32145403,
        84233184, 92953298, 75539217, 73953903, 73850845, 75677350, 75527189, 75539535, 83885625, 88437795, 86891740,
        94090483, 85544680, 86933302, 85788600, 88646203,
        88244580, 32144970, 93569580, 32145586, 82370915, 84232935, 95006346, 95008292, 75527138, 73887145, 87805405,
        73887218, 83928731, 82370753, 92320384, 88729222,
        72982217, 94003114, 85544388, 82370630, 95207880, 86982109, 82370559, 84232633, 32145730, 88244440, 32144512,
        88244202, 73986313, 3008932, 82370664, 73887170, 91590663,
        83830200, 90104951, 82445222, 94234906, 86006118, 82370702, 92069207, 13714473, 95012095, 95008500, 95207899,
        75539330, 83928790, 82370737, 82370885, 84417072,
        82445265, 94572593, 95008845, 91489520, 32145330, 91289548, 92761860, 32145667, 32145128, 82370800, 3067793,
        82370907, 82370923, 91581869, 94434107, 84233052,
        91494117, 83830456, 94959013, 93717393, 84232706, 32144024, 82370710, 75557584, 83928774, 82370940, 94765774,
        85485993, 87814773, 84232978, 32145497, 93053827,
        84232803, 32144709, 32145772, 75539411, 75557614, 33900973, 84382104, 91861675, 94852804, 21220247, 86011553,
        82370591, 3111636, 86930168, 84012374, 89351746,
        36083658, 75539462, 73986461, 82370699, 92414443, 94543828, 3016412, 95400206, 23816090, 86819747, 75538695,
        86820052, 82370494, 95367217, 83967460, 95370234,
        57040117, 58460560, 66359732, 95367225, 73392642, 94764182, 95400729, 91860849, 92054668, 92117791, 94338442,
        93889100, 94378398, 95400460, 95071512, 95071431,
        95367160, 8891141
    ]


    companyNumbers = [
        82370494, 82370648
    ]

    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + access_token
    }

    startdate = "2024-02-29"
    enddate = "2024-02-29"

    data_to_write = []  # Lista para armazenar os dados a serem escritos no Excel

    for companyNumber in companyNumbers:
        params = {
            "startDate": startdate,
            "endDate": enddate,
            "parentCompanyNumber": companyNumber
        }
        response = requests.get(url, params=params, headers=headers)
        if response.status_code == 200:
            content = response.json()['content']
            amount = content[0]['amount']
            total = content[0]['total']
            data_to_write.append((startdate, enddate, companyNumber, amount, total))
            print("Valores inseridos para a empresa", companyNumber)
        elif response.status_code == 204:
            data_to_write.append((startdate, enddate, companyNumber, 0, 0))
            print("Inserido vazio para a empresa", companyNumber)
        else:
            print(f"Erro para a empresa {companyNumber}: {response.status_code}")

    # Agora que todos os dados foram coletados, vamos escrevê-los no Excel
    excel_file_path = r"C:\Users\BG-PROVISORIO\Desktop\Teste-Recebiveis.xlsx"
    workbook = Workbook()
    sheet = workbook.active

    for row_data in data_to_write:
        sheet.append(row_data)

    workbook.save(excel_file_path)
    print("Dados escritos no arquivo Excel com sucesso!")
else:
    print("Erro na solicitação:", response.text)
