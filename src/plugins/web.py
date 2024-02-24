import requests


def wwcc_check(WWCC_Number, WWCC_Name):
    if WWCC_Number == "":
        return "Number Empty"
    if WWCC_Name == "":
        return "Name Empty"

    api_url = f"https://api-status-check.wwcc.service.vic.gov.au/v1/wwcc/status/check/cardnumber/{WWCC_Number}/surname/{WWCC_Name}?generateSvReferenceNo=true"
    headers = {"X-Api-Key": "iQlg2FyQc56ULtqNWkZc1abioqHMbDRY9Wy9OcZA"}

    response = requests.get(api_url, headers=headers)
    jsonResponse = response.json()

    result = jsonResponse["statusCheckMessage"]
    date = jsonResponse["datetimeChecked"]

    print(result)

    if "current" in result:
        # print("WWCC is good!")
        return "OK"
    else:
        # print("Oh No! Get them out of here!")
        return "BAD"
