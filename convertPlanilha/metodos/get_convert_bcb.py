import requests


def obter_taxas_de_conversao():
    url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.10813/dados?formato=json"
    response = requests.get(url)
    
    if response.status_code == 200:
        try:
            taxas = {item["data"]: float(item["valor"]) for item in response.json()}
            return taxas
        except (IndexError, ValueError):
            return None
    else:
        return None