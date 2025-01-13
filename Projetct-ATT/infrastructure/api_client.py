import requests

def fetch_api_data(base_url, designator):
    url = f"{base_url}/dados_enlace/{designator}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Erro ao consultar {url}: {response.status_code}")
        return None
