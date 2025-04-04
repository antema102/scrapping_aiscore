import requests

# Remplacez ces variables par votre clé API et ID CX
api_key = 'AIzaSyD10QRiDcYLg6tMwnTzWPjOCcLc02_Lf-s'  # Remplacez par votre clé API
cx = 'f65d1bd411f3c41ef'         # Remplacez par votre ID CX (Custom Search Engine)

# La requête que vous souhaitez envoyer à l'API
query = 'site:www.societe.com Meilleurtaux'

# URL de la requête API
url = f'https://www.googleapis.com/customsearch/v1?q={query}&cx={cx}&key={api_key}'

# Faire la requête GET
response = requests.get(url)

# Vérifier si la requête a réussi
if response.status_code == 200:
    results = response.json()
    if 'items' in results:
        for item in results['items']:
            print(f"Title: {item['title']}")
            print(f"Link: {item['link']}")
            print(f"Snippet: {item['snippet']}")
            print("-" * 50)
    else:
        print("Aucun résultat trouvé.")
else:
    print(f"Erreur: {response.status_code}")
