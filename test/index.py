from seleniumwire import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time

# Définition du proxy
proxy_host = "gate2.proxyfuel.com"
proxy_port = "2000"
proxy_user = "antema103.gmail.com"
proxy_pass = "9yucvu"

seleniumwire_options = {
    'proxy': {
        'http': f'http://{proxy_host}:{proxy_port}',
        'https': f'https://{proxy_host}:{proxy_port}',
        'no_proxy': 'localhost,127.0.0.1'
    }
}

# Service ChromeDriver
service = Service(r'C:\Users\antem\Desktop\scrapping_aiscore\chromedriver\chromedriver.exe')

# Options Chrome
options = Options()

try:
    # Lancer le navigateur
    driver = webdriver.Chrome(service=service, options=options, seleniumwire_options=seleniumwire_options)
    
    # Ajouter authentification manuelle (si nécessaire)
    driver.proxy_auth = (proxy_user, proxy_pass)

    # Tester le proxy
    driver.get("http://checkip.amazonaws.com")
    
    # Afficher l'IP obtenue
    print("✅ Connexion réussie :", driver.page_source)

except Exception as e:
    if "Invalid proxy server credentials supplied" in str(e):
        print("❌ Erreur : Identifiants proxy incorrects. Vérifie ton username/password.")
    else:
        print(f"❌ Autre erreur détectée : {e}")

finally:
    # Fermer le navigateur si ouvert
    try:
        driver.quit()
    except:
        pass
