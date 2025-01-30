from seleniumwire import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time

# Configuration du proxy avec authentification via URL
seleniumwire_options = {
    'proxy': {
        'http': 'http://antema103.gmail.com:9yucvu@gate2.proxyfuel.com:2000',
        'https': 'http://antema103.gmail.com:9yucvu@gate2.proxyfuel.com:2000',
    }
}

# Créez un objet Service avec le chemin du chromedriver
service = Service(r'C:\Users\antem\Desktop\scrapping_aiscore\chromedriver\chromedriver.exe')

# Créez les options pour Chrome
options = Options()

# Lancer le navigateur Chrome avec les options définies et le service
driver = webdriver.Chrome(service=service, options=options, seleniumwire_options=seleniumwire_options)

# Accéder à l'URL pour vérifier l'IP
driver.get("http://checkip.amazonaws.com")

# Afficher le contenu de la page
print(driver.page_source)

# Attendre un peu avant de fermer
time.sleep(100000)

# Fermer le navigateur
driver.quit()
#pip install selenium_wire
#pip install blinker==1.4
#pip install setuptools
#pip install webdriver-manager