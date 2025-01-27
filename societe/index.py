import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import sys  # Importer le module pour la sortie du script
import urllib3, socket
from urllib3.connection import HTTPConnection
import time
import os


# Fonction pour charger les éléments déjà traités depuis un fichier spécifique
def load_processed_elements(filename):
    if os.path.exists(filename):
        with open(filename, 'r', encoding="utf-8") as file:
            return set(file.read().splitlines())
    return set()

# Fonction pour enregistrer les éléments traités dans un fichier spécifique
def save_processed_element(element, filename):
    with open(filename, 'a', encoding="utf-8") as file:
        file.write(f"{element}\n")

def societe():
    HTTPConnection.default_socket_options = ( 
            HTTPConnection.default_socket_options + [
            (socket.SOL_SOCKET, socket.SO_SNDBUF, 1000000), #1MB in byte
            (socket.SOL_SOCKET, socket.SO_RCVBUF, 1000000)
        ])

    chrome_driver_path = r"C:\Users\Administrator\Desktop\scrapping_aiscore\chromedriver\chromedriver.exe"
    chrome_options = Options()
    chrome_options.add_argument("windows-size=800*600")  # Lance le navigateur en mode maximisé
    chrome_options.add_argument("--headless")  # Mode sans interface graphique
    chrome_options.add_argument("--disable-infobars")  # Désactive les barres d'information
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Empêche la détection d'automatisation
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    chrome_options.add_argument("--no-sandbox")  # Pour résoudre certains problèmes de sécurité
    chrome_options.add_argument("--disable-renderer-backgrounding")  # Évite la mise en arrière-plan des processus de rendu
    chrome_options.add_argument("--disable-background-timer-throttling")  # Empêche le ralentissement des minuteries en arrière-plan
    chrome_options.add_argument("--disable-backgrounding-occluded-windows")  # Évite la mise en arrière-plan des fenêtres occultées
    chrome_options.add_argument("--disable-client-side-phishing-detection")  # Désactive la détection de phishing côté client
    chrome_options.add_argument("--disable-crash-reporter")  # Désactive le rapporteur de crash
    chrome_options.add_argument("--disable-gpu")  # Désactive l'utilisation du GPU pour la compatibilité
    chrome_options.add_argument("--silent")  # Réduit les logs inutiles
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_extension("Block-image.crx")

    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)            
    processed_filename = f"C:\\Users\\Administrator\\Desktop\\scrapping_aiscore\\societe\\societe_dep_6_part_C.txt"
    processed_elements = load_processed_elements(processed_filename)

    #Array excel
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\Administrator\Desktop\scrapping_aiscore\credentials.json', scope)
    client = gspread.authorize(creds)
    sheet_id = "1krqcAXwqjFYF9F3ICVvB8hlSquKvL_iQKSliNMpa-Ss" 
    
    sheet = client.open_by_key(sheet_id)
    # wb = sheet.sheet1
    wb = sheet.worksheet("Worksheet")
    ws = wb.get_all_values()

    try:
        url = "https://www.societe.com/cgi-bin/recherche"
        driver.get(url)

        # Accepter les cookies
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '#didomi-notice-agree-button'))
            ).click()
        except Exception:
            print("Le bouton de cookies n'est pas apparu, poursuite de l'exécution...")

        total_elements = len(ws)  # Total des éléments dans la source de données
        processed_count = 0       # Compteur des éléments déjà traités

        #ne pas ignire ligne 1 si igner le 
        # si ignoer alors code est for i, row in enumerate(ws[1:], start=2):
        for i, row in enumerate(ws, start=1):  

            name_company=row[1] #Nom entreprise

            sirene_number = str(row[0])  # Convertit en chaîne de caractères

            last_four_digits_sirene = sirene_number[-4:]  # Prend les 4 derniers chiffres

            if name_company in processed_elements:
                print('element deja traiter')
                processed_count += 1
                continue

            cta_url = f'https://www.societe.com/cgi-bin/liste?ori=avance&nom={name_company}&exa=on&dirig=&pre=&ape=&dep=06'

            driver.execute_script("window.open(arguments[0]);", cta_url)

            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

            driver.switch_to.window(driver.window_handles[-1])

            try:                
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#search_details")))
                elements = driver.find_elements(By.CSS_SELECTOR, 'a.ResultBloc__link__content')  

                if not elements:
                    print("pas elements")
            
                for item in elements:
                    try: 
                        sirene = item.find_element(By.CSS_SELECTOR, 'p:nth-child(3)').text.strip()

                        if not sirene:
                            print("pas des sirene")

                        sirene_result=sirene.split(' ')
                        sirene_number = int(sirene_result[-1])
                        last_four_digits = str(sirene_number)[-4:] 

                        # Prend les 4 derniers chiffres
                        href = item.get_attribute("href")
                    
                        if last_four_digits == last_four_digits_sirene:

                            driver.get(href)
                            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#identite")))
                            
                            try:
                                salarier = driver.find_element(By.CSS_SELECTOR, "#trancheeff-histo-description").text.strip()

                                if salarier:
                                    # Extraction du dernier élément après découpage par espaces
                                    salarie_parts = salarier.split()
                                    
                                    # Parcourt à l'envers pour trouver le premier chiffre
                                    for part in reversed(salarie_parts):
                                        if part.isdigit():
                                            salarier_text = part
                                            break
                                    else:
                                        salarier_text = ''
                                else:
                                    salarier_text = ''

                            except Exception as e:
                                salarier_text = ''

                            li=driver.find_elements(By.CSS_SELECTOR,'.co-resume > ul > li')

                            for item in li:
                                try:
                                    span_text=item.find_element(By.CSS_SELECTOR,'span.ui-label').text.strip()
                                    if span_text =='ADRESSE':
                                        span_adresse = item.find_element(By.CSS_SELECTOR, 'span:nth-child(2) > a').text.strip()
                                        span_adresse_parts = span_adresse.split(',')
                                        # Vérifie si des parties existent avant d'accéder à l'index
                                        if len(span_adresse_parts) > 0:
                                            span_adresse_str = str(span_adresse_parts[0])
                                            wb.update_cell(i, 3, span_adresse_str)
                                        else:
                                            span_adresse_str = ''
                                except Exception as e: 
                                    span_adresse_str = ''
                            
                                        
                            print(f"Sirène trouvé : noms {name_company} numero {sirene} addresse {span_adresse_str} salarié {salarier_text} ") 
                            # Mise à jour de la colonne B avec le nouveau sirene
                            wb.update_cell(i, 1, sirene_number)#update de la numero sirene
                            wb.update_cell(i, 7, salarier_text)#update de la nombre salarier  
                            processed_elements.add(name_company)
                            save_processed_element(name_company, processed_filename)
                            processed_count += 1          

                    except Exception as e:
                        print(f"Erreur lors du traitement de l'élément : {e}")


                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                                                      
            except Exception as e:
                print("pas de data")
                processed_elements.add(name_company)
                save_processed_element(name_company, processed_filename)
                processed_count += 1
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            
            # Vérifiez si tous les éléments ont été traités
        if processed_count >= total_elements:
            print("Tous les éléments ont été traités.")
            driver.close()  # Fermer l'onglet actif
            driver.quit()   # Fermer complètement le navigateur
            sys.exit("Script arrêté car aucune correspondance n'a été trouvée.")  # Sortie immédiate
           
    except Exception as e:
        print(f"Erreur lors de l'exécution : {e}")    
        driver.quit()  # Nettoyer correctement le driver
        
# # Répéter la fonction tout en nettoyant les ressources
while True:
    try:
        societe()
    except Exception as e:
        print(f"Relance de la fonction societe après une erreur globale : {e}")
        # Si une erreur globale survient, fermez toutes les instances de WebDriver avant de relancer