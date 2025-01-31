from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
import sys  # Importer le module pour la sortie du script
import socket
from urllib3.connection import HTTPConnection
import os
from multiprocessing import Process,Lock
import requests
import time
files_and_sheets=[]
# Verrou global
lock = Lock()

# Configuration du proxy avec authentification via URL
seleniumwire_options = {
    'proxy': {
        'http': 'http://antema103.gmail.com:9yucvu@gate2.proxyfuel.com:2000',
        'https': 'http://antema103.gmail.com:9yucvu@gate2.proxyfuel.com:2000',
    }
}


for i in range(18,19):  # Départements de 8 à 12
    dep_formatted = str(i).zfill(2)
    parts = [f"part_{j}" for j in range(1, 2)]  # Générer part_1 à part_6
    files_and_sheets.append(
        (f"C:\\Users\\Administrator\\Desktop\\scrapping_aiscore\\societe\\Multi\\DEPT_{dep_formatted}.xlsx", parts)
    )


def check_internet(url="https://www.google.com", timeout=5):
    """Teste la connexion Internet en envoyant une requête à Google."""
    try:
        response = requests.get(url, timeout=timeout)
        return response.status_code == 200
    except requests.ConnectionError:
        return False
    

def h1(url="https://www.google.com", timeout=5):
    """Teste la connexion Internet en envoyant une requête à Google."""
    try:
        response = requests.get(url, timeout=timeout)
        return response.status_code == 200
    except requests.ConnectionError:
        return False

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

def societe(file_path,sheets):
    HTTPConnection.default_socket_options = ( 
            HTTPConnection.default_socket_options + [
            (socket.SOL_SOCKET, socket.SO_SNDBUF, 1000000), #1MB in byte
            (socket.SOL_SOCKET, socket.SO_RCVBUF, 1000000)
        ])

    chrome_driver_path = r"C:\Users\Administrator\Desktop\scrapping_aiscore\chromedriver\chromedriver.exe"
    chrome_options = Options()
    chrome_options.add_argument("--window-size=800,600")  # Dimensions de la fenêtre
    # chrome_options.add_argument("--headless")  # Mode sans interface graphique
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

    # Désactiver JavaScript via les préférences
    prefs = {"profile.managed_default_content_settings.javascript": 2}
    chrome_options.add_experimental_option("prefs", prefs)

    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options, seleniumwire_options=seleniumwire_options)
    processed_text= os.path.splitext(os.path.basename(file_path))[0]
    number = processed_text.split("_")[-1]  # Sépare à "_" et prend la 2e partie
    processed_filename = f"C:\\Users\\Administrator\\Desktop\\scrapping_aiscore\\societe\\Multi\\societe_{processed_text}_{sheets}.txt"
    processed_elements = load_processed_elements(processed_filename)
    new_file_path= f"C:\\Users\\Administrator\\Desktop\\scrapping_aiscore\\societe\\Multi\\{processed_text}_{sheets}.xlsx"

    workbook = load_workbook(file_path)

    # Parcourir toutes les feuilles et supprimer celles qui ne sont pas 'sheets'
    for feuille in workbook.sheetnames:
        if feuille != sheets:  # Si ce n'est pas l'onglet à garder
            ws = workbook[feuille]
            workbook.remove(ws)  # Supprimer l'onglet
            print(f"Onglet {feuille} supprimé.")  # Afficher le nom de l'onglet supprimé

    # Sélectionner la feuille active
    worksheet_name = sheets  # Nom de la feuille à garder dans le fichier Excel
    worksheet = workbook[worksheet_name]

    try:
        url = "https://www.societe.com/cgi-bin/recherche"
        driver.get(url)

        total_elements = worksheet.max_row  # Total des éléments dans la source de données
        processed_count = 0       # Compteur des éléments déjà traités

        #ne pas ignire ligne 1 si igner le 
        # si ignoer alors code est for i, row in enumerate(ws[1:], start=2):
        for i, row in enumerate(worksheet.iter_rows(min_row=1, values_only=True), start=1):  # Ignore la première ligne si c'est un en-tête

            name_company=row[1] #Nom entreprise

            sirene_number = str(row[0])  # Convertit en chaîne de caractères

            last_four_digits_sirene = sirene_number[-4:]  # Prend les 4 derniers chiffres

            if name_company in processed_elements:
                print('element deja traiter')
                processed_count += 1
                continue

            cta_url = f'https://www.societe.com/cgi-bin/liste?ori=avance&nom={name_company}&exa=on&dirig=&pre=&ape=&dep={number}'

            driver.execute_script("window.open(arguments[0]);", cta_url)

            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

            driver.switch_to.window(driver.window_handles[-1])


            if not check_internet():
                print("❌ Pas de connexion Internet. Fermeture du script.")
                driver.close()
                driver.quit()
                return  # Quitte immédiatement

            try:                
                elements = driver.find_elements(By.CSS_SELECTOR, 'a.ResultBloc__link__content')

                if not elements:
                    
                    h1 = driver.find_element(By.CSS_SELECTOR, '#appMain > div > section > div > h1')

                    if h1:
                        print("ip bloquer")
                        driver.close()
                        driver.quit()
                        return False

                    processed_elements.add(name_company)
                    save_processed_element(name_company, processed_filename)
                    processed_count += 1    
                else:
                    for item in elements:
                        try: 
                            sirene = item.find_element(By.CSS_SELECTOR, 'p:nth-child(3)').text.strip()
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
                                                worksheet.cell(row=i, column=3, value=span_adresse_str)
                                            else:
                                                span_adresse_str = ''
                                    except Exception as e: 
                                        span_adresse_str = ''
                                
                                            
                                print(f"Sirène trouvé : noms {name_company} numero {sirene} addresse {span_adresse_str} salarié {salarier_text} ") 
                                # Mise à jour de la colonne B avec le nouveau sirene
                                worksheet.cell(row=i, column=1, value=sirene_number)
                                worksheet.cell(row=i, column=7, value=salarier_text) 

                                # Sauvegarder les modifications dans le fichier Excel
                                workbook.save(new_file_path) # Mise à jour du texte des salariés

                        except Exception as e:
                            print(f"Erreur lors du traitement de l'élément : {e}")

                    processed_elements.add(name_company)
                    save_processed_element(name_company, processed_filename)
                    processed_count += 1    

                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                                                      
            except Exception as e:
                print("Error",e)
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

        # Vérifiez si tous les éléments ont été traités
        if processed_count >= total_elements:
            print("Tous les éléments ont été traités.")
            driver.close()  # Fermer l'onglet actif
            driver.quit()   # Fermer complètement le navigateur
            print("Script arrêté car aucune correspondance n'a été trouvée.")  # Sortie immédiate
            return True 
           
    except Exception as e:
        print(f"Erreur lors de l'exécution : {e}")    
        driver.quit()  # Nettoyer correctement le driver
        return False  # Retourne False pour signaler une erreur

def retry_societe(file_path, sheet_name):
    """
    Fonction pour exécuter et relancer le traitement si une erreur se produit.
    """
    while True:  # Boucle infinie jusqu'à ce que le traitement soit terminé avec succès
        success = societe(file_path, sheet_name)
        time.sleep(10)
        if success:
            break  # Sort de la boucle si le traitement est terminé
        else:
            print(f"Relance du traitement pour {file_path} - {sheet_name}")

        
def launch_processes():
    """
    Fonction pour lancer les traitements en simultané.
    """
    processes = []  # Liste pour stocker les processus

    for file_path, sheets in files_and_sheets:
        for sheet_name in sheets:
            print(f"Création d'un processus pour {file_path} - {sheet_name}")
            # Créer un processus pour chaque combinaison fichier/feuille
            process = Process(target=retry_societe, args=(file_path, sheet_name))
            processes.append(process)
            process.start()  # Lancer le processus

    # Attendre que tous les processus soient terminés
    for process in processes:
        process.join()

if __name__ == "__main__":
    print("Lancement des traitements en simultané...")
    launch_processes()
    print("Tous les traitements sont terminés.")