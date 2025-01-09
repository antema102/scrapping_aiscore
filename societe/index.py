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
    chrome_driver_path = r"C:\Users\Administrator\Desktop\scrapping_aiscore\chromedriver\chromedriver.exe"
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")  # Démarrer en mode maximisé
    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)            
    processed_filename = f"C:\\Users\\Administrator\\Desktop\\scrapping_aiscore\\societe\\societe_dep_1_part_3.txt"
    processed_elements = load_processed_elements(processed_filename)

    #Array excel
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\Administrator\Desktop\scrapping_aiscore\credentials.json', scope)
    client = gspread.authorize(creds)
    sheet_id = "1QZV0VJrHosEoeuPraz3oK_N06hiVFcVOD2qbgoRIFyk" 
    
    sheet = client.open_by_key(sheet_id)
    # wb = sheet.sheet1
    wb = sheet.worksheet("part_3")
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


        for i, row in enumerate(ws[1:], start=1):  # Commence à la ligne 3 (index Excel)

            name_company=row[1] #Nom entreprise

            sirene_number = str(row[0])  # Convertit en chaîne de caractères

            last_four_digits_sirene = sirene_number[-4:]  # Prend les 4 derniers chiffres

            if name_company in processed_elements:
                print('element deja traiter')
                continue

            cta_url = f'https://www.societe.com/cgi-bin/liste?ori=avance&nom={name_company}&dirig=&pre=&ape=&dep=01'

            driver.execute_script("window.open(arguments[0]);", cta_url)

            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

            driver.switch_to.window(driver.window_handles[-1])

            try:                
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#search_details")))
                elements = driver.find_elements(By.CSS_SELECTOR, 'a.ResultBloc__link__content')    
                new_data_found = False 
            
                for item in elements:
                    try: 
                        sirene = item.find_element(By.CSS_SELECTOR, 'p:nth-child(3)').text.strip()
                        sirene_result=sirene.split(' ')
                        sirene_number = int(sirene_result[-1])
                        last_four_digits = str(sirene_number)[-4:] 
                        
                        # Prend les 4 derniers chiffres
                        href = item.get_attribute("href")

                        if last_four_digits == last_four_digits_sirene:
                            new_data_found=True  
                            driver.get(href)

                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#identite")))
                            
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

                    except Exception as e:
                        print(f"Erreur lors du traitement de l'élément : {e}")

                if not new_data_found:
                    print("Aucune correspondance trouvée, fermeture de l'onglet.")

                processed_elements.add(name_company)
                save_processed_element(name_company, processed_filename)  

                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                                                 
            except Exception as e:
                print("pas de data")
                processed_elements.add(name_company)
                save_processed_element(name_company, processed_filename)
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

    except Exception as e:
        print(f"Erreur lors de l'exécution : {e}")
        

societe()