import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time
import requests
import os
from datetime import datetime,timedelta


# Fonction pour charger les éléments déjà traités depuis un fichier spécifique
def load_processed_elements(filename):
    if os.path.exists(filename):
        with open(filename, 'r',encoding="utf-8") as file:
            return set(file.read().splitlines())
    return set()

# Fonction pour enregistrer les éléments traités dans un fichier spécifique
def save_processed_element(element, filename):
    with open(filename, 'a',encoding="utf-8") as file:
        file.write(f"{element}\n")

# Fonction pour extraire les données et les enregistrer dans un fichier Excel
def process_url(url):
    try:
        #Intialisations 
        chrome_driver_path = r"C:\Users\Administrator\Desktop\scrapping_aiscore\chromedriver\chromedriver.exe" 
        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service)
        
        #Google sheets
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\Administrator\Desktop\scrapping_aiscore\credentials.json', scope)
        client = gspread.authorize(creds)
        sheet_id = "16c6lONHjvr5--8A6Ed729mNY1lCorTShVYZUo-fwPsE"
        sheet = client.open_by_key(sheet_id)
        wb =sheet.sheet1
        ws = wb.get_all_values()

        #Date
        data_str = url.split('/')[-1]
        processed_filename = f"C:\\Users\\Administrator\\Desktop\\scrapping_aiscore\\scrapping_cote_initiales\\processed_elements_{data_str}.txt"
        date_object = datetime.strptime(str(data_str),"%Y%m%d")

        # Formater la date
        formatted_date = date_object.strftime("%d/%m/%Y")
        date_present=False
        
        for row in ws[1:]:
            cell_score=row[0]
            if cell_score == formatted_date:
                date_present=True
                break

        if not date_present:
            wb.append_row([formatted_date,"", "",""])
        else:
            print("date deja présent")

        #Charger Elements
        processed_elements = load_processed_elements(processed_filename)

        try:
            driver.get(url)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#app")))

            tous_selector = "#app > DIV:nth-of-type(3) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(1) > SPAN:nth-of-type(1)"
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, tous_selector))).click()

            time.sleep(4)

            second_selector = "#app > DIV:nth-of-type(3) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > DIV:nth-of-type(2) > LABEL > SPAN > SPAN"
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, second_selector))).click()

            time.sleep(4)

            while True:
                match_elements = driver.find_elements(By.CSS_SELECTOR, 'a.match-container')
                new_data_found = False
                # Traiter les éléments de match
                for element in match_elements:
                    nomEquipe = element.find_element(By.CSS_SELECTOR, 'span.name.minitext.maxWidth160').text.strip()

                    if nomEquipe in processed_elements:
                        print(f"L'équipe {nomEquipe} a déjà été traitée, passage au suivant.")
                        continue  # Passe au prochain élément sans traiter

                    href = element.get_attribute("href")
                    processed_elements.add(nomEquipe)
                    save_processed_element(nomEquipe, processed_filename)
                    new_data_found = True

                    # Ouvre le lien et récupère les données
                    driver.execute_script("window.open(arguments[0]);", href)
                    driver.switch_to.window(driver.window_handles[-1])

                    # Cliquer sur la troisieme bouton
                    thirst_selector = "#app > div.detail.view.border-box.back > div.tab-bar > div > div > a:nth-child(2)"
                    thirst_selectorText=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, thirst_selector))).text.strip()

                    if thirst_selectorText =="Cotes":
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, thirst_selector))).click()
                        
                        #affiche les cotes lorsque je clique sur le bouton cote
                        score1 = '#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.flex.w-bar-100.homeBox > div.h-top-center.matchStatus3 > div.font-bold.home-score > span:last-child'
                        score2 = '#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.flex.w-bar-100.homeBox > div.h-top-center.matchStatus3 > div.font-bold.away-score > span'

                        cote_selectorBet365_1 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(1) > span > span'
                        cote_selectorBet365_2 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(2) > span'
                        cote_selectorBet265_3 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(3) > span > span'

                        cote_nombreButBet365  ="#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(1) > span"
                        cote_plusBut365=       "#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(2) > span"
                        cote_moinsBut365=      "#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(3) > span"

 
                        try:
                            score_text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, score1))).text.strip()
                            score_text1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, score2))).text.strip()

                            try:
                                cote_Bet365 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_selectorBet365_1))).text.strip().replace('.', '')
                                cote_Bet365_2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_selectorBet365_2))).text.strip().replace('.', '')
                                cote_Bet365_3 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_selectorBet265_3))).text.strip().replace('.', '')
                                cotesBet365 = f"{cote_Bet365}/{cote_Bet365_2}/{cote_Bet365_3}"
                            except Exception:
                                cotesBet365 = ''

                            try:
                                # Récupère la donnée brute de nombre de buts
                                nombreButBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_nombreButBet365))).text.strip()

                                # Initialisation des valeurs
                                TotalButBet365 = ''

                                # Vérifie s'il y a un '/' pour extraire les deux nombres
                                if '/' in nombreButBet:
                                    # Sépare les deux valeurs
                                    valeurs = nombreButBet.split('/')
                                    premier_nombre = float(valeurs[0])
                                    second_nombre = float(valeurs[1])

                                    # Vérifie si les deux valeurs sont entre 2 et 3
                                    if 2 <= premier_nombre <= 3 and 2 <= second_nombre <= 3:
                                        # Récupère les valeurs pour plus et moins si les conditions sont respectées
                                        plusButBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_plusBut365))).text.strip().replace('.', '')
                                        moinsButBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_moinsBut365))).text.strip().replace('.', '')
                                        TotalButBet365 = f"{plusButBet}/{moinsButBet}"
                                else:
                                    # Si aucun '/' alors on vérifie si le nombre unique est entre 2 et 3
                                    unique_nombre = float(nombreButBet)
                                    if 2 <= unique_nombre <= 3:
                                        plusButBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_plusBut365))).text.strip().replace('.', '')
                                        moinsButBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_moinsBut365))).text.strip().replace('.', '')
                                        TotalButBet365 = f"{plusButBet}/{moinsButBet}"

                            except Exception:
                                TotalButBet365 = ''  # En cas d'erreur, TotalButBet365 est vide
                            
                            wb.append_row([f"{score_text}-{score_text1}",cotesBet365,TotalButBet365,formatted_date])
                            print(f"{score_text}-{score_text1}",cotesBet365,TotalButBet365,formatted_date,nomEquipe)
                        
                        except Exception as e:
                            print("match reporté")

                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                    else:
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                    # Faire défiler vers le bas pour charger plus de tâches

                els = driver.execute_script("return document.getElementsByClassName('match-container');")

                if els:
                    driver.execute_script("arguments[0].scrollIntoView();", els[-1])  # Défile vers le dernier élément
                    time.sleep(2)  # Attendre un peu pour permettre le chargement

                    # Si aucun nouvel élément n'a été trouvé, sortir de la boucle
                if not new_data_found:
                    print("Aucun nouvel élément trouvé, sortie de la boucle.")
                    break
                
                print(f"Tous les éléments traités pour {url}.")  # Afficher que tous les éléments ont été traités pour cette URL.

        except Exception as e:
                print("Erreur dans le selenium",e) 
                
    except Exception as e:
        print("Erreur dans le debuts",e)  

    finally:
        driver.quit()



# Configuration
base_url = "https://www.aiscore.com/fr"
# Date de début
start_date = datetime.strptime("20241214", "%Y%m%d")  

# Fonction pour générer des URLs avec des dates
def generate_urls_until_yesterday(base_url, start_date):
    try:
        urls = []
        today = datetime.now()
        yesterday = today - timedelta(days=1)  # Calculer hier
        current_date = start_date

        # Boucle pour générer des URLs jusqu'à hier
        while current_date <= yesterday:
            formatted_date = current_date.strftime("%Y%m%d")
            urls.append(f"{base_url}/{formatted_date}")
            current_date += timedelta(days=1)

        return urls

    except Exception as e:
        print("Erreur date",e)

# Générer les URLs
urls = generate_urls_until_yesterday(base_url, start_date)

#Traiter chaque URL et enregistre les donnés
for url in urls:
    process_url(url)