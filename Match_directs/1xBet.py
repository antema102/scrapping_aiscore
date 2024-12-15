# import openpyxl
# from openpyxl import Workbook,load_workbook
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
from datetime import datetime
import re 


# Fonction pour charger les éléments déjà traités depuis un fichier spécifique
def load_processed_elements(filename):
    if os.path.exists(filename):
        with open(filename, 'r',encoding='utf-8') as file:
            return set(file.read().splitlines())
    return set()

# Fonction pour enregistrer les éléments traités dans un fichier spécifique
def save_processed_element(element, filename):
    with open(filename, 'a',encoding="utf-8") as file:
        file.write(f"{element}\n")

def remove_element(element,filename):
    elements=load_processed_elements(filename)
    if element in elements:
         elements.remove(element)

    with open(filename,'w',encoding="utf-8") as file:
         for el in elements:
              file.write(f"{el}\n")

def process_url():
    # Définir le chemin vers le ChromeDriver
    try:
        chrome_driver_path = r"C:\Users\Administrator\Desktop\scrapping_aiscore\chromedriver-win64\chromedriver.exe" 
        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service)
        url="https://www.aiscore.com/fr/"
        driver.get(url)
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:/Users/Administrator/Desktop/scrapping_aiscore/credentials.json", scope)
        client = gspread.authorize(creds)
        sheet_id = "1YhuJe7DZ-2IP3l1DoeBdk1xnALlE1OOtXmX2gFw6c4c"
        sheet = client.open_by_key(sheet_id)
        wb_comparison =sheet.sheet1
        ws_comparison = wb_comparison.get_all_values()
        processed_filename = f"processed_elements.txt"
        # Créer ou accéder à des onglets dans Google Sheets
        try:
            ws_main = sheet.worksheet("Match Data")  # Onglet principal
            rows_main = ws_main.get_all_values()
        except gspread.exceptions.WorksheetNotFound:
            ws_main = sheet.add_worksheet(title="Match Data", rows=1000, cols=20)
            ws_main.append_row(["Date","Heure","Ligue","Match", "Score","1XBET ODDS"])
            rows_main = ws_main.get_all_values()

        try:
            ws_buts_total_bet365 = sheet.worksheet("Buts Total 1XBET")  # Onglet Bet365
            rows_buts_total_bet365  = ws_buts_total_bet365.get_all_values()
        except gspread.exceptions.WorksheetNotFound:
            ws_buts_total_bet365 = sheet.add_worksheet(title="Buts Total 1XBET", rows=1000, cols=20)
            ws_buts_total_bet365.append_row(["Date","Heure","Ligue","Match", "Moins de 2 buts", "Plus de 2 buts ou égales", "1XBET O/U 2.5"])
            rows_buts_total_bet365  = ws_buts_total_bet365.get_all_values()


        processed_elements = load_processed_elements(processed_filename)

        try:
            second_selector = "#app > DIV:nth-of-type(3) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > DIV:nth-of-type(2) > LABEL > SPAN > SPAN"
            WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.CSS_SELECTOR, second_selector))).click()
            time.sleep(5) # Attendre un peu pour permettre le chargement
            # Initialiser la variable date_deja_presente
            aujourd_hui = datetime.today().date()
            aujourd_hui_str = aujourd_hui.strftime('%Y-%m-%d')
            date_deja_presenteMain = False
            date_deja_presenteBet365 = False
            date_deja_presente1xBet = False

            for row in rows_main[1:]:  # values_only=True pour obtenir les valeurs sans les objets cellule
                cell_date = row[0]  # Supposons que la date se trouve dans la première colonne
                if cell_date == aujourd_hui_str:
                    date_deja_presenteMain = True
                    break  # On peut sortir de la boucle si on trouve la date
            
            for row in rows_buts_total_bet365[1:]:  # values_only=True pour obtenir les valeurs sans les objets cellule
                cell_date = row[0]  # Supposons que la date se trouve dans la première colonne
                if cell_date == aujourd_hui_str:
                    date_deja_presenteBet365 = True
                    break  # On peut sortir de la boucle si on trouve la date
            

            # Ajouter la date si elle n'est pas déjà présente
            if not date_deja_presenteMain:
                ws_main.append_row([aujourd_hui_str,"","","","","","",""])
            else:
                print("La date est déjà présente Main.")

            # Ajouter la date si elle n'est pas déjà présente
            if not date_deja_presenteBet365:
                ws_buts_total_bet365.append_row([aujourd_hui_str,"","","","","",""])
            else:
                print("La date est déjà présente Bet.")
                

            while True: 
                
                    match_elements = driver.find_elements(By.CSS_SELECTOR, 'a.match-container')
                    # Traiter les éléments de match
                    for element in match_elements:
                        nomEquipe = element.find_element(By.CSS_SELECTOR, 'span.name.minitext.maxWidth160').text.strip()
                        if nomEquipe in processed_elements:
                            print(f"L'équipe {nomEquipe} a déjà été traitée, passage au suivant.")
                            continue  # Passe au prochain élément sans traiter

                        times = element.find_element(By.CSS_SELECTOR,'span.time.minitext').text
                        href = element.get_attribute("href")
                        # new_data_found = True
                        try:
                            # Ouvre le lien et récupère les données
                            driver.execute_script("window.open(arguments[0]);", href)
                            driver.switch_to.window(driver.window_handles[-1])
                            thirst_selector = "#app > div.detail.view.border-box.back > div.tab-bar > div > div > a:nth-child(2)"         
                            thirst_selectorText= WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, thirst_selector))).text.strip()
                            temps = "#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.flex.w-bar-100.homeBox > div.h-top-center.matchStatus2 > div.flex-1.text-center.scoreBox > div.h-16.m-b-4 > span > span:nth-child(1)"
                            ligue = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR,"div.top.color-333.flex-col.flex.align-center > div.comp-name > a"))).text.strip()
                            
                            try: 
                                temps__text = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, temps)))
                                
                                if thirst_selectorText == "Cotes":

                                    #Cliquer sur le Bouton si cotes 
                                    WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, thirst_selector))).click()

                                    deuxiemeEquipe__selector ="#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.flex.w-bar-100.homeBox > div.away-box > div > a"

                                    deuxiementEquipe = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, deuxiemeEquipe__selector))).text.strip()

                                    ws_main.append_row(["",times,ligue,f"{nomEquipe} vs {deuxiementEquipe}","","",""])

                                    cote_selectorBet365_1 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(1) > span > span'
                                    cote_selectorBet365_2 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(2) > span'
                                    cote_selectorBet365_3 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(3) > span > span'
                                    cote_nombreButBet365  = "#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div.flex.w100.borderBottom > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(1) > span"
                                    cote_plusBut365       = "#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(2) > span"
                                    cote_moinsBut365      = "#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(3) > span"

                                    # Ligue ="#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.comp-name > a"

                                    try:
                                            cote_Bet365   = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_selectorBet365_1))).text.strip().replace('.', '')
                                            cote_Bet365_2 = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_selectorBet365_2))).text.strip().replace('.', '')
                                            cote_Bet365_3 = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_selectorBet365_3))).text.strip().replace('.', '')
                                            cotesBet365   = f"{cote_Bet365}/{cote_Bet365_2}/{cote_Bet365_3}"
                                            
                                            for row in ws_comparison[1:]:  # Commence à la ligne 2 pour ignorer l'en-tête
                                                # On récupère la cellule dans la colonne B de chaque ligne
                                                cell_a = row[0]  
                                                cell_b = row[3]
                                                if cell_b == cotesBet365:
                                                    scoreBet365 = cell_a
                                                    ws_main.append_row(["","","","",scoreBet365,cotesBet365])  # Colonne de score à gauche, colonne de cote 1xBet vide pour l'instant                      
                                    except Exception:
                                            cotesBet365 = ''

                                    try:
                                            # Récupère la donnée brute de nombre de buts
                                            nombreButBet = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_nombreButBet365))).text.strip()
                                            # Initialisation des valeurs
                                            TotalButBet365 = ''
                                            ws_buts_total_bet365.append_row(["",times,ligue,f"{nomEquipe} vs {deuxiementEquipe}","","", ""])  # Ajoute une ligne vide pour séparer les matchs
                                            total_moins_de_2_buts = 0
                                            total_2_buts_ou_plus = 0
                                            # Vérifie s'il y a un '/' pour extraire les deux nombres
                                            if '/' in nombreButBet:
                                                # Sépare les deux valeurs
                                                valeurs = nombreButBet.split('/')
                                                premier_nombre = float(valeurs[0])
                                                second_nombre = float(valeurs[1])
                                                # Vérifie si les deux valeurs sont entre 2 et 3
                                                if 2 <= premier_nombre <= 3 and 2 <= second_nombre <= 3.0:
                                                    # Récupère les valeurs pour plus et moins si les conditions sont respectées
                                                    plusButBet = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_plusBut365))).text.strip().replace('.', '')
                                                    moinsButBet = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_moinsBut365))).text.strip().replace('.', '')
                                                    TotalButBet365 = f"{plusButBet}/{moinsButBet}"

                                                    for row in ws_comparison[1:]:
                                                        cell_a = row[0]
                                                        cell_c = row[4]
                                                        if cell_c == TotalButBet365:  
                                                            scoreButBetToTaux = cell_a
                                                            equipe_domicile, equipe_exterieur = map(int, scoreButBetToTaux.split("-"))
                                                            total_buts = equipe_domicile + equipe_exterieur
                                                            if total_buts < 2:
                                                                total_moins_de_2_buts +=1
                                                            else:
                                                                total_2_buts_ou_plus  += 1
                                                
                                                ws_buts_total_bet365.append_row(["","","","",total_moins_de_2_buts,total_2_buts_ou_plus,TotalButBet365])
                                            else:
                                                # Si aucun '/' alors on vérifie si le nombre unique est entre 2 et 3
                                                unique_nombre = float(nombreButBet)
                                                if 2 <= unique_nombre <= 3:
                                                        plusButBet = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_plusBut365))).text.strip().replace('.', '')
                                                        moinsButBet = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_moinsBut365))).text.strip().replace('.', '')
                                                        TotalButBet365 = f"{plusButBet}/{moinsButBet}"

                                                        for row in ws_comparison[1:]:

                                                            cell_a = row[0]
                                                            cell_c = row[4]   

                                                            if cell_c == TotalButBet365:
                                                                scoreButBetToTaux = cell_a
                                                                equipe_domicile, equipe_exterieur = map(int, scoreButBetToTaux.split("-"))
                                                                total_buts = equipe_domicile + equipe_exterieur
                                                                if total_buts < 2:
                                                                    total_moins_de_2_buts +=1
                                                                else:
                                                                    total_2_buts_ou_plus  += 1
                                                                
                                                        ws_buts_total_bet365.append_row(["","","","",total_moins_de_2_buts,total_2_buts_ou_plus,TotalButBet365]) 
                                    except Exception:
                                            TotalButBet365 = ''  # En cas d'erreur, TotalButBet365 est vide

                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[0])
                                else:
                                    print("match sans cote")
                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[0])    

                                print(nomEquipe)
                                processed_elements.add(nomEquipe)
                                save_processed_element(nomEquipe, processed_filename)  

                            except Exception as e:
                                print(nomEquipe)
                                remove_element(nomEquipe,processed_filename)
                                print("Fin matchs en direct")
                                driver.quit()
                                return
                            
                        except Exception as e:
                            print("Erreur lors l'ouverture de la nouvelle onglet",e)
                    driver.execute_script("els = document.getElementsByClassName('match-container'); els[els.length-1].scrollIntoView();")
                    time.sleep(5)  # Attendre un peu pour le chargement des nouvelles données

                
        except Exception as e:
            print("Erreur dans le selenium",e)
            
    except Exception as e:
        print(f"Erreur dans le debuts du code",e)

    finally:
        driver.quit()

while True:
    try:
        process_url()
    except Exception:
        print("Erreur")

    time.sleep(10)
