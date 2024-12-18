# import openpyxl
# from openpyxl import Workbook,load_workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
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
# Fonction pour convertir une cote en liste d'entiers
def parse_cote(cote_str):
    return list(map(int, cote_str.split('/')))

def process_url():
    # Définir le chemin vers le ChromeDriver
    try:
        chrome_driver_path = r"C:\Users\antem\Desktop\scrapping_aiscore\chromedriver\chromedriver.exe" 
        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service)
        url="https://www.aiscore.com/fr/"
        driver.get(url)
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\antem\Desktop\scrapping_aiscore/credentials.json", scope)
        client = gspread.authorize(creds)

        sheet_id = "13h2YXSaJcKjaNa6PO19AJphd-s1VxdO4aBwORng_6Qs"
        sheet = client.open_by_key(sheet_id)

        sheet_data_id="1ROzI-Xnz-Y4y-QGP_-KsCir3FBquSCNteUIWt2AB5DM"
        sheet_data=client.open_by_key(sheet_data_id)

        wb_comparison=sheet_data.sheet1
        ws_comparison = wb_comparison.get_all_values()

        # wb_comparison =sheet.sheet1
        id_recuperation="19F_f1opf_dWCzhFsLe-TcF9nHdOyUomuU7V8p_V0W4c"
        sheet_recuperation=client.open_by_key(id_recuperation)

        #Home
        ws_home_150 = sheet_recuperation.sheet1
        ws_home_150_200 = sheet_recuperation.worksheet("Domicile 150 < x <= 200")
        ws_home_200 =sheet_recuperation.worksheet("Domicile  x > 200")
        #Away
        ws_away_150 = sheet_recuperation.worksheet("Extérieur x <= 150")
        ws_away_150_200 = sheet_recuperation.worksheet("Extérieur 150 < x <= 200")
        ws_away_200 =sheet_recuperation.worksheet("Extérieur > 200")

        #Donné Home
        wb_home_150= ws_home_150.get_all_values()
        wb_home_150_200= ws_home_150_200.get_all_values()
        wb_home_200= ws_home_200.get_all_values()

        #Donné away
        wb_away_150= ws_away_150.get_all_values()
        wb_away_150_200= ws_away_150_200.get_all_values()
        wb_away_200= ws_away_200.get_all_values()

        processed_filename = f"C:\\Users\\antem\\Desktop\\scrapping_aiscore\\Match_directs\\processed_elements.txt"
        # Créer ou accéder à des onglets dans Google Sheets
        try:
            ws_main = sheet.worksheet("Match Data")  # Onglet principal
            rows_main = ws_main.get_all_values()
        except gspread.exceptions.WorksheetNotFound:
            ws_main = sheet.add_worksheet(title="Match Data",rows=1000,cols=40)

            # Ligne principale (colonnes générales avec les mois)
            ws_main.append_row([
                "Date", "Heure", "Ligue", "Match", "Score", "1XBET ODDS", "Date Matchs",
                "1XBET O/U 2.5", "Rapproche", "Score", "Date Matchs",
                "Septembre", ".", "Octobre", ".", "Novembre", ".", "Décembre", "."
            ])

            # Ligne secondaire (sous-colonnes UNDER et OVER alignées sous les mois)
            ws_main.append_row([
                "", "", "", "", "", "", "", "", "", "", "",
                "UNDER", "OVER", "UNDER", "OVER", "UNDER", "OVER", "UNDER", "OVER"
            ])

            # Fusion des cellules pour chaque mois
            ws_main.merge_cells("L1:M1")  # Fusion pour "Septembre"
            ws_main.merge_cells("N1:O1")  # Fusion pour "Octobre"
            ws_main.merge_cells("P1:Q1")  # Fusion pour "Novembre"
            ws_main.merge_cells("R1:S1")  # Fusion pour "Décembre"

            # Formatage pour centrer les mois
            ws_main.format("A1:S1000", {
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE"
            })

        rows_main = ws_main.get_all_values()
        processed_elements = load_processed_elements(processed_filename)

        try:
            # Initialiser la variable date_deja_presente
            aujourd_hui = datetime.today().date()
            aujourd_hui_str = aujourd_hui.strftime('%Y/%m/%d')
            date_deja_presenteMain = False

            for row in rows_main[1:]:  # values_only=True pour obtenir les valeurs sans les objets cellule
                cell_date = row[0]  # Supposons que la date se trouve dans la première colonne
                if cell_date == aujourd_hui_str:
                    date_deja_presenteMain = True
                    break  # On peut sortir de la boucle si on trouve la date

            data_array=[] 

            # Ajouter la date si elle n'est pas déjà présente
            if not date_deja_presenteMain:              
                data_array.append([aujourd_hui_str,"","","","","","","","","","","","","","","","","",""])
            else:
                print("La date est déjà présente Main.")

            # second_selector = "#app > DIV:nth-of-type(3) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > DIV:nth-of-type(2) > LABEL > SPAN > SPAN"
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'span.el-checkbox__inner[data-v-d4c6fef0]'))).click()
            time.sleep(5) # Attendre un peu pour permettre le chargement

            while True: 
                    try:              
                        match_elements = driver.find_elements(By.CSS_SELECTOR, 'a.match-container')
                        if match_elements :
                            # Traiter les éléments de match
                            for element in match_elements:
                                # temps_home = element.find_element(By.CSS_SELECTOR, 'span.status.minitext.on').text.strip()
                                nomEquipe = element.find_element(By.CSS_SELECTOR, 'span.name.minitext.maxWidth160').text.strip()
                                times = element.find_element(By.CSS_SELECTOR,'span.time.minitext').text

                                if not nomEquipe:
                                        print(f"Aucun nom d'équipe trouvé pour cet élément.")
                                        continue

                                if nomEquipe in processed_elements:
                                        print(f"L'équipe {nomEquipe} a déjà été traitée, passage au suivant.")
                                        continue  # Passe au prochain élément sans traiter

                                if not times:
                                        print(f"Aucun temps trouvé pour l'équipe {nomEquipe}")
                                        continue

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
                                                    deuxiementEquipe = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, deuxiemeEquipe__selector))).text.strip()

                                                    # ws_main.append_row(["",times,ligue,f"{nomEquipe} vs {deuxiementEquipe}", "", "", "","", "","", "","", "", "", "", "", "", "", ""])
                                                    data_array.append(["",times,ligue,f"{nomEquipe} vs {deuxiementEquipe}", "", "", "","", "","", "","", "", "", "", "", "", "", ""])

                                                    cote_selectorBet365_1 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(2) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(1) > span > span'
                                                    cote_selectorBet365_2 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(2) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(2) > span'
                                                    cote_selectorBet365_3 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(2) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(3) > span > span'
                                                    cote_nombreButBet365  = "#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(2) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(1) > span"
                                                    cote_plusBut365       = "#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(2) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(2) > span"
                                                    cote_moinsBut365      = "#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(2) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(3) > span"

                                                    # Ligue ="#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.comp-name > a"

                                                    try:
                                                            cote_Bet365   = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_selectorBet365_1))).text.strip().replace('.', '')
                                                            cote_Bet365_2 = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_selectorBet365_2))).text.strip().replace('.', '')
                                                            cote_Bet365_3 = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_selectorBet365_3))).text.strip().replace('.', '')
                                                            cotesBet365   = f"{cote_Bet365}/{cote_Bet365_2}/{cote_Bet365_3}"
                                                            
                                                            for row in ws_comparison[1:]:  # Commence à la ligne 2 pour ignorer l'en-tête
                                                                # On récupère la cellule dans la colonne B de chaque ligne
                                                                cell_a = row[0]  
                                                                cell_b = row[1]
                                                                cell_c = row[3]

                                                                if cell_b == cotesBet365:
                                                                    scoreBet365 = cell_a
                                                                    dateMatchs  = cell_c
                                                                    # ws_main.append_row(["", "", "", "",scoreBet365, cotesBet365,dateMatchs,"", "", "", "", "", "", "", "", "", "", "", ""])
                                                                    data_array.append(["", "", "", "",scoreBet365, cotesBet365,dateMatchs,"", "", "", "", "", "", "", "", "", "", "", ""])
                                                                    # Colonne de score à gauche, colonne de cote 1xBet vide pour l'instant    
                                                                                    
                                                    except Exception:
                                                            cotesBet365 = ''

                                                    try:
                                                            # Récupère la donnée brute de nombre de buts
                                                            nombreButBet = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_nombreButBet365))).text.strip()
                                                            # Initialisation des valeurs
                                                            TotalButBet365 = ''
                                                            # ws_buts_total_bet365.append_row(["",times,ligue,f"{nomEquipe} vs {deuxiementEquipe}","","", ""])  # Ajoute une ligne vide pour séparer les matchs
                                                            total_moins_de_3_buts = 0
                                                            total_3_buts_ou_plus = 0
                                                            tolerance = 10
                                                            
            
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
                                                                    # Marge de tolérance définie, par exemple, 10 unités
                                                                    tolerance = 5

                                                                    # On suppose que 'ws_comparison' contient les données récupérées du Google Sheet
                                                                    for row in ws_comparison[1:]:  # Une seule boucle
                                                                        cell_a = row[0]  # Score, par exemple: '0-1'
                                                                        cell_b = row[1]  # Cotes, par exemple: '140/414/995'
                                                                        cell_c = row[2]  # Cotes '1XBET O/U 2.5', par exemple: '205/176'
                                                                        cell_d = row[3]  # La date, par exemple: '20/09/2024'

                                                                        if cell_c == TotalButBet365:  
                                                                            scoreButBetToTaux = cell_a
                                                                            cote_1xBet = cell_b
                                                                            date=cell_d
                                                                            # Séparer les cotes de Bet365
                                                                            favoris = cotesBet365.split('/')

                                                                            try:
                                                                                # On suppose que les cotes sont au format '100/200/300'
                                                                                premiere_favoris = int(favoris[0])  # Première cote (domicile)
                                                                                troisieme_favoris = int(favoris[-1])  # Troisième cote (extérieur)
                                                                                # Comparer les cotes dans ws_comparison
                                                                                cote_comparaison = cote_1xBet.split('/')  # Séparer la première cote de la ligne du sheet
                                                                                try:
                                                                                    cote_dom = int(cote_comparaison[0])  # Cote de domicile
                                                                                    cote_ext = int(cote_comparaison[-1])  # Cote extérieure
                                                                                    # Comparaison en fonction du favori
                                                                                    if premiere_favoris < troisieme_favoris:
                                                                                                        
                                                                                        # Fvoris domicile : comparer la cote de domicile7
                                                                                            diff_dom = abs(cote_dom - premiere_favoris)
                                                                                            if diff_dom <= tolerance:     

                                                                                                if (cote_dom <= 150 ):
                                                                                                    for row in wb_home_150[2:]:
                                                                                                        if row[0] == TotalButBet365:
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      
                                                                                                elif( 150 < cote_dom <= 200):
                                                                                                    for row in wb_home_150_200[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      
                                                                            
                                                                                                else:
                                                                                                    for row in wb_home_200[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      

                                                                                    else:
                                                                                            diff_ext = abs(cote_ext - troisieme_favoris)

                                                                                            if diff_ext <= tolerance:                                                                                  
                                                                                                if (cote_ext <= 150 ):
                                                                                                    for row in wb_away_150[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      

                                                                                                elif( 150 < cote_ext <= 200):
                                                                                                    for row in wb_away_150_200[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      

                                                                                                else:
                                                                                                    for row in wb_away_200[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      

                                                                                    
                                                                                except ValueError:
                                                                                    print("Erreur dans la conversion des cotes en entiers:", cote_comparaison)

                                                                            except ValueError:
                                                                                print("Erreur dans la conversion des cotes en entiers:", favoris)

                                                            else:
                                                                # Si aucun '/' alors on vérifie si le nombre unique est entre 2 et 3
                                                                unique_nombre = float(nombreButBet)
                                                                if 2 <= unique_nombre <= 3:
                                                                        plusButBet = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_plusBut365))).text.strip().replace('.', '')
                                                                        moinsButBet = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR, cote_moinsBut365))).text.strip().replace('.', '')
                                                                        TotalButBet365 = f"{plusButBet}/{moinsButBet}"
                                                                        # Marge de tolérance définie, par exemple, 10 unités
                                                                        tolerance = 5

                                                                        # On suppose que 'ws_comparison' contient les données récupérées du Google Sheet
                                                                        for row in ws_comparison[1:]:  # Une seule boucle
                                                                            cell_a = row[0]  # Score, par exemple: '0-1'
                                                                            cell_b = row[1]  # Cotes, par exemple: '140/414/995'
                                                                            cell_c = row[2]  # Cotes '1XBET O/U 2.5', par exemple: '205/176'
                                                                            cell_d = row[3]  # La date, par exemple: '20/09/2024'

                                                                            if cell_c == TotalButBet365:  
                                                                                scoreButBetToTaux = cell_a
                                                                                cote_1xBet = cell_b
                                                                                date = cell_d 
                                                                                # Séparer les cotes de Bet365
                                                                                favoris = cotesBet365.split('/')

                                                                                try:
                                                                                    # On suppose que les cotes sont au format '100/200/300'
                                                                                    premiere_favoris = int(favoris[0])  # Première cote (domicile)
                                                                                    troisieme_favoris = int(favoris[2])  # Troisième cote (extérieur)

                                                                                    # Comparer les cotes dans ws_comparison
                                                                                    cote_comparaison = cote_1xBet.split('/')  # Séparer la première cote de la ligne du sheet
                                                                                    try:
                                                                                        cote_dom = int(cote_comparaison[0])  # Cote de domicile
                                                                                        cote_ext = int(cote_comparaison[-1])  # Cote extérieure

                                                                                        # Comparaison en fonction du favori
                                                                                        if premiere_favoris < troisieme_favoris:
                                                                                            # Favoris domicile : comparer la cote de domicile7
                                                                                            diff_dom = abs(cote_dom - premiere_favoris)
                                                                                            if diff_dom <= tolerance:                                                                                  
                                                                                                if (cote_dom <= 150 ):
                                                                                                    for row in wb_home_150[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      

                                                                                                elif( 150 < cote_dom <= 200):
                                                                                                    for row in wb_home_150_200[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      

                                                                                                else:
                                                                                                    for row in wb_home_200[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      

                                                                                        else:
                                                                                            diff_ext = abs(cote_ext - troisieme_favoris)
                                                                                            if diff_ext <= tolerance:                                                                                  
                                                                                                if (cote_ext <= 150 ):
                                                                                                    for row in wb_away_150[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      


                                                                                                elif( 150 < cote_ext <= 200):
                                                                                                    for row in wb_away_150_200[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      

                                                                                                else:
                                                                                                    for row in wb_away_200[2:]:
                                                                                                        if row[0] == TotalButBet365: 
                                                                                                            data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet,scoreButBetToTaux,date, row[1] , row[2],  row[3],  row[4], row[5], row[6],row[7], row[8]])                                      


                                                                                    except ValueError:
                                                                                        print("Erreur dans la conversion des cotes en entiers:", cote_comparaison)

                                                                                except ValueError:
                                                                                    print("Erreur dans la conversion des cotes en entiers:", favoris)
                
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
                                            continue
                                        
                                except Exception as e:
                                    print("match sans cote")
                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[0])    
                    
                            driver.execute_script("els = document.getElementsByClassName('match-container'); els[els.length-1].scrollIntoView();")
                            time.sleep(3)  # Attendre un peu pour le chargement des nouvelles données
                        else:
                            print(f"pas de matchs")    
                    except (NoSuchElementException, StaleElementReferenceException) as e:
                            print(f"L'élément a été supprimé ou n'est plus présent: {e}. Passage au suivant.")
                            continue  # Passe au prochain élément
            
        except Exception as e:
            print("Erreur dans le selenium",e)
            
    except Exception as e:
        print(f"Erreur dans le debuts du code",e)

    finally:
        ws_main.append_rows(data_array)
        driver.quit()

process_url()

while True:
    try:
        process_url()
    except Exception:
        print("Erreur")

    time.sleep(10)
