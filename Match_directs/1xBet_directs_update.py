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
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import requests
import os
from datetime import datetime
import re

# Récupérer l'utilisateur courant
user_name = os.getlogin()

# Fonction pour charger les éléments déjà traités depuis un fichier spécifique


def load_processed_elements(filename):
    if os.path.exists(filename):
        with open(filename, 'r', encoding='utf-8') as file:
            return set(file.read().splitlines())
    return set()

# Fonction pour enregistrer les éléments traités dans un fichier spécifique


def save_processed_element(element, filename):
    with open(filename, 'a', encoding="utf-8") as file:
        file.write(f"{element}\n")

# Fonction pour convertir une cote en liste d'entiers


def parse_cote(cote_str):
    return list(map(int, cote_str.split('/')))


def process_data(data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue):
    date_obj = datetime.strptime(date, "%d/%m/%Y")

    if date_obj.month == 2:
        match_info = f"{nomEquipe} vs {deuxiementEquipe}"
        total_but = f"{TotalButBet365}"
        score = ""
        score_info = f"{scoreButBetToTaux} {date}"

        # Vérifier si la ligne existe déjà dans data_test
        if not data_test or data_test[-1][2] != match_info:
            data_test.append(
                [times, ligue, match_info, total_but, score, score_info]),
        else:
            # Ajouter les nouvelles informations sur la même ligne
            data_test[-1].append(score_info)


def process_url():
    # Définir le chemin vers le ChromeDriver
    try:
        chrome_options = Options()
        # Démarrer en mode maximisé
        chrome_options.add_argument("--start-maximized")
        # service = Service(chrome_driver_path)
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        url = "https://www.aiscore.com/fr/"
        driver.get(url)
        scope = ["https://www.googleapis.com/auth/spreadsheets"]

        # 1 xbet
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            f"C:\\Users\\{user_name}\\Desktop\\scrapping_aiscore\\xbet-identifiants.json", scope)

        # ptyhon
        credsPython = ServiceAccountCredentials.from_json_keyfile_name(
            f"C:\\Users\\{user_name}\\Desktop\\scrapping_aiscore\\credentials.json", scope)

        client = gspread.authorize(creds)
        client_eternal = gspread.authorize(credsPython)

        # Sheet pour la match en directs
        sheet_id = "1IfTZnG5ggyU5P-t_5y6BlOEUbfQWsCnk3GQ0ZID6qAs"
        sheet = client.open_by_key(sheet_id)

        # Sheet pour les donnés des matchs en directs
        sheet_data_id = "1ROzI-Xnz-Y4y-QGP_-KsCir3FBquSCNteUIWt2AB5DM"
        sheet_data = client.open_by_key(sheet_data_id)

        # sheet pour les analyses de donnés
        sheet_id = "1lH8i7RCzVkdo4CN71_h13I6yteebocQTqP6ZxHL1nYo"
        sheet_analyse_test = client_eternal.open_by_key(sheet_id)

        try:
            # Tentative de récupération de la feuille "data"
            sheet_result = sheet_analyse_test.worksheet("data")
            print("La feuille 'Résultat' existe déjà.")

        except gspread.exceptions.WorksheetNotFound:
            # Si la feuille n'existe pas, création d'une nouvelle feuille
            sheet_result = sheet_analyse_test.add_worksheet(
                title="data", rows="100", cols="20")
            print("La feuille 'data' a été créée.")

            # Définir l'en-tête de la feuille
            header = ["TIME", "LEAGUE", "MATCH", "Predictions",
                      "SCORE", "COTE MATCH"]
            sheet_result.append_row(header)

        # Donné analyse
        array_analyse = sheet_result.get_all_values()
        wb_comparison = sheet_data.sheet1
        ws_comparison = wb_comparison.get_all_values()

        # wb_comparison =sheet.sheet1
        id_recuperation = "1li-ihCIXsx-L9WGMxdbKq1yASnIGIKd_AWnHPR3NJ18"
        sheet_recuperation = client.open_by_key(id_recuperation)

        # Home
        ws_home_150 = sheet_recuperation.sheet1
        ws_home_150_200 = sheet_recuperation.worksheet(
            "Domicile 150 < x <= 200")
        ws_home_200 = sheet_recuperation.worksheet("Domicile  x > 200")
        # Away
        ws_away_150 = sheet_recuperation.worksheet("Extérieur x <= 150")
        ws_away_150_200 = sheet_recuperation.worksheet(
            "Extérieur 150 < x <= 200")
        ws_away_200 = sheet_recuperation.worksheet("Extérieur > 200")

        # Donné Home
        wb_home_150 = ws_home_150.get_all_values()
        wb_home_150_200 = ws_home_150_200.get_all_values()
        wb_home_200 = ws_home_200.get_all_values()

        # Donné away
        wb_away_150 = ws_away_150.get_all_values()
        wb_away_150_200 = ws_away_150_200.get_all_values()
        wb_away_200 = ws_away_200.get_all_values()

        processed_filename = f"C:\\Users\\{user_name}\\Desktop\\scrapping_aiscore\\Match_directs\\processed_elements.txt"
        # Créer ou accéder à des onglets dans Google Sheets

        try:
            ws_main = sheet.worksheet("Match_Data")  # Onglet principal
            rows_main = ws_main.get_all_values()
        except gspread.exceptions.WorksheetNotFound:
            ws_main = sheet.add_worksheet(
                title="Match_Data", rows=1000, cols=40)

            # Ligne principale (colonnes générales avec les mois)
            ws_main.append_row([
                "Date", "Time", "League", "Match", "Score", "1XBET ODDS", "Match Date",
                "1XBET O/U 2.5", "Approach", "Score", "Match Date",
                "September", ".", "October", ".", "November", ".", "December", ".", "January", "."
            ])

            # Ligne secondaire (sous-colonnes UNDER et OVER alignées sous les mois)
            ws_main.append_row([
                "", "", "", "", "", "", "", "", "", "", "",
                "UNDER", "OVER", "UNDER", "OVER", "UNDER", "OVER", "UNDER", "OVER", "UNDER", "OVER"
            ])

            # Fusion des cellules pour chaque mois
            ws_main.merge_cells("L1:M1")  # Fusion pour "Septembre"
            ws_main.merge_cells("N1:O1")  # Fusion pour "Octobre"
            ws_main.merge_cells("P1:Q1")  # Fusion pour "Novembre"
            ws_main.merge_cells("R1:S1")  # Fusion pour "Décembre"
            ws_main.merge_cells("T1:U1")  # Fusion pour "January"

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
            date_deja_Result = False

            # values_only=True pour obtenir les valeurs sans les objets cellule
            for row in rows_main[1:]:
                # Supposons que la date se trouve dans la première colonne
                cell_date = row[0]
                if cell_date == aujourd_hui_str:
                    date_deja_presenteMain = True
                    break  # On peut sortir de la boucle si on trouve la date

            for row in array_analyse:
                # Supposons que la date se trouve dans la première colonne
                cell_date = row[0]
                print(cell_date)
                if cell_date == aujourd_hui_str:
                    date_deja_Result = True
                    break  # On peut sortir de la boucle si on trouve la date

            data_array = []
            data_test = []

            # Ajouter la date si elle n'est pas déjà présente
            if not date_deja_presenteMain:
                ws_main.append_row([aujourd_hui_str, "", "", "", "", "",
                                   "", "", "", "", "", "", "", "", "", "", "", "", "", ""])
            else:
                print("La date est déjà présente Main.")

            # Ajouter la date si elle n'est pas déjà présente
            if not date_deja_Result:
                sheet_result.append_row([aujourd_hui_str, ""])
            else:
                print("La date est déjà présente Main.")

            WebDriverWait(driver, 2).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'span.el-checkbox__inner[data-v-d4c6fef0]'))).click()
            time.sleep(5)  # Attendre un peu pour permettre le chargement

            script = """
                    window.lastScrollPosition = 0;  // Initialisation à 0
                    window.addEventListener('scroll', () => {
                    window.lastScrollPosition = window.scrollY;  // Stockage dans window
                    });
                    """

            driver.execute_script(script)
            print("Écouteur de scroll ajouté.")

            while True:
                try:
                    match_elements = driver.find_elements(
                        By.CSS_SELECTOR, 'a.match-container')

                    if not match_elements:
                        print(f"pas de matchs")
                        time.sleep(10)

                    max_no_team_count = 4  # e d'occurrences successives sans nom d'équipe
                    no_team_counter = 0    # Compteur d'occurrences successives sans nom d'équipe

                    # Traiter les éléments de match
                    for element in match_elements:
                        # temps_home = element.find_element(By.CSS_SELECTOR, 'span.status.minitext.on').text.strip()
                        nomEquipe = element.find_element(
                            By.CSS_SELECTOR, 'span.name.minitext.maxWidth160').text.strip()
                        times = element.find_element(
                            By.CSS_SELECTOR, 'span.time.minitext').text

                        scroll_position = driver.execute_script(
                            "return window.lastScrollPosition;")

                        if scroll_position is None:
                            scroll_position = 0  # Par sécurité, éviter les None

                        # Vérifier si l'équipe a déjà été traitée
                        if nomEquipe in processed_elements:
                            print(
                                f"L'équipe {nomEquipe} a déjà été traitée, passage au suivant")
                            continue  # Passe au prochain élément

                        if not element.is_displayed():
                            print(
                                "L'élément n'est plus visible, passage au suivant.")
                            continue

                        if not nomEquipe:
                            no_team_counter += 1
                            print(
                                f"Aucun nom d'équipe trouvé pour cet élément. ({no_team_counter}/{max_no_team_count})")

                            if no_team_counter >= max_no_team_count and scroll_position == 0:
                                print(
                                    "Trop d'éléments sans nom d'équipe, arrêt du traitement.")
                                return  # Arrête la boucle si le compteur atteint la limite

                            continue  # Passe au prochain élément

                        # Réinitialiser le compteur si un nom d'équipe est trouvé
                        no_team_counter = 0

                        if not times:
                            print(
                                f"Aucun temps trouvé pour l'équipe {nomEquipe}")
                            continue

                        href = element.get_attribute("href")
                        # new_data_found = True
                        try:
                            # Ouvre le lien et récupère les données
                            driver.execute_script(
                                "window.open(arguments[0]);", href)
                            driver.switch_to.window(driver.window_handles[-1])
                            thirst_selector = "#app > div.detail.view.border-box.back > div.tab-bar > div > div > a:nth-child(2)"
                            thirst_selectorText = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, thirst_selector))).text.strip()
                            temps = "#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.flex.w-bar-100.homeBox > div.h-top-center.matchStatus2 > div.flex-1.text-center.scoreBox > div.h-16.m-b-4 > span > span:nth-child(1)"

                            ligue = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                (By.CSS_SELECTOR, "div.top.color-333.flex-col.flex.align-center > div.comp-name > a"))).text.strip()

                            try:
                                temps__text = WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, temps)))
                                try:
                                    if thirst_selectorText == "Cotes":

                                        deuxiemeEquipe__selector = "#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.flex.w-bar-100.homeBox > div.away-box > div > a"
                                        deuxiementEquipe = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, deuxiemeEquipe__selector))).text.strip()

                                        # ws_main.append_row(["",times,ligue,f"{nomEquipe} vs {deuxiementEquipe}", "", "", "","", "","", "","", "", "", "", "", "", "", ""])
                                        data_array.append(
                                            ["", times, ligue, f"{nomEquipe} vs {deuxiementEquipe}", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""])

                                        cote_selectorBet365_1 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.flex.odds > div.table.flex-1.eu > div:nth-child(3) > div > div:nth-child(1) > span'
                                        cote_selectorBet365_2 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.flex.odds > div.table.flex-1.eu > div:nth-child(3) > div > div:nth-child(2) > span'
                                        cote_selectorBet365_3 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.flex.odds > div.table.flex-1.eu > div:nth-child(3) > div > div:nth-child(3) > span'
                                        cote_nombreButBet365 = "#app > div.detail.view.border-box.back > div.content-box > span > div > div.flex.odds > div.table.flex-1.bs > div:nth-child(3) > div > div:nth-child(1) > span"
                                        cote_plusBut365 = "#app > div.detail.view.border-box.back > div.content-box > span > div > div.flex.odds > div.table.flex-1.bs > div:nth-child(3) > div > div:nth-child(2) > span"
                                        cote_moinsBut365 = "#app > div.detail.view.border-box.back > div.content-box > span > div > div.flex.odds > div.table.flex-1.bs > div:nth-child(3) > div > div:nth-child(3) > span"

                                        # Ligue ="#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.comp-name > a"

                                        try:
                                            cote_Bet365 = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                                (By.CSS_SELECTOR, cote_selectorBet365_1))).text.strip().replace('.', '')
                                            cote_Bet365_2 = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                                (By.CSS_SELECTOR, cote_selectorBet365_2))).text.strip().replace('.', '')
                                            cote_Bet365_3 = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                                (By.CSS_SELECTOR, cote_selectorBet365_3))).text.strip().replace('.', '')
                                            cotesBet365 = f"{cote_Bet365}/{cote_Bet365_2}/{cote_Bet365_3}"

                                            # Commence à la ligne 2 pour ignorer l'en-tête
                                            for row in ws_comparison[1:]:
                                                # On récupère la cellule dans la colonne B de chaque ligne
                                                cell_a = row[0]
                                                cell_b = row[1]
                                                cell_c = row[3]

                                                if cell_b == cotesBet365:
                                                    scoreBet365 = cell_a
                                                    dateMatchs = cell_c
                                                    # ws_main.append_row(["", "", "", "",scoreBet365, cotesBet365,dateMatchs,"", "", "", "", "", "", "", "", "", "", "", ""])
                                                    data_array.append(
                                                        ["", "", "", "", scoreBet365, cotesBet365, dateMatchs, "", "", "", "", "", "", "", "", "", "", "", "", "", ""])
                                                    # Colonne de score à gauche, colonne de cote 1xBet vide pour l'instant

                                        except Exception:
                                            cotesBet365 = ''

                                        try:
                                            # Récupère la donnée brute de nombre de buts
                                            nombreButBet = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                                (By.CSS_SELECTOR, cote_nombreButBet365))).text.strip()
                                            # Initialisation des valeurs
                                            TotalButBet365 = ''
                                            # ws_buts_total_bet365.append_row(["",times,ligue,f"{nomEquipe} vs {deuxiementEquipe}","","", ""])  # Ajoute une ligne vide pour séparer les matchs
                                            total_moins_de_3_buts = 0
                                            total_3_buts_ou_plus = 0
                                            tolerance = 10

                                            if '/' in nombreButBet:
                                                # Sépare les deux valeurs
                                                valeurs = nombreButBet.split(
                                                    '/')
                                                premier_nombre = float(
                                                    valeurs[0])
                                                second_nombre = float(
                                                    valeurs[1])

                                                # Vérifie si les deux valeurs sont entre 2 et 3
                                                if 2 <= premier_nombre <= 3 and 2 <= second_nombre <= 3.0:
                                                    # Récupère les valeurs pour plus et moins si les conditions sont respectées
                                                    plusButBet = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                                        (By.CSS_SELECTOR, cote_plusBut365))).text.strip().replace('.', '')
                                                    moinsButBet = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                                        (By.CSS_SELECTOR, cote_moinsBut365))).text.strip().replace('.', '')
                                                    TotalButBet365 = f"{plusButBet}/{moinsButBet}"
                                                    # Marge de tolérance définie, par exemple, 10 unités
                                                    tolerance = 5

                                                    # On suppose que 'ws_comparison' contient les données récupérées du Google Sheet
                                                    # Une seule boucle
                                                    for row in ws_comparison[1:]:
                                                        # Score, par exemple: '0-1'
                                                        cell_a = row[0]
                                                        # Cotes, par exemple: '140/414/995'
                                                        cell_b = row[1]
                                                        # Cotes '1XBET O/U 2.5', par exemple: '205/176'
                                                        cell_c = row[2]
                                                        # La date, par exemple: '20/09/2024'
                                                        cell_d = row[3]

                                                        if cell_c == TotalButBet365:
                                                            scoreButBetToTaux = cell_a
                                                            cote_1xBet = cell_b
                                                            date = cell_d
                                                            # Séparer les cotes de Bet365
                                                            favoris = cotesBet365.split(
                                                                '/')

                                                            try:
                                                                # On suppose que les cotes sont au format '100/200/300'
                                                                # Première cote (domicile)
                                                                premiere_favoris = int(
                                                                    favoris[0])
                                                                # Troisième cote (extérieur)
                                                                troisieme_favoris = int(
                                                                    favoris[-1])
                                                                # Comparer les cotes dans ws_comparison
                                                                # Séparer la première cote de la ligne du sheet
                                                                cote_comparaison = cote_1xBet.split(
                                                                    '/')
                                                                try:
                                                                    # Cote de domicile
                                                                    cote_dom = int(
                                                                        cote_comparaison[0])
                                                                    # Cote extérieure
                                                                    cote_ext = int(
                                                                        cote_comparaison[-1])
                                                                    # Comparaison en fonction du favori
                                                                    if premiere_favoris < troisieme_favoris:

                                                                        # Fvoris domicile : comparer la cote de domicile7
                                                                        diff_dom = abs(
                                                                            cote_dom - premiere_favoris)
                                                                        if diff_dom <= tolerance:

                                                                            if (cote_dom <= 150):
                                                                                for row in wb_home_150[2:]:

                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])

                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                            elif (150 < cote_dom <= 200):
                                                                                for row in wb_home_150_200[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])

                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                            else:
                                                                                for row in wb_home_200[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])

                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                    else:
                                                                        diff_ext = abs(
                                                                            cote_ext - troisieme_favoris)

                                                                        if diff_ext <= tolerance:
                                                                            if (cote_ext <= 150):
                                                                                for row in wb_away_150[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])

                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                            elif (150 < cote_ext <= 200):
                                                                                for row in wb_away_150_200[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])
                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                            else:
                                                                                for row in wb_away_200[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])
                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                except ValueError:
                                                                    print(
                                                                        "Erreur dans la conversion des cotes en entiers:", cote_comparaison)

                                                            except ValueError:
                                                                print(
                                                                    "Erreur dans la conversion des cotes en entiers:", favoris)

                                            else:
                                                # Si aucun '/' alors on vérifie si le nombre unique est entre 2 et 3
                                                unique_nombre = float(
                                                    nombreButBet)
                                                if 2 <= unique_nombre <= 3:
                                                    plusButBet = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                                        (By.CSS_SELECTOR, cote_plusBut365))).text.strip().replace('.', '')
                                                    moinsButBet = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                                        (By.CSS_SELECTOR, cote_moinsBut365))).text.strip().replace('.', '')
                                                    TotalButBet365 = f"{plusButBet}/{moinsButBet}"
                                                    # Marge de tolérance définie, par exemple, 10 unités
                                                    tolerance = 5

                                                    # On suppose que 'ws_comparison' contient les données récupérées du Google Sheet
                                                    # Une seule boucle
                                                    for row in ws_comparison[1:]:
                                                        # Score, par exemple: '0-1'
                                                        cell_a = row[0]
                                                        # Cotes, par exemple: '140/414/995'
                                                        cell_b = row[1]
                                                        # Cotes '1XBET O/U 2.5', par exemple: '205/176'
                                                        cell_c = row[2]
                                                        # La date, par exemple: '20/09/2024'
                                                        cell_d = row[3]

                                                        if cell_c == TotalButBet365:
                                                            scoreButBetToTaux = cell_a
                                                            cote_1xBet = cell_b
                                                            date = cell_d
                                                            # Séparer les cotes de Bet365
                                                            favoris = cotesBet365.split(
                                                                '/')

                                                            try:
                                                                # On suppose que les cotes sont au format '100/200/300'
                                                                # Première cote (domicile)
                                                                premiere_favoris = int(
                                                                    favoris[0])
                                                                # Troisième cote (extérieur)
                                                                troisieme_favoris = int(
                                                                    favoris[2])

                                                                # Comparer les cotes dans ws_comparison
                                                                # Séparer la première cote de la ligne du sheet
                                                                cote_comparaison = cote_1xBet.split(
                                                                    '/')
                                                                try:
                                                                    # Cote de domicile
                                                                    cote_dom = int(
                                                                        cote_comparaison[0])
                                                                    # Cote extérieure
                                                                    cote_ext = int(
                                                                        cote_comparaison[-1])

                                                                    # Comparaison en fonction du favori
                                                                    if premiere_favoris < troisieme_favoris:
                                                                        # Favoris domicile : comparer la cote de domicile7
                                                                        diff_dom = abs(
                                                                            cote_dom - premiere_favoris)
                                                                        if diff_dom <= tolerance:
                                                                            if (cote_dom <= 150):
                                                                                for row in wb_home_150[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])
                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                            elif (150 < cote_dom <= 200):
                                                                                for row in wb_home_150_200[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])
                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                            else:
                                                                                for row in wb_home_200[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])
                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                    else:
                                                                        diff_ext = abs(
                                                                            cote_ext - troisieme_favoris)
                                                                        if diff_ext <= tolerance:
                                                                            if (cote_ext <= 150):
                                                                                for row in wb_away_150[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])
                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                            elif (150 < cote_ext <= 200):
                                                                                for row in wb_away_150_200[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])
                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                            else:
                                                                                for row in wb_away_200[2:]:
                                                                                    if row[0] == TotalButBet365:
                                                                                        data_array.append(["", "", "", "", "", "", "", TotalButBet365, cote_1xBet, scoreButBetToTaux, date, row[
                                                                                                          1], row[2],  row[3],  row[4], row[5], row[6], row[7], row[8], row[9], row[10]])
                                                                                        process_data(
                                                                                            data_test, TotalButBet365, scoreButBetToTaux, date, nomEquipe, deuxiementEquipe, times, ligue)

                                                                except ValueError as e:
                                                                    print(
                                                                        "Erreur dans la conversion des cotes en entiers:", e, cote_comparaison)

                                                            except ValueError as e:
                                                                print(
                                                                    "Erreur dans la conversion des cotes en entiers:", e, favoris)

                                        except Exception:
                                            TotalButBet365 = ''  # En cas d'erreur, TotalButBet365 est vide

                                        driver.close()
                                        driver.switch_to.window(
                                            driver.window_handles[0])
                                    else:
                                        print("match sans cote")
                                        driver.close()
                                        driver.switch_to.window(
                                            driver.window_handles[0])

                                    print(nomEquipe, "ajouter")
                                    ws_main.append_rows(data_array)
                                    sheet_result.append_rows(data_test)
                                    data_array = []
                                    data_test = []
                                    save_processed_element(
                                        nomEquipe, processed_filename)
                                    processed_elements.add(nomEquipe)

                                except Exception as e:
                                    print("Error", e)
                                    continue

                            except Exception as e:
                                print(nomEquipe, "match pas en courts")
                                driver.quit()
                                return

                        except Exception as e:
                            print("match sans cote")
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])

                    driver.execute_script(
                        "els = document.getElementsByClassName('match-container'); els[els.length-1].scrollIntoView();")
                    # Attendre un peu pour le chargement des nouvelles données
                    time.sleep(3)

                except (NoSuchElementException, StaleElementReferenceException) as e:
                    print(
                        f"L'élément a été supprimé ou n'est plus présent: {e}. Passage au suivant.")
                    continue  # Passe au prochain élément

        except Exception as e:
            print("Erreur dans le selenium", e)

    except Exception as e:
        print(f"Erreur dans le debuts du code", e)

    finally:
        print(f"script fin")
        driver.quit()


while True:
    try:
        process_url()
    except Exception:
        print("Erreur")

    time.sleep(60)
