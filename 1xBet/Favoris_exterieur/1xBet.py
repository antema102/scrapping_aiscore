import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import time
import requests
import os
from datetime import datetime, timedelta


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

# Fonction pour extraire les données et les enregistrer dans un fichier Excel


def process_url():
    try:
        # Intialisations
        chrome_driver_path = r"C:\Users\Administrator\Desktop\scrapping_aiscore\chromedriver\chromedriver.exe"
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)

        # Google sheets
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            r'C:\Users\Administrator\Desktop\scrapping_aiscore\xbet-identifiants.json', scope)
        client = gspread.authorize(creds)
        sheet_id = "1ROzI-Xnz-Y4y-QGP_-KsCir3FBquSCNteUIWt2AB5DM"
        sheet = client.open_by_key(sheet_id)
        wb = sheet.sheet1
        ws = wb.get_all_values()

        # Date
        hier = datetime.today().date()-timedelta(days=1)
        hier_str = hier.strftime('%Y%m%d')
        dateHier = hier.strftime('%d/%m/%Y')
        url = f"https://www.aiscore.com/fr/{hier_str}"
        processed_filename = f"C:\\Users\\Administrator\\Desktop\\scrapping_aiscore\\1xBet\\Favoris_exterieur\\processed_elements_{hier_str}.txt"
        date_present = False

        for row in ws[1:]:
            cell_score = row[0]
            if cell_score == dateHier:
                date_present = True
                break

        if not date_present:
            wb.append_row([dateHier, "", "", "",])
        else:
            print("date deja présent")

        # Charger Elements
        processed_elements = load_processed_elements(processed_filename)

        driver.get(url)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#app")))

        tous_selector = "#app > DIV:nth-of-type(3) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(1) > SPAN:nth-of-type(1)"
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, tous_selector))).click()

        second_selector = "#app > DIV:nth-of-type(3) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > DIV:nth-of-type(2) > LABEL > SPAN > SPAN"
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, second_selector))).click()

        while True:

            match_elements = driver.find_elements(
                By.CSS_SELECTOR, 'a.match-container')
            new_data_found = False
            # Traiter les éléments de match

            for element in match_elements:
                nomEquipe = element.find_element(
                    By.CSS_SELECTOR, 'span.name.minitext.maxWidth160').text.strip()

                if nomEquipe in processed_elements:
                    print(
                        f"L'équipe {nomEquipe} a déjà été traitée, passage au suivant.")
                    continue  # Passe au prochain élément sans traiter

                href = element.get_attribute("href")
                processed_elements.add(nomEquipe)
                save_processed_element(
                    nomEquipe, processed_filename)

                try:
                    # Ouvre le lien et récupère les données
                    driver.execute_script(
                        "window.open(arguments[0]);", href)
                    driver.switch_to.window(driver.window_handles[-1])

                    # Cliquer sur la troisieme bouton
                    thirst_selector = "#app > div.detail.view.border-box.back > div.tab-bar > div > div > a:nth-child(2)"
                    thirst_selectorText = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, thirst_selector))).text.strip()

                    if thirst_selectorText == "Cotes":

                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, thirst_selector))).click()

                        # affiche les cotes lorsque je clique sur le bouton cote
                        score1 = '#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.flex.w-bar-100.homeBox > div.h-top-center.matchStatus3 > div.font-bold.home-score > span:last-child'
                        score2 = '#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.flex.w-bar-100.homeBox > div.h-top-center.matchStatus3 > div.font-bold.away-score > span'

                        cote_selector1xBet_1 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(1) > span > span'
                        cote_selector1xBet_2 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(2) > span'
                        cote_selector1xBet_3 = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(1) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(3) > span > span'

                        cote_nombreBut1xBet = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(1) > span'
                        cote_plusBut1xBet = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(2) > span'
                        cote_moinBut1xBet = '#app > div.detail.view.border-box.back > div.content-box > span > div > div.newOdds > div:nth-child(3) > div:nth-child(1) > div.flex-1 > div > div:nth-child(3) > div.box.flex.w100.brr.preMatchBg1 > div > div:nth-child(3) > span'

                        try:
                            score_text = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, score1))).text.strip()
                            score_text1 = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, score2))).text.strip()

                            try:
                                cote_1xBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, cote_selector1xBet_1))).text.strip().replace('.', '')
                                cote_1xBet_2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, cote_selector1xBet_2))).text.strip().replace('.', '')
                                cote_1xBet_3 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, cote_selector1xBet_3))).text.strip().replace('.', '')
                                cotes1xBet = f"{cote_1xBet}/{cote_1xBet_2}/{cote_1xBet_3}"
                            except Exception:
                                cotes1xBet = ''

                            try:
                                # Récupère la donnée brute de nombre de buts
                                nombreBut1xBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, cote_nombreBut1xBet))).text.strip()

                                # Initialisation de la variable
                                TotalBut1xBet = ''

                                # Vérifie s'il y a un '/' pour extraire les deux nombres
                                if '/' in nombreBut1xBet:
                                    # Sépare les deux valeurs
                                    valeurs = nombreBut1xBet.split('/')
                                    premier_nombre = float(valeurs[0])
                                    second_nombre = float(valeurs[1])

                                    # Vérifie si les deux valeurs sont entre 2 et 3
                                    if 2 <= premier_nombre <= 3 and 2 <= second_nombre <= 3:
                                        # Récupère les valeurs pour plus et moins si les conditions sont respectées
                                        plusBut1xBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, cote_plusBut1xBet))).text.strip().replace('.', '')
                                        moinsBut1xBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, cote_moinBut1xBet))).text.strip().replace('.', '')
                                        TotalBut1xBet = f"{plusBut1xBet}/{moinsBut1xBet}"
                                else:
                                    # Si aucun '/' alors on vérifie si le nombre unique est entre 2 et 3
                                    unique_nombre = float(
                                        nombreBut1xBet)
                                    if 2 <= unique_nombre <= 3:
                                        plusBut1xBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, cote_plusBut1xBet))).text.strip().replace('.', '')
                                        moinsBut1xBet = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, cote_moinBut1xBet))).text.strip().replace('.', '')
                                        TotalBut1xBet = f"{plusBut1xBet}/{moinsBut1xBet}"

                            except Exception:
                                TotalBut1xBet = ''  # En cas d'erreur, TotalBut1xBet est vide

                            wb.append_row(
                                [f"{score_text}-{score_text1}", cotes1xBet, TotalBut1xBet, dateHier])
                            print(f"{score_text}-{score_text1}",
                                  cotes1xBet, TotalBut1xBet, dateHier)

                        except Exception as e:
                            print("match reporté")

                        driver.close()
                        driver.switch_to.window(
                            driver.window_handles[0])
                    else:
                        driver.close()
                        driver.switch_to.window(
                            driver.window_handles[0])

                except Exception as e:
                    print('erreur pour louveture tabs', e)
                    driver.quit()
                    
            driver.execute_script(
                "els = document.getElementsByClassName('match-container'); els[els.length-1].scrollIntoView();")
            # Attendre un peu pour le chargement des nouvelles données
            time.sleep(3)

    except Exception as e:
        print("Erreur dans le debuts", e)
        driver.quit()
    
    finally:
        print(f"script fin")
        driver.quit()
        
while True:
    try:
        print("Relance")
        process_url()

    except Exception as e:
        print('erreur')
