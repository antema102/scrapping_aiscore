# import openpyxl
# from openpyxl import Workbook,load_workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import os
import requests
from datetime import datetime, timedelta

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


def check_internet(url="https://www.google.com", timeout=5):
    """Teste la connexion Internet en envoyant une requête à Google."""
    try:
        response = requests.get(url, timeout=timeout)
        return response.status_code == 200
    except requests.ConnectionError:
        return False


def process_url():
    try:
        chrome_options = Options()
        chrome_options.add_argument(
            "--disable-blink-features=AutomationControlled")
        chrome_options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
        # service = Service(chrome_driver_path)
        chrome_options.add_experimental_option(
            "prefs", {"profile.managed_default_content_settings.images": 2})

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        url = "https://www.google.com/"
        driver.get(url)
        scope = ["https://www.googleapis.com/auth/spreadsheets"]

        # ptyhon
        credsPython = ServiceAccountCredentials.from_json_keyfile_name(
            f"C:\\Users\\{user_name}\\Desktop\\scrapping_aiscore\\credentials.json", scope)

        client_eternal = gspread.authorize(credsPython)

        # sheet pour les analyses de donnés
        sheet_id = "1lH8i7RCzVkdo4CN71_h13I6yteebocQTqP6ZxHL1nYo"

        sheet_analyse_test = client_eternal.open_by_key(sheet_id)

        processed_filename = f"C:\\Users\\{user_name}\\Desktop\\scrapping_aiscore\\Match_directs\\processed_elements_sofascore.txt"

        processed_elements = load_processed_elements(processed_filename)

        try:
            sheet_result = sheet_analyse_test.worksheet("data")

            array_analyse = sheet_result.get_all_values()

            for i, row in enumerate(array_analyse, start=1):
                # Récupère les trois premières colonnes
                times, league, match = row[:3]

                matchs = match.split('vs')

                home = matchs[0]

                if not check_internet():
                    print("❌ Pas de connexion Internet. Fermeture du scripts.")
                    driver.close()
                    driver.quit()
                    return False

                if home in processed_elements:
                    print(f"✅ {home} déjà traité, passage au suivant...")
                    continue

                google_url = f'https://www.google.com/search?q={f'sofascore {times} {league} {match}'}'

                driver.execute_script("window.open(arguments[0]);", google_url)

                WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

                driver.switch_to.window(driver.window_handles[-1])

                elements = driver.find_elements(
                    By.CSS_SELECTOR, 'span[jscontroller="msmzHf"] a')

                if elements:

                    first_link = elements[0]

                    href = first_link.get_attribute("href")

                    try:
                        driver.get(href)

                        live = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                            By.CSS_SELECTOR, 'div.Box.klGMtt.sc-46a9cb1a-1.cBKodw > span > span > div')).text.strip()

                        print(live)

                        if live in ' Terminé Reporté Annulé ':

                            try:
                                score_home = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                    (By.CSS_SELECTOR,
                                     'div.Box.iCtkKe > span > span:nth-child(1)')
                                )).text.strip()

                                away_home = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                    (By.CSS_SELECTOR,
                                     'div.Box.iCtkKe > span > span:nth-child(3)')
                                )).text.strip()

                                score = f'{score_home}-{away_home}'
                                sheet_result.update_cell(i, 5, score)
                                print(f'update score {home}')

                            except:
                                print('')

                            save_processed_element(
                                home, processed_filename)

                        else:
                            "print la match n'est pas encore términé"
                            driver.close()
                            driver.quit()
                            return

                    except Exception as e:
                        print("pas de donné")
                        save_processed_element(home, processed_filename)

                driver.close()
                driver.switch_to.window(driver.window_handles[0])

        except Exception as e:
            print("❌ error", e)

    except Exception as e:
        print(f"Erreur dans le debuts du code", e)

    finally:
        print(f"script fin")
        driver.close()
        driver.quit()


while True:
    try:
        process_url()
    except Exception:
        print("Erreur")

    time.sleep(60)
