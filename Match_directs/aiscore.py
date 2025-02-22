import gspread
from oauth2client.service_account import ServiceAccountCredentials
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import os
import requests
import undetected_chromedriver as uc
import random

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
        chrome_options = uc.ChromeOptions()
        chrome_options.add_argument(
            "--disable-blink-features=AutomationControlled")
        chrome_options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
        chrome_options.add_experimental_option(
            "prefs", {"profile.managed_default_content_settings.images": 2})
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_argument(
            "--disable-blink-features=AutomationControlled")

        service = Service(ChromeDriverManager().install())
        driver = uc.Chrome(service=service, options=chrome_options)
        url = "https://www.google.com/"

        driver.get(url)

        cookie_selector = "#L2AGLb div"

        try:
            bouton = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, cookie_selector)))
            bouton.click()
        except Exception as e:
            print("Une erreur s'est produite lors du clic sur le bouton de consentem")

        scope = ["https://www.googleapis.com/auth/spreadsheets"]

        # ptyhon
        credsPython = ServiceAccountCredentials.from_json_keyfile_name(
            f"C:\\Users\\{user_name}\\Desktop\\scrapping_aiscore\\score102.json", scope)

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

                if not home or home in processed_elements:
                    print(f"✅ {home} déjà traité ou vide, passage au suivant...")
                    continue

                wait_time = random.randint(2, 5)

                google_url = f'https://www.google.com/search?q=site:aiscore.com {league} {match}'

                time.sleep(wait_time)

                driver.execute_script("window.open(arguments[0]);", google_url)

                WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

                driver.switch_to.window(driver.window_handles[-1])

                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'span[jscontroller="msmzHf"] a')))

                elements = driver.find_elements(
                    By.CSS_SELECTOR, 'span[jscontroller="msmzHf"] a')

                # input("'testttt'")

                if elements:

                    first_link = elements[0]

                    href = first_link.get_attribute("href")

                    try:
                        driver.get(href)

                        try:
                            live = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                By.CSS_SELECTOR, '#app > div.detail.view.border-box.back > div.top.color-333.flex-col.flex.align-center > div.flex.w-bar-100.homeBox > div.h-top-center.matchStatus2 > div.flex-1.text-center.scoreBox > div.h-16.m-b-4 > span > span:nth-child(1)')).text.strip()
                            print('match en cours')

                            driver.close()
                            driver.quit()
                            return

                        except:
                            # input("Appuyez sur Entrée pour continuer...")

                            score1 = 'div.font-bold.home-score > span'

                            score2 = 'div.font-bold.away-score > span'

                            try:
                                # score web
                                try:
                                    score_home = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                        (By.CSS_SELECTOR, score1)
                                    )).text.strip()

                                    away_home = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                        (By.CSS_SELECTOR, score2))).text.strip()
                                    score = f'{score_home}-{away_home}'
                                except:
                                    print("")

                                # score mobile
                                try:
                                    score_mob = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                        (By.CSS_SELECTOR, '#app > div.detail.view.border-box > div.w100.contentBox > div.innerMatchInfo > div.matchTop > div.matchInfo > div > div.score')
                                    )).text.strip()
                                    score = score_mob

                                except:
                                    print("")

                                sheet_result.update_cell(i, 5, score)
                                print(f'update score {home} {score}')

                            except:
                                print('')

                            save_processed_element(home, processed_filename)

                    except Exception as e:
                        print("erreur dans la nouvelle onglet", e)
                else:
                    print("Aucun lien trouvé dans la recherche.")
                    driver.close()
                    driver.quit()
                    return
                
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

        except Exception as e:
            print(" errors ", e)

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

    time.sleep(900)
