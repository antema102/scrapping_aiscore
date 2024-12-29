import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
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

def attendre_element(driver, selector, condition=EC.element_to_be_clickable, timeout=10):
    return WebDriverWait(driver, timeout).until(condition((By.CSS_SELECTOR, selector)))

def recuperer_cotes(driver,name,ui_table="div.ui-table__row"):
    ui_table_elements = driver.find_elements(By.CSS_SELECTOR, ui_table)
    result = []  # Liste pour stocker les résultats

    for item in ui_table_elements:
        try:
            image = item.find_element(By.CSS_SELECTOR, "img.prematchLogo")
            image_text = image.get_attribute("title")
            cotes = []

            if image_text == "1xBet":
                odds_links = item.find_elements(By.CSS_SELECTOR, "a.oddsCell__odd")

                for link in odds_links:
                    cote = link.find_element(By.CSS_SELECTOR, "span").text.strip().replace('.', '')
                    cotes.append(cote)

                if name == "1X2":
                    n = 3
                else:
                    n = 2

                # Vérifier si assez de cotes sont disponibles

                if len(cotes) >= n:
                    if n == 3:
                        cote_1, cote_X, cote_2 = cotes[:3]
                        result=(f"{cote_1}/{cote_X}/{cote_2}")
                    elif name=="Both_teams_to_score" and n == 2: 
                        cote_1, cote_2 = cotes[:2]
                        result=(f"{cote_1}/{cote_2}")
                    else:
                        cote_1, cote_2 = cotes[:2]
                        result=(f"{cote_1}/{cote_2}")
                else:
                    print(f"{name}: cotes insuffisantes - {cotes}")

        except Exception as e:
            print("error", e)

    return result 

def recuper_cotes_over_under(driver,score,date,cote_over_under,cotes1x2,ui_table="div.ui-table__row"):
    result=[]
    ui_table_elements = driver.find_elements(By.CSS_SELECTOR, ui_table)
    for item in ui_table_elements:
        try:
            image = item.find_element(By.CSS_SELECTOR, "img.prematchLogo")
            odds_span = item.find_element(By.CSS_SELECTOR, "span.oddsCell__noOddsCell").text.strip()
            image_text = image.get_attribute("title")

            if image_text == '1xBet' and float(odds_span) % 1 == 0.5:
                cotes = []
                odds_links = item.find_elements(By.CSS_SELECTOR, "a.oddsCell__odd")

                for link in odds_links[:2]:
                    cote = link.find_element(By.CSS_SELECTOR, "span").text.strip().replace('.', '')
                    cotes.append(cote)

                # Ajouter à result uniquement si 2 cotes sont présentes
                if len(cotes) == 2:
                    result.append([score,cotes1x2,odds_span, f"{cotes[0]}/{cotes[1]}", date])

        except Exception as e:
            print("error cotes_over_under", e)

    cote_over_under.append_rows(result)

    return result

def flashScore(url):
    try:
        # Chemin vers le driver Chrome
        chrome_driver_path = r"C:\Users\etech\Desktop\scrapping_aiscore\chromedriver\chromedriver.exe"
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")  # Démarrer en mode maximisé

        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)

        # URL cible
        driver.get(url)

        # Attente jusqu'à ce que les éléments soient visibles
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.sportName.soccer")))
        # time.sleep(5)  # Attendre un peu pour permettre le chargement

        # Charger les éléments déjà traités
        date_text = url.split('/')[-3]
        years=int(date_text.split('-')[-1])

        processed_filename = f"processed_elements_{date_text}.txt"
        processed_elements = load_processed_elements(processed_filename)

        #Google sheets
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\etech\Desktop\scrapping_aiscore\credentials.json', scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key("1-agugik6J7Bo6XU2GqioAC68PWGDFkDW97TacDdA8SY")  
    
        #Créer ou réinitialiser les feuilles pour les résultats
        try:
            cotes_1x2_both_matchs = spreadsheet.worksheet("Cotes_Total")
        except gspread.exceptions.WorksheetNotFound:
            cotes_1x2_both_matchs = spreadsheet.add_worksheet(title="Cotes_Total", rows="100", cols="10")
            cotes_1x2_both_matchs.append_row(["Score","1XBET ODDS","Both_teams_to_score","ODD/EVENT","Date"])

        #Créer ou réinitialiser les feuilles pour les résultats
        try:
            cotes_Both_teams_to_score = spreadsheet.worksheet("Cotes_Both_teams_to_score")
        except gspread.exceptions.WorksheetNotFound:
            cotes_Both_teams_to_score = spreadsheet.add_worksheet(title="Cotes_Both_teams_to_score", rows="100", cols="10")
            cotes_Both_teams_to_score.append_row(["Score","1XBET ODDS","Both_teams_to_score","Date"])

        #Créer ou réinitialiser les feuilles pour les résultats
        try:
            cotes_ODD_event = spreadsheet.worksheet("Cotes_ODD/EVENT")
        except gspread.exceptions.WorksheetNotFound:
            cotes_ODD_event = spreadsheet.add_worksheet(title="Cotes_ODD/EVENT", rows="100", cols="10")
            cotes_ODD_event.append_row(["Score","1XBET ODDS","ODD/EVENT","Date"])

        try:
            cote_over_under = spreadsheet.worksheet("cote_over_under")
        except gspread.exceptions.WorksheetNotFound:
            cote_over_under = spreadsheet.add_worksheet(title="cote_over_under", rows="100", cols="10")
            cote_over_under.append_row(["Score","1XBET ODDS","Total","OVER/UNDER","Date"])
  
        attendre_element(driver, "#onetrust-accept-btn-handler").click()

        while True:
            try:
                match_elements = driver.find_elements(By.CSS_SELECTOR, '.event__match')
                initial_count = len(match_elements)

                if not match_elements:
                    print("Pas de matchs trouvés.")
                    break

                print(f"{initial_count} matchs trouvés, traitement en cours...")

                new_data_found = False

                for element in match_elements:
                    try:
                        match_id = element.get_attribute("id")

                        if match_id in processed_elements:
                            print(f"[{match_id}] Déjà traité, passage au suivant.")
                            continue

                        link_element = element.find_element(By.CSS_SELECTOR, 'a.eventRowLink')
                        href = link_element.get_attribute("href")

                        print(f"[{match_id}] Nouveau match trouvé : {href}")

                        save_processed_element(match_id, processed_filename)
                        processed_elements.add(match_id)
                        new_data_found = True

                        tableau_reponse = []
                        tableaux_over_under=[]
                        try:
                            driver.execute_script("window.open(arguments[0]);", href)

                            #attend que le nouvelle onglet est ouvert
                            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

                            #Basculer vers la nouvelle onglet
                            driver.switch_to.window(driver.window_handles[-1])

                            # Récupération de la chaîne contenant la date et l'heure
                            date_str  = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR,"#detail > div.duelParticipant > div.duelParticipant__startTime > div"))).text.strip()

                            # Conversion de la chaîne en objet datetime
                            date_obj = datetime.strptime(date_str, "%d.%m.%Y %H:%M")

                            # Formatage de l'objet datetime en chaîne au format souhaité
                            formatted_date = date_obj.strftime("%d/%m/%Y")
                            score_home = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR,"#detail > div.duelParticipant > div.duelParticipant__score > div > div.detailScore__wrapper > span:nth-child(1)"))).text.strip()
                            score_away = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CSS_SELECTOR,"#detail > div.duelParticipant > div.duelParticipant__score > div > div.detailScore__wrapper > span:nth-child(3)"))).text.strip()

                            score=f"{score_home}-{score_away}"

                            if years < 2022:
                                odds="#detail > div.detailOver > div > a:nth-child(2)"
                            else:
                                odds="#detail > div.detailOver > div > a:nth-child(3)"

                            attendre_element(driver, odds).click()
                            
                            attendre_element(driver, "div.ui-table__body", EC.presence_of_element_located, 20)

                            #Récuperation de la cotes 1X2
                            cotes_1x2 = recuperer_cotes(driver,"1X2")

                            attendre_element(driver, "div.filterOver.filterOver--indent > div > a:nth-child(2)").click()

                            attendre_element(driver, "div.ui-table__body", EC.presence_of_element_located, 20)

                            #Récuperation de la cotes  over under
                            recuper_cotes_over_under(driver,score,formatted_date,cote_over_under,cotes_1x2)

                            attendre_element(driver, "div.filterOver.filterOver--indent > div > a:nth-child(4)").click()

                            attendre_element(driver, "div.ui-table__body", EC.presence_of_element_located, 20)

                            #Recuperationde la cotes both teams to score
                            cote_both=recuperer_cotes(driver,"Both_teams_to_score")
                            cotes_Both_teams_to_score.append_row([score,cotes_1x2,cote_both,formatted_date])

                            attendre_element(driver, "div.filterOver.filterOver--indent > div > a:nth-child(10)").click()
                            attendre_element(driver, "div.ui-table__body", EC.presence_of_element_located, 20)

                            #Récuperation de la cotes odd events
                            cotes_odd_event = recuperer_cotes(driver,"cotes odd event")
                            if cotes_odd_event:
                                cotes_1x2_both_matchs.append_row([score,cotes_1x2,cote_both,cotes_odd_event,formatted_date])
                                cotes_ODD_event.append_row([score,cotes_1x2,cotes_odd_event,formatted_date])
                            else:
                                cotes_1x2_both_matchs.append_row([score,cotes_1x2,cote_both,"",formatted_date])

                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])    
                                
                        except Exception as e:
                            print(f"Erreur lors de l'ouverture de nouvelle onglets de l'élément {match_id} url {href} : {e}")
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])   
    
                    except Exception as e:
                        print(f"Erreur lors du traitement de l'élément {match_id} : {e}")

                # Si aucun nouveau match n'a été trouvé, arrêt
                if not new_data_found:
                    print("Aucun nouvel élément, arrêt du scrapping.")
                    break

                # Masquer l'iframe publicitaire
                try:
                    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
                    for iframe in iframes:
                        if "lsadvert" in iframe.get_attribute("id"):
                            driver.execute_script("arguments[0].style.display = 'none';", iframe)
                            print("Iframe publicitaire masqué.")
                            time.sleep(1)
                except Exception as e:
                    print(f"Erreur lors du masquage de l'iframe publicitaire : {e}")

                # Clic sur "Show more matches"
                try:
                    show_more_button = driver.find_element(By.CSS_SELECTOR, 'a.event__more.event__more--static')
                    driver.execute_script("arguments[0].scrollIntoView(true);", show_more_button)
                    time.sleep(2)  # Attendre un peu pour s'assurer que le défilement est terminé
                    driver.execute_script("arguments[0].click();", show_more_button)
                    print("Bouton 'Show more matches' cliqué.")
                    
                    # Attendre que de nouveaux matchs apparaissent
                    WebDriverWait(driver, 5).until(lambda d: len(d.find_elements(By.CSS_SELECTOR, '.event__match')) > initial_count)
                    time.sleep(2)

                except Exception as e:
                    print("Plus de bouton 'Show more matches' ou aucun nouvel élément chargé :", e)
                    break

            except Exception as e:
                print("Erreur dans la boucle principale :", e)

    except Exception as e:
        print("Erreur générale :", e)

    finally:
        driver.quit()
        print("Fin du scrapping.")

# Traiter chaque URL et enregistrer les données
urls = ["https://www.flashscore.com/football/england/premier-league-2014-2015/results/",
        "https://www.flashscore.com/football/england/premier-league-2015-2016/results/",
        "https://www.flashscore.com/football/england/premier-league-2016-2017/results/",
        "https://www.flashscore.com/football/england/premier-league-2017-2018/results/",
        "https://www.flashscore.com/football/england/premier-league-2018-2019/results/",
        "https://www.flashscore.com/football/england/premier-league-2019-2020/results/",
        "https://www.flashscore.com/football/england/premier-league-2020-2021/results/",
        "https://www.flashscore.com/football/england/premier-league-2021-2022/results/",
        "https://www.flashscore.com/football/england/premier-league-2022-2023/results/",
        "https://www.flashscore.com/football/england/premier-league-2023-2024/results/",
        ]

for url in urls:
    flashScore(url)