import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
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

def flashScore():
    try:
        chrome_driver_path = r"C:\Users\antem\Desktop\scrapping_aiscore\chromedriver\chromedriver.exe" 
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")  # Démarrer en mode maximisé
        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get("https://www.flashscore.com/football/england/premier-league-2014-2015/results/")

        # Attendre que l'élément soit visible, par exemple la liste des résultats
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".event.event--results")))

        # Ajoutez une action de scrapping ici (extraction des données, etc.)

        print("Page chargée et prête à être scrappée !")
        
        # Attendre une action manuelle pour éviter la fermeture immédiate
        input("Appuyez sur Entrée pour fermer le navigateur...")
    
    except Exception as e:
        print("Erreur:", e)
    finally:
        driver.quit()
        print("Fin du scrapping.")

flashScore()
