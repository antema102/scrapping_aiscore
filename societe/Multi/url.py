import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import re

user_name = os.getlogin()

# 🔹 Authentification Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    f'C:/Users/{user_name}/Desktop/scrapping_aiscore/credentials.json', scope
)
client = gspread.authorize(creds)

# 🔹 Ouvrir la Google Sheet principale
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/13u9VHP1rOi2FLke70vSZLtlUM9Onb4SyLoEqMtGAff8"
sheet = client.open_by_url(SPREADSHEET_URL).sheet1  # Feuille active

# 🔹 Fonction pour extraire les vrais URLs des hyperliens
def extract_url(cell):
    if cell and cell.startswith('=HYPERLINK'):
        match = re.search(r'"(https?://[^"]+)"', cell)  # Extraire l'URL de la formule
        return match.group(1) if match else ""
    return cell.strip()  # Retourner directement la valeur si ce n'est pas un HYPERLINK

# 🔹 Récupérer les cellules des colonnes B et G
cells_B = sheet.range(f"B2:B{sheet.row_count}")  # Récupérer toutes les cellules non vides
cells_G = sheet.range(f"G2:G{sheet.row_count}")

# 🔹 Extraire les vrais liens des colonnes B et G
urls_B = [extract_url(cell.value) for cell in cells_B if cell.value]
urls_G = [extract_url(cell.value) for cell in cells_G if cell.value]

# 🔹 Fonction pour compter les lignes d'une Google Sheet donnée
def count_rows(sheet_url):
    try:
        sub_sheet = client.open_by_url(sheet_url).sheet1  # Ouvrir la feuille
        data = sub_sheet.get_all_values()  # Récupérer toutes les données
        return len(data)  # Compter les lignes non vides
    except Exception as e:
        print(f"⚠️ Erreur avec {sheet_url}: {e}")
        return 

# 🔹 Récupérer le nombre de lignes pour chaque URL
counts_B = [count_rows(url) for url in urls_B]
counts_G = [count_rows(url) for url in urls_G]

# 🔹 Mise à jour dans les colonnes D et F
update_data_D = [[count] for count in counts_B]  # Une seule colonne (D)
update_data_F = [[count] for count in counts_G]  # Une seule colonne (F)

# Vérifier et mettre à jour les colonnes
if update_data_D:
    sheet.update(f"D2:D{len(update_data_D) + 1}", update_data_D)
if update_data_F:
    sheet.update(f"F2:F{len(update_data_F) + 1}", update_data_F)

print("✅ Mise à jour terminée !")
