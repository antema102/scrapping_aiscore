import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime


def trie_1_x_2(spreadsheet):
    sheet = spreadsheet.sheet1 
    data = sheet.get_all_values()
    
    home_array = []
    away_array = []
    sorted_home = []
    sorted_away = []

    # Créer ou réinitialiser les feuilles pour les résultats
    try:
        home_sheet = spreadsheet.worksheet("FD 1X2")
        home_sheet.clear()  # Vider les données existantes
    except gspread.exceptions.WorksheetNotFound:
        home_sheet = spreadsheet.add_worksheet(title="FD 1X2", rows="100", cols="3")

    try:
        away_sheet = spreadsheet.worksheet("FE 1X2")
        away_sheet.clear()  # Vider les données existantes
    except gspread.exceptions.WorksheetNotFound:
        away_sheet = spreadsheet.add_worksheet(title="FE 1X2", rows="100", cols="3")

    # Ajouter les en-têtes
    home_sheet.append_row(["Score", "1XBET ODDS", "Date"])
    away_sheet.append_row(["Score", "1XBET ODDS", "Date"])

    for row in data[1:]:
        date = row[4]
        score = row[0]
        cote_1xBET = row[1].split('/')
        premier = int(cote_1xBET[0])
        dernier = int(cote_1xBET[-1])
        
        # Conversion de la date
        date_obj = datetime.strptime(date, "%d/%m/%Y")  # Adapte le format si nécessaire
        
        if premier < dernier:
            home_array.append((score, row[1],date_obj,premier))  
        else:
            away_array.append((score, row[1],date_obj,dernier))

    home_array.sort(key=lambda x: (x[3], -x[2].timestamp()))
    away_array.sort(key=lambda x: (x[3], -x[2].timestamp()))

    # Fonction pour ajouter les séparateurs dans les tableaux triés
    def ajouter_avec_separateur(array, sorted_array):
        previous_cote = None
        for i, item in enumerate(array):
            cote = item[3]  # Récupère la cote
            sorted_array.append([item[0], item[1], item[2].strftime("%d/%m/%Y")])
            
            # Ajoute une ligne de séparation lorsque la cote change
            if i == len(array) - 1 or cote != array[i + 1][3]:
                sorted_array.append(["-----", "--- ----", "---"])

    # Appliquer la logique de séparation
    ajouter_avec_separateur(home_array, sorted_home)
    ajouter_avec_separateur(away_array, sorted_away)

    home_sheet.append_rows(sorted_home)
    away_sheet.append_rows(sorted_away)

def process_odds(spreadsheet, _name, column_index, sheet_suffix=""):
    # Tableaux pour stocker les résultats
    home_rows_to_add = []
    away_rows_to_add = []

    sheet = spreadsheet.worksheet(_name)

    # Extraire les données de la feuille
    data = sheet.get_all_values()[1:]

    # Obtenir les valeurs uniques dans la colonne ciblée (ignorant les vides)
    unique_odds = sorted(set(str(row[column_index]) for row in data if row[column_index] and row[column_index] != ""))

    # Créer ou réinitialiser les feuilles pour les résultats
    home_title = f"FD {sheet_suffix} {_name}".strip()
    away_title = f"FE {sheet_suffix} {_name}".strip()

    try:
        home_sheet = spreadsheet.worksheet(home_title)
        home_sheet.clear()
    except gspread.exceptions.WorksheetNotFound:
        home_sheet = spreadsheet.add_worksheet(title=home_title, rows="100", cols="4")

    try:
        away_sheet = spreadsheet.worksheet(away_title)
        away_sheet.clear()
    except gspread.exceptions.WorksheetNotFound:
        away_sheet = spreadsheet.add_worksheet(title=away_title, rows="100", cols="4")

    # Ajouter les en-têtes
    if "OVER/UNDER" in sheet_suffix:
        headers = ["Score", "Total", "1XBET ODDS", "OVER/UNDER", "Date"]
    else:
        headers = ["Score", "1XBET ODDS", f"{_name}", "Date"]

    home_sheet.append_row(headers)
    away_sheet.append_row(headers)

    for odd in unique_odds:
        matching_rows = [row for row in data if str(row[column_index]) == odd]

        # Trier les lignes par "1XBET ODDS"
        sorted_rows = sorted(
            matching_rows,
            key=lambda row: [int(i) for i in (row[column_index - 1].split('/') if row[column_index - 1] else []) if i.isdigit()]
        )

        # Ajouter une séparation entre groupes de cotes
        home_rows_to_add.append(["-----------", "--------", "-------", "-------"])
        away_rows_to_add.append(["-----------", "--------", "-------", "-------"])

        # Trier les lignes pour Favoris domicile et extérieur
        home_rows = []
        away_rows = []

        for row in sorted_rows:
            if row[column_index - 1]:  # Vérifier que "1XBET ODDS" n'est pas vide
                odds = [int(i) for i in row[column_index - 1].split('/') if i.isdigit()]
                if len(odds) >= 3:
                    if odds[0] == min(odds):
                        home_rows.append(row)
                    elif odds[2] == min(odds):
                        away_rows.append(row)

        # Trier les favoris extérieurs par la dernière cote
        away_rows_sorted = sorted(
            away_rows,
            key=lambda row: int(row[column_index - 1].split('/')[-1]) if row[column_index - 1] and row[column_index - 1].split('/')[-1].isdigit() else float('inf')
        )

        home_rows_to_add.extend(home_rows)
        away_rows_to_add.extend(away_rows_sorted)

    # Ajouter les lignes dans les feuilles
    home_sheet.append_rows(home_rows_to_add)
    away_sheet.append_rows(away_rows_to_add)


def trie_flashscore():
    # Authentification avec l'API Google Sheets
    scope = ["https://spreadsheets.google.com/feeds"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\antem\Desktop\scrapping_aiscore\credentials.json", scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key("1-agugik6J7Bo6XU2GqioAC68PWGDFkDW97TacDdA8SY") 
    trie_1_x_2(spreadsheet)
    process_odds(spreadsheet,'Cotes_Both_teams_to_score',column_index=2)
    process_odds(spreadsheet,'Cotes_ODD/EVENT',column_index=2)

    # for i in range(10):
    #     process_odds(spreadsheet, f"{i}.5", column_index=3, sheet_suffix="OVER/UNDER")


trie_flashscore()

