import gspread
from oauth2client.service_account import ServiceAccountCredentials

def filter_odds_and_sort():
    # Authentification avec l'API Google Sheets
    scope = ["https://spreadsheets.google.com/feeds"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("C:/Users/etech/Desktop/scrapping_aiscore/credentials.json", scope)
    client = gspread.authorize(creds)

    # Ouvrir le fichier Google Sheets via son ID
    spreadsheet = client.open_by_key("1uQSHqJShVfICDFJLfCCNzm7xFrTZMtAe_3n0dFQGi9E")  # Remplacez par l'ID de votre feuille
    sheet = spreadsheet.sheet1  # Par défaut, la première feuille active

    # Extraire les données de la feuille
    data = sheet.get_all_values()[1:]  # Extraire toutes les valeurs sauf la première ligne (en-têtes)

    # Obtenir les valeurs uniques dans la colonne "1XBET O/U 2.5", en ignorant les vides
    unique_odds = sorted(set(str(row[2]) for row in data if row[2] and row[2] != ""))

    # Créer ou réinitialiser les feuilles pour les résultats
    try:
        home_sheet = spreadsheet.worksheet("Favoris domicile")
        home_sheet.clear()  # Vider les données existantes
    except gspread.exceptions.WorksheetNotFound:
        home_sheet = spreadsheet.add_worksheet(title="Favoris domicile", rows="100", cols="4")

    try:
        away_sheet = spreadsheet.worksheet("Favoris extérieur")
        away_sheet.clear()  # Vider les données existantes
    except gspread.exceptions.WorksheetNotFound:
        away_sheet = spreadsheet.add_worksheet(title="Favoris extérieur", rows="100", cols="4")

    # Ajouter les en-têtes
    home_sheet.append_row(["Score", "1XBET ODDS", "1XBET O/U 2.5", "Date"])
    away_sheet.append_row(["Score", "1XBET ODDS", "1XBET O/U 2.5", "Date"])

    # Tableaux pour stocker les résultats
    home_rows_to_add = []
    away_rows_to_add = []

    # Parcourir chaque valeur unique dans "1XBET O/U 2.5" triée
    for odd in unique_odds:
        # Filtrer les lignes correspondantes
        matching_rows = [row for row in data if str(row[2]) == odd]

        # Trier les lignes en fonction des valeurs de "1XBET ODDS" (du plus petit au plus grand)
        sorted_rows = sorted(
            matching_rows,
            key=lambda row: [int(i) for i in (row[1].split('/') if row[1] else []) if i.isdigit()]  # Vérification de None
        )

        # Ajouter une ligne vide avant chaque groupe de valeurs
        home_rows_to_add.append(["-----------", "--------", "-------", "-------"])
        away_rows_to_add.append(["-----------", "--------", "-------", "-------"])

        # Trier les lignes pour Favoris extérieur et domicile
        home_rows = []
        away_rows = []

        for row in sorted_rows:
            if row[1]:  # Vérifier que "1XBET ODDS" n'est pas vide
                odds = [int(i) for i in row[1].split('/') if i.isdigit()]
                if len(odds) >= 3:  # S'assurer qu'il y a au moins 3 cotes
                    if odds[0] == min(odds):  # Si la première cote est la plus petite
                        home_rows.append(row)
                    elif odds[2] == min(odds):  # Si la troisième cote est la plus petite
                        away_rows.append(row)

        # Trier les lignes des favoris extérieur par le dernier chiffre de 1XBET ODDS
        away_rows_sorted = sorted(
            away_rows,
            key=lambda row: int(row[1].split('/')[-1]) if row[1] and row[1].split('/')[-1].isdigit() else float('inf')
        )

        # Ajouter les résultats filtrés aux tableaux
        home_rows_to_add.extend(home_rows)
        away_rows_to_add.extend(away_rows_sorted)

    # Ajouter les lignes filtrées dans les feuilles respectives en une seule fois
    home_sheet.append_rows(home_rows_to_add)
    away_sheet.append_rows(away_rows_to_add)

    print("Filtrage, tri et ajout dans les onglets terminé!")

# Exécution de la fonction
filter_odds_and_sort()
