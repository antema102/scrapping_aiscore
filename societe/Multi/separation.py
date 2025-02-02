import pandas as pd
from openpyxl import load_workbook
import os

user_name = os.getlogin()

# Liste des fichiers Excel à traiter
input_files = []  # Initialiser la liste des fichiers

# Générer les noms de fichiers pour les départements 7 à 12
for i in range(9,10):  # i va de 7 à 12
    dep_formatted = str(i).zfill(2)
    input_files.append(f"C:\\Users\\{user_name}\\Desktop\\scrapping_aiscore\\societe\\Multi\\DEPT_{dep_formatted}.xlsx")

# Préfixe pour les fichiers de sortie
output_file_prefix = 'partie_'

# Fonction pour diviser un fichier Excel en plusieurs parties et ajouter des onglets
def split_excel(input_file):
    # Charger les données
    df = pd.read_excel(input_file)

    # Vérifier combien de lignes il y a dans le fichier
    total_rows = len(df)

    # Calculer la taille des différentes parties (environ 10 000 lignes par partie)
    part_size = total_rows // 16  # On divise en 6 parties égales

    # Ouvrir le fichier Excel existant
    with pd.ExcelWriter(input_file, engine='openpyxl', mode='a') as writer:
        # Créer les 6 parties et les ajouter dans des onglets
        for i in range(16):
            # Définir l'indice de début et de fin pour chaque partie
            start_idx = i * part_size
            if i == 15:  # Dernière partie, prendre tout ce qui reste
                end_idx = total_rows
            else:
                end_idx = (i + 1) * part_size

            # Extraire la partie du DataFrame
            df_part = df.iloc[start_idx:end_idx]

            # Ajouter la partie comme un nouvel onglet dans le fichier Excel
            sheet_name = f'part_{i + 1}'  # Nom de l'onglet
            df_part.to_excel(writer, index=False, sheet_name=sheet_name)

            print(f"{input_file} - Partie {i + 1} ajoutée sous l'onglet {sheet_name}")

# Traiter chaque fichier de la liste
for input_file in input_files:
    split_excel(input_file)

print("Séparation terminée pour tous les fichiers avec ajout d'onglets !")
