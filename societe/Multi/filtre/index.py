import pandas as pd
import os

def filtrer_csv_et_convertir(input_csv, output_xlsx):
    try:
        # Colonnes à conserver
        colonnes_conservees = [
            "siren", "dénomination", "adresse2", "code postal", 
            "commune", "code ape établissement", "tranche effectif entreprise"
        ]
        
        # Lire le fichier CSV
        df = pd.read_csv(input_csv, dtype=str, encoding="ISO-8859-1", sep=";", on_bad_lines="skip")
        
        # Garder uniquement les colonnes spécifiées
        df_filtre = df[colonnes_conservees]
        
        df_filtre.drop_duplicates(inplace=True)

        # Sauvegarder en format Excel (XLSX)
        df_filtre.to_excel(output_xlsx, index=False)
        
        print(f"Fichier converti avec succès : {output_xlsx}")
    except Exception as e:
        print(f'error pour le departement {input_file}',e)

# Boucle sur les fichiers dep_01.csv à dep_90.csv
for i in range(6,7):
    input_file = f"dep_{i:02}.csv"
    output_file = f"dep_{i:02}_sources.xlsx"
    
    if os.path.exists(input_file):
        filtrer_csv_et_convertir(input_file, output_file)
    else:
        print(f"Fichier introuvable : {input_file}")
