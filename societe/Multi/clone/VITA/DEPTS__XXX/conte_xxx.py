import pandas as pd
import os

def compter_lignes_excel(fichier_excel):
    try:
        # Charger le fichier Excel
        df = pd.read_excel(fichier_excel)
        
        # Compter le nombre total de lignes
        return df.shape[0]
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier {fichier_excel} : {e}")
        return None

# Liste pour stocker les résultats
donnees = []

# Boucle sur les fichiers DEPT_01.xlsx à DEPT_90.xlsx
for i in range(1, 100):
    fichier = f"DEPT_{i:02d}.xlsx"
    if os.path.exists(fichier):
        nombre_lignes = compter_lignes_excel(fichier)
        if nombre_lignes is not None:
            donnees.append([f"Dept {i}", nombre_lignes])
    else:
        print(f"Fichier non trouvé : {fichier}")

# Créer un DataFrame et enregistrer dans un fichier Excel
resultat_df = pd.DataFrame(donnees, columns=["Département", "Nombre de lignes"])
resultat_df.to_excel("resultats_news_xxx.xlsx", index=False)

print("Les résultats ont été enregistrés dans resultats.xlsx")