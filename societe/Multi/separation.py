import pandas as pd
import os

# Liste des fichiers d'entrée
input_files = []
user_name = os.getlogin()
for i in range(75,76):  # i va de 45 à 49
    dep_formatted = str(i).zfill(2)
    input_files.append(
        f"C:\\Users\\{user_name}\\Desktop\\scrapping_aiscore\\societe\\Multi\\DEPT\\DEPT_{dep_formatted}.xlsx")

# Fonction pour diviser un fichier Excel en plusieurs fichiers

def split_excel(input_file):
    # Charger les données
    df = pd.read_excel(input_file)

    # Vérifier combien de lignes il y a dans le fichier
    total_rows = len(df)

    # Calculer la taille des différentes parties (environ 10 000 lignes par partie)
    part_size = total_rows // 10  # On divise en 16 parties égales

    # Obtenir le chemin du dossier et le nom du fichier sans extension
    dir_name = os.path.dirname(input_file)
    base_name = os.path.splitext(os.path.basename(input_file))[0]

    # Créer le nouveau dossier DEP_<numéro département>
    dep_folder = os.path.join(dir_name, f"DEPT_{base_name[-2:]}")
    if not os.path.exists(dep_folder):
        os.makedirs(dep_folder)

    # Créer les 16 fichiers
    for i in range(10):
        # Définir l'indice de début et de fin pour chaque partie
        start_idx = i * part_size
        if i == 9:  # Dernière partie, prendre tout ce qui reste
            end_idx = total_rows
        else:
            end_idx = (i + 1) * part_size

        # Extraire la partie du DataFrame
        df_part = df.iloc[start_idx:end_idx]

        # Nom du fichier de sortie (avec le numéro de la partie)
        output_file = os.path.join(dep_folder, f"{base_name}_part_{i + 1}.xlsx")

        # Nom de la feuille dans le fichier Excel
        sheet_name = f"part_{i + 1}"

        # Utilisation d'ExcelWriter pour définir le nom de la feuille
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df_part.to_excel(writer, index=False, sheet_name=sheet_name)

        print(f"Fichier créé : {output_file} avec la feuille {sheet_name}")


# Traiter chaque fichier de la liste
for input_file in input_files:
    if os.path.exists(input_file):
        split_excel(input_file)
    else:
        print(f"Fichier non trouvé : {input_file}")
