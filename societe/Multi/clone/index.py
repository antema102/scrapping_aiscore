import pandas as pd
import os


# def load_excel(file_path):
#     """Charge un fichier Excel si disponible, sinon renvoie None."""
#     if os.path.exists(file_path):
#         return pd.read_excel(file_path, dtype={'siren': str})
#     print(f"⚠️ Fichier introuvable : {file_path}")
#     return None

def load_excel(file_path):
    """Charge un fichier Excel si disponible, sinon renvoie None."""
    if os.path.exists(file_path):
        try:
            df = pd.read_excel(file_path, dtype={'siren': str}, engine='openpyxl')
            if 'siren' not in df.columns or 'dénomination' not in df.columns:
                print(f"⚠️ Colonnes manquantes dans {file_path} -> {df.columns.tolist()}")
                return None
            
            return df
        except Exception as e:
            print(f"❌ Erreur de lecture du fichier {file_path} : {e}")
            return None
    print(f"⚠️ Fichier introuvable : {file_path}")
    return None


def extract_last4_digits(df):
    """Ajoute une colonne contenant les 4 derniers chiffres du numéro sirene."""
    df['sirene_last4'] = df['siren'].str[-4:]
    return df


def merge_dataframes(df1, df2):
    """Fusionne les deux DataFrames sur 'sirene_last4' et 'nom societe'."""
    print("🔄 Fusion des fichiers...")
    return df1.merge(df2, on=['sirene_last4', 'dénomination'], how='inner', suffixes=('', '_df2'))


def clean_dataframe(df):
    """Supprime les colonnes dupliquées de df2 et 'sirene_last4' si présente."""
    cols_to_drop = [col for col in df.columns if col.endswith('_df2')]
    if 'sirene_last4' in df.columns:
        cols_to_drop.append('sirene_last4')
    return df.drop(columns=cols_to_drop)


def filter_df2(df2, merged_df):
    """Filtre df2 en supprimant uniquement les lignes qui ont exactement matché avec df1."""
    print("📌 Filtrage des données restantes...")
    merged_pairs = set(zip(merged_df['sirene_last4'], merged_df['dénomination']))
    
    filtered_df2 = df2[
        ~df2.apply(lambda row: (row['sirene_last4'], row['dénomination']) in merged_pairs, axis=1)
    ]
    
    return filtered_df2.drop(columns=['sirene_last4'])

def remove_duplicates(file_path):
    """Supprime les doublons d'un fichier Excel et le sauvegarde."""
    try:
        df = pd.read_excel(file_path, engine='openpyxl')

        # Suppression des doublons en fonction de toutes les colonnes
        df_cleaned = df.drop_duplicates()

        # Sauvegarde du fichier nettoyé
        df_cleaned.to_excel(file_path, index=False)
        print(f"✅ Doublons supprimés pour {file_path}")
    
    except Exception as e:
        print(f"❌ Erreur lors de la suppression des doublons dans {file_path} : {e}")

# 📢 Boucle sur les départements (08 à 90)
for dep in range(1,2):
    try:
        # Formate en deux chiffres (ex : '08', '09', '10')
        dep_str = f"{dep:02d}"

        # Définition des fichiers dynamiquement
        file_df1 = f"news_dep_{dep_str}.xlsx"
        file_df2 = f"dep_{dep_str}_sources.xlsx"
        output_matched = f"news_dep_{dep_str}.xlsx"
        output_filtered = f"news_dep_{dep_str}_xxx.xlsx"

        print(f"\n🚀 Traitement des fichiers pour le département {dep_str}...")

        # Chargement des fichiers
        df1 = load_excel(file_df1)
        df2 = load_excel(file_df2)

        if df1 is None or df2 is None:
            print(f"⏩ Département {dep_str} ignoré (fichiers manquants).")
            continue

        # Extraction des 4 derniers chiffres du numéro sirene
        df1 = extract_last4_digits(df1)
        df2 = extract_last4_digits(df2)

        # Fusion des DataFrames
        merged_df = merge_dataframes(df1, df2)

        # Filtrage de df2
        df2_filtered = filter_df2(df2, merged_df)

        # Nettoyage des colonnes en doublon
        print("🧹 Nettoyage des colonnes en doublon...")

        merged_df = clean_dataframe(merged_df)

        # Sauvegarde des résultats
        print("💾 Sauvegarde des fichiers...")

        merged_df.to_excel(output_matched, index=False)

        df2_filtered.to_excel(output_filtered, index=False)

        remove_duplicates(output_matched)

        remove_duplicates(output_filtered)

        print(f"✅ Suppresion fichiers sources {file_df2} traité ! 🚀")

        os.remove(file_df2)

        print(f"✅ Département {dep_str} traité avec succès ! 🚀")

    except Exception as e:
        print(f'erreur lors de conversion departements {dep_str}', e)

print("\n🎉 Traitement terminé pour tous les départements !")
