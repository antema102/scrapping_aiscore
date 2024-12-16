import os
import shutil

def supprimer_dossier(chemin):
    if os.path.isfile(chemin):
        try:
            os.remove(chemin)  # Supprime le dossier et son contenu
            print(f"Dossier supprimé : {chemin}")
        except Exception as e:
            print(f"Erreur lors de la suppression : {e}")
    else:
        print(f"Le dossier n'existe pas : {chemin}")

if __name__ == "__main__":
    chemin_dossier = r"C:\Users\Administrator\Desktop\scrapping_aiscore\Match_directs\processed_elements.txt"  # Remplacez par le chemin du dossier à supprimer
    supprimer_dossier(chemin_dossier)
