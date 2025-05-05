# scrapping_aiscore
## Instructions pour utiliser le script Python de récupération des données d'AIScore

Ce script Python permet de récupérer des données depuis le site AIScore. Suivez les étapes ci-dessous pour configurer et exécuter le script correctement.

### Prérequis
1. **Installer Python** : Téléchargez et installez Python depuis [python.org](https://www.python.org/).
2. **Installer Google Chrome** : Assurez-vous que Google Chrome est installé sur votre machine.
3. **Télécharger ChromeDriver** : Téléchargez la version de ChromeDriver correspondant à votre version de Google Chrome depuis [chromedriver.chromium.org](https://chromedriver.chromium.org/).

### Installation des dépendances
1. Clonez ou téléchargez le projet sur votre machine.
2. Installez les dépendances Python nécessaires en exécutant la commande suivante dans le terminal :
    ```bash
    pip install -r requirements.txt
    ```

### Configuration du certificat SSL
Pour utiliser un certificat SSL valide avec Selenium Wire :
1. Générez le certificat en exécutant la commande suivante :
    ```bash
    python -m seleniumwire extractcert
    ```
2. Installez le certificat généré en utilisant la commande suivante (remplacez `/path/to/ca.crt` par le chemin réel du fichier généré) :
    ```bash
    certutil -addstore -f "Root" /path/to/ca.crt
    ```

### Exécution du script
Une fois toutes les étapes ci-dessus terminées, vous pouvez exécuter le script Python pour récupérer les données depuis AIScore.
