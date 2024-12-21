function calculerDifferenceLignes() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const feuille = spreadsheet.getSheetByName("Match Data"); // Remplacez par le nom de votre feuille
    
    // Index de la colonne contenant les heures
    const indexColonne = 2; // Par exemple, la deuxième colonne
    
    // Déterminer la hauteur des données
    const dernieresLignes = feuille.getLastRow();
    
    // Récupérer les données de la colonne (à partir de la ligne 4)
    const colonne = feuille.getRange(4, indexColonne, dernieresLignes - 3).getValues(); // Ligne 4 à la dernière
    
    // Initialiser une liste pour stocker les lignes contenant des horaires
    const differences = [];
    let dernierIndex = null; // Conserver l'index de la dernière cellule avec une valeur
    
    for (let i = 0; i < colonne.length; i++) {
      const valeur = colonne[i][0]; // Accéder à la valeur dans le tableau 2D
      
      if (valeur) { // Vérifier si la cellule n'est pas vide
        if (dernierIndex !== null) {
          const difference = i - dernierIndex; // Calculer la différence de lignes
          differences.push({ heure: valeur, difference });
        }
        dernierIndex = i; // Mettre à jour l'index de la dernière cellule
      }
    }
    
    // Afficher les résultats
    Logger.log("Différences entre les heures :");
    differences.forEach(d => Logger.log(`Heure : ${d.heure}, Différence : ${d.difference}`));
  }
  