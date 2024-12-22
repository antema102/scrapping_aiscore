function verifierEtSupprimerDepassement() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const feuille = spreadsheet.getSheetByName("Match Data"); // Remplacez par le nom de votre feuille
    
    const indexColonne = 2; // Colonne contenant les heures
    const heureActuelle = new Date(); // Heure actuelle du système
    const limiteEnHeures = 2 * 60 * 60 * 1000; // 2 heures en millisecondes
    
    const dernieresLignes = feuille.getLastRow();
    const colonne = feuille.getRange(4, indexColonne, dernieresLignes - 3).getValues(); // Récupérer les données de la colonne
    
    let lignesASupprimer = [];
    let ligneDebutASupprimer = null;
  
    for (let i = 0; i < colonne.length; i++) {
      const valeur = colonne[i][0]; // Récupérer la valeur de la cellule (heure)
  
      if (valeur) {
        // Si une heure est trouvée, calculer la différence avec l'heure actuelle
        const heureColonne = new Date(`${heureActuelle.toDateString()} ${valeur}`);
  
        if (isNaN(heureColonne.getTime())) {
          Logger.log(`Valeur ignorée : ${valeur} (non valide comme heure)`);
          continue;
        }
  
        const difference = heureActuelle - heureColonne;
  
        if (difference > limiteEnHeures) {
          // Si l'heure dépasse 2 heures, marquer pour suppression
          if (ligneDebutASupprimer === null) {
            ligneDebutASupprimer = i + 4; // Ajouter l'index de la ligne à partir de 4
          }
          lignesASupprimer.push(i + 4);
        } else {
          // Si l'heure est valide, réinitialiser la suppression
          ligneDebutASupprimer = null;
        }
      } else {
        // Ligne vide
        if (ligneDebutASupprimer !== null) {
          lignesASupprimer.push(i + 4); // Ajouter la ligne vide pour suppression
        }
      }
    }
  
    // Suppression par lots
    if (lignesASupprimer.length > 0) {
      // Suppression sécurisée pour éviter les erreurs hors limites
      lignesASupprimer.reverse().forEach((ligne) => {
        if (ligne <= dernieresLignes) {
          feuille.deleteRow(ligne);
        }
      });
    }
  }
  