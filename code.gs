function envoiTache() {
  // Sélectionnez la plage de données dans la feuille Liste des tâches
  var classeur = SpreadsheetApp.getActiveSpreadsheet();
  var feuille = classeur.getSheetByName("Liste des tâches");
  var dernierRang = feuille.getLastRow();
  
  var plagededonnees = feuille.getRange(5,1,dernierRang-3,9).getValues();
  Logger.log(plagededonnees.length);
  
  // Créer un horodatage à mentionner lorsque le mail a été envoyé
  var timestamp = new Date();
  
  // Faire une boucle et envoyer un mail si l'option "Oui" est choisie
  for (var i = 0; i < plagededonnees.length; i++) {
    if (plagededonnees[i][6] == "A envoyer") {
      
      // Choisir Propriétaire, Acteur ou Les deux
      switch (plagededonnees[i][7]) {
        case "Propriétaire":
          // Envoyer un courriel au propriétaire en appelant la fonction envoiEmail
          envoiEmailProprietaire(plagededonnees[i]);
          break;
          
        case "Acteur":
          // Envoyer un courriel à l'acteur en appelant la fonction envoiEmail
          envoiEmailActeur(plagededonnees[i]);
          break;
          
        case "Les deux":
          // envoi aux deux
          envoiEmailProprietaire(plagededonnees[i]);
          envoiEmailActeur(plagededonnees[i]);
          break;
      }
      
      // Ajouter un Non pour le statut 
      feuille.getRange(i+5,9,1,1).setValue(timestamp);
      
      // Modifier le statut que le mail a été envoyé
      feuille.getRange(i+5,7,1,1).setValue("Envoyé");
    };
  }
}

// Fonction pour créer et envoyer des emails
function envoiEmailProprietaire(plage) {
  var timeZone =Session.getScriptTimeZone();
  var horodatage = Utilities.formatDate(new Date(), timeZone, "dd-MM-yyyy' à 'HH:mm:ss");
  
  var dateEcheance = Utilities.formatDate(plage[5], timeZone, "dd-MM-yyyy");
  
  MailApp.sendEmail({
    to: plage[3],
    subject: "Information sur la tâche " + plage[0] ,
    htmlBody: 
    "Bonjour" + ",<br><br>" +
    "Ci-dessous, retrouvez la tâche <b>" + plage[0] + "</b><br><br>" +
    "<table  border='1'><tr><td><b>Description de la tâche</b></td>" +
    "<td><b>Propriétaire</b></td>" +
    "<td><b>Acteur</b></td>" +
    "<td><b>Date échéance</b></td></tr>" +
    "<tr><td>" + plage[1] + "</td>" +
    "<td>" + plage[3] +"</td>" +
    "<td>" + plage[4] + "</td>" +
    "<td>" + dateEcheance + "</td></tr></table>" +
    "<br><b>Commentaires: </b>" + plage[2] + "<br>" +
    "<br>Date de l'envoi: " + horodatage +
    "<br><br>Fourni par un outil de présenté par <a href='https://plus.google.com/u/0/communities/116171205102166939198'>Fabrice Faucheux</a>"
  });
}

// Fonction pour créer et envoyer des emails
function envoiEmailActeur(plage) {
  var timeZone =Session.getScriptTimeZone();
  var horodatage = Utilities.formatDate(new Date(), timeZone, "dd-MM-yyyy' à 'HH:mm:ss");
    
  var dateEcheance = Utilities.formatDate(plage[5], timeZone, "dd-MM-yyyy");
  
  MailApp.sendEmail({
    to: plage[4],
    subject: "Information tâche " + plage[0] + " à réaliser !",
    htmlBody: 
    "Bonjour" + ",<br><br>" +
    "Ci-dessous, retrouvez la tâche <b>" + plage[0] + "</b> à réaliser.<br><br>" +
    "<table  border='1'><tr><td><b>Description de la tâche</b></td>" +
    "<td><b>Propriétaire</b></td>" +
    "<td><b>Acteur</b></td>" +
    "<td><b>Date échéance</b></td></tr>" +
    "<tr><td>" + plage[1] + "</td>" +
    "<td>" + plage[3] +"</td>" +
    "<td>" + plage[4] + "</td>" +
    "<td>" + dateEcheance + "</td></tr></table>" +
    "<br><b>Commentaires: </b>" + plage[2] + "<br>" +
    "<br>Date de l'envoi: " + horodatage +
    "<br><br>Fourni par un outil de présenté par <a href='https://plus.google.com/u/0/communities/116171205102166939198'>Fabrice Faucheux</a>"
  });
}
