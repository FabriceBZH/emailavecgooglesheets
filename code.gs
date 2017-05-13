function envoiTache() {
  // Sélectionnez la plage de données dans la feuille Liste des tâches
  var classeur = SpreadsheetApp.getActiveSpreadsheet();
  var feuille = classeur.getSheetByName("Liste des tâches en cours");
  var dernierRang = feuille.getLastRow();
  
  var plagededonnees = feuille.getRange(5,1,dernierRang-3,11).getValues();
  Logger.log(plagededonnees.length);
  
  // Créer un horodatage à mentionner lorsque le mail a été envoyé
  var timestamp = new Date();
  
  // Faire une boucle et envoyer un mail si l'option "Oui" est choisie
  for (var i = 0; i < plagededonnees.length; i++) {
    if (plagededonnees[i][8] == "A envoyer") {
      
      // Choisir Propriétaire, Acteur ou Les deux
      switch (plagededonnees[i][9]) {
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
      feuille.getRange(i+5,11,1,1).setValue(timestamp);
      
      // Modifier le statut que le mail a été envoyé
      feuille.getRange(i+5,9,1,1).setValue("Envoyé");
    };
  }
}

// Fonction pour créer et envoyer des emails
function envoiEmailProprietaire(plage) {
  var timeZone =Session.getScriptTimeZone();
  var horodatage = Utilities.formatDate(new Date(), timeZone, "dd-MM-yyyy' à 'HH:mm:ss");  
  var dateEcheance = Utilities.formatDate(plage[7], timeZone, "dd-MM-yyyy");
  var titreRDV = "Echéance sur projet " + plage[0] + " - Tâche " + plage[1] ;
  var lienAgenda = "https://www.google.com/calendar/render?action=TEMPLATE&text=" + titreRDV + 
    "&details=" + plage[2] + "&trp=false&sf=true&output=xml";
  
  
  //Your+Event+Name&dates=20140127T224000Z/20140320T221500Z&details=For+details,+link+here:+http://www.example.com&location=Waldorf+Astoria,+301+Park+Ave+,+New+York,+NY+10022&sf=true&output=xml
  
  
  MailApp.sendEmail({
    to: plage[4],
    subject: "PROJET " + plage[0] + " - Tâche " + plage[1] ,
    htmlBody: 
    "Bonjour" + ",<br><br>" +
    "Ci-dessous, retrouvez la tâche <b>" + plage[1] + "</b><br><br>" +
    "<table  border='1'><tr><td><b>Description de la tâche</b></td>" +
    "<td><b>Propriétaire</b></td>" +
    "<td><b>Acteur</b></td>" +
    "<td><b>Action</b></td>" +
    "<td><b>Date échéance</b></td></tr>" +
    "<tr><td>" + plage[2] + "</td>" +
    "<td>" + plage[4] +"</td>" +
    "<td>" + plage[5] + "</td>" +
    "<td>" + plage[6] + "</td>" +
    "<td>" + dateEcheance + "</td></tr></table>" +
    "<br><b>Commentaires: </b>" + plage[3] + "<br>" +
    "<br>Date de l'envoi: " + horodatage +
    "<br><a href='" + lienAgenda + "'>Ajouter à mon calendrier</a>" + 
    "<br><br>Fourni par un outil de présenté par <a href='https://plus.google.com/u/0/communities/116171205102166939198'>Fabrice Faucheux</a>"
  });
}

// Fonction pour créer et envoyer des emails
function envoiEmailActeur(plage) {
  var timeZone =Session.getScriptTimeZone();
  var horodatage = Utilities.formatDate(new Date(), timeZone, "dd-MM-yyyy' à 'HH:mm:ss"); 
  var dateEcheance = Utilities.formatDate(plage[7], timeZone, "dd-MM-yyyy");
  var titreRDV = "Projet " + plage[0] + " - Action " + plage[6] ;
  var lienAgenda = "https://www.google.com/calendar/render?action=TEMPLATE&text=" + titreRDV + 
    "&details=" + plage[2] + "&trp=false&sf=true&output=xml";
  
  MailApp.sendEmail({
    to: plage[4],
    subject: "PROJET " + plage[0] + "- Tâche " + plage[1] ,
    htmlBody: 
    "Bonjour" + ",<br><br>" +
    "Ci-dessous, retrouvez la tâche <b>" + plage[1] + "</b> à réaliser.<br><br>" +
    "<table  border='1'><tr><td><b>Description de la tâche</b></td>" +
    "<td><b>Propriétaire</b></td>" +
    "<td><b>Acteur</b></td>" +
    "<td><b>Action</b></td>" +
    "<td><b>Date échéance</b></td></tr>" +
    "<tr><td>" + plage[2] + "</td>" +
    "<td>" + plage[4] +"</td>" +
    "<td>" + plage[5] + "</td>" +
    "<td>" + plage[6] + "</td>" +
    "<td>" + dateEcheance + "</td></tr></table>" +
    "<br><b>Commentaires: </b>" + plage[3] + "<br>" +
    "<br>Date de l'envoi: " + horodatage +
    "<br><a href='" + lienAgenda + "'>Ajouter à mon calendrier</a>" +
    "<br><br>Fourni par un outil de présenté par <a href='https://plus.google.com/u/0/communities/116171205102166939198'>Fabrice Faucheux</a>"
  });
}
