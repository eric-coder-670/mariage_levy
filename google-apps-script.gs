/**
 * Faire-part Levy & Adrienne — réception des réponses RSVP dans Google Sheets.
 * Déploiement : Déployer > Nouveau déploiement > Application Web
 *   Exécuter en tant que : Moi  |  Accès : Tout le monde
 * Après toute modification du code, créez une NOUVELLE VERSION du déploiement.
 */
const SHEET_NAME = 'Réponses RSVP';
const ENTETES = ['Date', 'Nom', 'Contact', 'Présence', 'Événements', 'Nombre de personnes', 'Message'];
const LIBELLES = { mairie: 'Traditionnel et civil', eglise: 'Mariage religieux', soiree: 'Réception' };

function doPost(e) {
  const verrou = LockService.getScriptLock();
  try {
    verrou.waitLock(15000);
    const donnees = (e && e.parameter) || {};
    const nom = nettoyer(donnees.nom, 120);
    if (!nom) return reponse({ success: false, error: 'nom manquant' });

    const absent = donnees.absent === 'oui';
    const brut = e.parameters && e.parameters.evenements ? e.parameters.evenements : [];
    const evenements = brut.map(function (v) { return LIBELLES[v] || nettoyer(v, 40); }).join(', ');
    const nombre = absent ? '' : Math.min(Math.max(parseInt(donnees.nombre, 10) || 1, 1), 10);

    const feuille = obtenirFeuille();
    feuille.appendRow([
      new Date(),
      nom,
      nettoyer(donnees.contact, 120),
      absent ? 'Absent' : 'Présent',
      evenements,
      nombre,
      nettoyer(donnees.message, 1000)
    ]);
    return reponse({ success: true });
  } catch (err) {
    return reponse({ success: false, error: String(err) });
  } finally {
    try { verrou.releaseLock(); } catch (x) { /* verrou non obtenu */ }
  }
}

/** Test rapide : ouvrir l'URL du déploiement dans un navigateur doit afficher { "success": true }. */
function doGet() {
  return reponse({ success: true, service: 'rsvp' });
}

function obtenirFeuille() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  let feuille = classeur.getSheetByName(SHEET_NAME);
  if (!feuille) feuille = classeur.insertSheet(SHEET_NAME);
  if (feuille.getLastRow() === 0) {
    feuille.appendRow(ENTETES);
    feuille.getRange(1, 1, 1, ENTETES.length).setFontWeight('bold').setBackground('#DCEEFA');
    feuille.setFrozenRows(1);
  }
  return feuille;
}

/** Limite la longueur et neutralise les formules (=, +, -, @) pour éviter l'injection dans Sheets. */
function nettoyer(valeur, max) {
  let t = String(valeur == null ? '' : valeur).trim().slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

function reponse(objet) {
  return ContentService.createTextOutput(JSON.stringify(objet)).setMimeType(ContentService.MimeType.JSON);
}
