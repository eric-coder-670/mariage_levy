    const SHEET_NAME = 'Réponses RSVP';

    function doPost(e) {
    const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
        || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);

    if (feuille.getLastRow() === 0) {
        feuille.appendRow([
        'Date',
        'Nom',
        'Contact',
        'Présence',
        'Événements',
        'Nombre de personnes',
        'Message'
        ]);
    }

    const donnees = e.parameter || {};
    const evenements = e.parameters && e.parameters.evenements
        ? e.parameters.evenements.join(', ')
        : String(donnees.evenements || '');

    feuille.appendRow([
        new Date(),
        donnees.nom || '',
        donnees.contact || '',
        donnees.absent === 'oui' ? 'Absent' : 'Présent',
        evenements,
        donnees.nombre || '1',
        donnees.message || ''
    ]);

    return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
