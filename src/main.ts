// Entry points for Google Apps Script
// Apps Script does not support Node-style modules at runtime; entry points must be global.

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('France Travail')
    .addItem('Configurer FT_CLIENT_ID / FT_CLIENT_SECRET', 'ftConfigureSecrets_')
    .addSeparator()
    .addItem('MAJ Travailleurs sociaux (<24h)', 'ftUpdateTravailleurSocial_24h')
    .addToUi();
}
