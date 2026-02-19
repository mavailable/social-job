type SheetExclusions = { intitule: string[]; entreprise: string[] };

function ftBuildRunSheetName_(): string {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
  return ts.replace(/[\\/\?\*\[\]]/g, '-');
}

function setupFTSheet_(sh: GoogleAppsScript.Spreadsheet.Sheet) {
  const headers = ftHeaders_();
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([...STATUS_VALUES], true)
    .setAllowInvalid(false)
    .build();

  sh.getRange(2, ftCol_('status'), sh.getMaxRows(), 1).setDataValidation(rule);
}

function createRunSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActive();
  const name = ftBuildRunSheetName_();

  let finalName = name;
  let i = 2;
  while (ss.getSheetByName(finalName)) {
    finalName = `${name} (${i++})`;
  }

  const sh = ss.insertSheet(finalName);
  setupFTSheet_(sh);
  return sh;
}

function getOrCreateSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  if (sh.getLastRow() === 0) {
    setupFTSheet_(sh);
  }

  return sh;
}

function ftLoadExclusions_(): SheetExclusions {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(EXCLUSIONS_SHEET_NAME);
  if (!sh) return { intitule: [], entreprise: [] };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { intitule: [], entreprise: [] };

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();

  const intitule: string[] = [];
  const entreprise: string[] = [];

  for (const [a, b] of values) {
    const ka = normalizeKeyword_(a);
    const kb = normalizeKeyword_(b);
    if (ka) intitule.push(ka);
    if (kb) entreprise.push(kb);
  }

  return { intitule, entreprise };
}

function appendRows_(sh: GoogleAppsScript.Spreadsheet.Sheet, rows: any[][]): number {
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  sh.setRowHeights(startRow, rows.length, 21);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([...STATUS_VALUES], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(startRow, ftCol_('status'), rows.length, 1).setDataValidation(rule);

  return startRow;
}

function applyRichTexts_(
  sh: GoogleAppsScript.Spreadsheet.Sheet,
  startRow: number,
  rowsLength: number,
  urlsForRow: string[]
) {
  const intituleRange = sh.getRange(startRow, ftCol_('intitule'), rowsLength, 1);
  const intitules = intituleRange.getValues().flat();

  const richValues = intitules.map((txt, i) => {
    return SpreadsheetApp.newRichTextValue()
      .setText((txt as string) || 'Offre')
      .setLinkUrl(urlsForRow[i])
      .build();
  });

  const entrepriseRange = sh.getRange(startRow, ftCol_('entreprise_nom'), rowsLength, 1);
  const lieuRange = sh.getRange(startRow, ftCol_('lieu_libelle'), rowsLength, 1);
  const entreprises = entrepriseRange.getValues().flat() as string[];
  const lieux = lieuRange.getValues().flat() as string[];

  const entrepriseRichValues = entreprises.map((name, i) => {
    if (!name) return SpreadsheetApp.newRichTextValue().setText('').build();

    const searchUrl = buildGoogleSearchUrl_(name, lieux[i]);
    if (!searchUrl) return SpreadsheetApp.newRichTextValue().setText(name).build();

    return SpreadsheetApp.newRichTextValue().setText(name).setLinkUrl(searchUrl).build();
  });
  entrepriseRange.setRichTextValues(entrepriseRichValues.map((v) => [v]));

  const contactNomRange = sh.getRange(startRow, ftCol_('contact_nom'), rowsLength, 1);
  const contactNoms = contactNomRange.getValues().flat() as string[];

  const contactNomRichValues = contactNoms.map((contactNom, i) => {
    const txt = contactNom || '';
    if (!txt) return SpreadsheetApp.newRichTextValue().setText('').build();

    const linkedinUrl = buildLinkedInPeopleSearchUrl_(txt, entreprises[i]);
    if (!linkedinUrl) return SpreadsheetApp.newRichTextValue().setText(txt).build();

    return SpreadsheetApp.newRichTextValue().setText(txt).setLinkUrl(linkedinUrl).build();
  });
  contactNomRange.setRichTextValues(contactNomRichValues.map((v) => [v]));

  intituleRange.setRichTextValues(richValues.map((v) => [v]));
  intituleRange.setNotes(urlsForRow.map((u) => [u]));
}
