type SheetExclusions = { intitule: string[]; entreprise: string[] };

// URLs des offres déjà importées : on les stocke dans les Notes de la colonne `intitule`.
function ftLoadExistingOfferUrls_(sh: GoogleAppsScript.Spreadsheet.Sheet): Set<string> {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return new Set();

  const range = sh.getRange(2, ftCol_('intitule'), lastRow - 1, 1);
  const notes = range.getNotes(); // string[][]

  const urls = new Set<string>();
  for (const row of notes) {
    const note = (row?.[0] || '').trim();
    if (note) urls.add(note);
  }

  return urls;
}

function ftBuildRunSheetName_(): string {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
  return ts.replace(/[\\/\?\*\[\]]/g, '-');
}

function setupFTSheet_(sh: GoogleAppsScript.Spreadsheet.Sheet) {
  const headers = ftHeaders_();
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  // Force un style de lignes "fixe" (pas de Fit to data) sur toute la zone data.
  // Sans ça, certaines lignes peuvent rester en auto-ajustement et grossir selon le contenu.
  const dataRows = sh.getMaxRows() - 1;
  if (dataRows > 0) {
    sh.setRowHeights(2, dataRows, 21);
    const dataRange = sh.getRange(2, 1, dataRows, sh.getMaxColumns());
    dataRange.setWrap(false);
    // Force une stratégie d'affichage qui n'agrandit pas les lignes.
    // (Sur certaines feuilles, le mode "wrap" peut être réappliqué via styles.)
    dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    dataRange.setVerticalAlignment('middle');
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([...STATUS_VALUES], true)
    .setAllowInvalid(false)
    .build();

  sh.getRange(2, ftCol_('status'), sh.getMaxRows(), 1).setDataValidation(rule);

  // Rend la colonne "Status" plus premium dès la création.
  const statusRange = sh.getRange(2, ftCol_('status'), sh.getMaxRows(), 1);
  statusRange.setFontFamily('Inter');
  statusRange.setFontSize(10);
  statusRange.setFontWeight('bold');
  statusRange.setHorizontalAlignment('center');
  statusRange.setVerticalAlignment('middle');

  // IMPORTANT: pour respecter la validation de données, la cellule doit contenir
  // une valeur de STATUS_VALUES. Les couleurs sont appliquées via formats conditionnels.
  ftEnsureStatusConditionalFormatting_(sh);
}

function ftEnsureStatusConditionalFormatting_(sh: GoogleAppsScript.Spreadsheet.Sheet) {
  const range = sh.getRange(2, ftCol_('status'), sh.getMaxRows(), 1);

  const existing = sh.getConditionalFormatRules();
  const kept = existing.filter((r) => {
    const ranges = r.getRanges();
    return !ranges.some((rg) => rg.getA1Notation() === range.getA1Notation());
  });

  const rules: GoogleAppsScript.Spreadsheet.ConditionalFormatRule[] = [];

  (STATUS_VALUES as readonly string[]).forEach((s) => {
    const status = s as any;
    const st = ftStatusStyle_(status);

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(String(s))
        .setBackground(st.bg)
        .setFontColor(st.fg)
        .setBold(true)
        .setRanges([range])
        .build()
    );
  });

  sh.setConditionalFormatRules([...kept, ...rules]);
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

  // Empêche le contenu (wrap) de faire paraître les lignes plus hautes.
  // (Google Sheets peut ajuster l'affichage si le texte est renvoyé à la ligne.)
  const rowRange = sh.getRange(startRow, 1, rows.length, rows[0].length);
  rowRange.setWrap(false);
  rowRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  rowRange.setVerticalAlignment('middle');

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([...STATUS_VALUES], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(startRow, ftCol_('status'), rows.length, 1).setDataValidation(rule);

  // Assure que les couleurs "badge" sont bien en place.
  ftEnsureStatusConditionalFormatting_(sh);

  return startRow;
}

// Supprimé: ancienne implémentation qui écrivait un label enrichi (● ...)
// dans la cellule, ce qui viole la data validation.

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

    const lieu = ftFormatLieuLibelle_(lieux[i]);
    const searchUrl = buildGoogleSearchUrl_(name, lieu);
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

  // Réapplique une hauteur fixe: l'écriture de RichText peut déclencher
  // un auto-fit implicite selon le contenu/format.
  // IMPORTANT: un simple setRowHeights(..., 21) n'est pas toujours suffisant:
  // Sheets peut conserver un rendu "fit". On force donc un vrai changement
  // (20 -> flush -> 21) comme quand on le fait à la main.
  SpreadsheetApp.flush();
  sh.setRowHeights(startRow, rowsLength, 20);
  SpreadsheetApp.flush();
  sh.setRowHeights(startRow, rowsLength, 21);

  // Et empêche le wrap après application des RichText.
  const rowRange = sh.getRange(startRow, 1, rowsLength, sh.getLastColumn());
  rowRange.setWrap(false);
  rowRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Force aussi l'option UI "Overflow" (pas seulement CLIP).
  // Différence: CLIP est une stratégie de wrap, OVERFLOW force l'affichage sur 1 ligne.
  rowRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

  // Verrouille une dernière fois la hauteur sur toute la zone data existante.
  // (Sheets peut encore recalculer après avoir appliqué les styles/rich text.)
  SpreadsheetApp.flush();
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.setRowHeights(2, lastRow - 1, 20);
    SpreadsheetApp.flush();
    sh.setRowHeights(2, lastRow - 1, 21);
  }
}
