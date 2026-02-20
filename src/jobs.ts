function ftUpdateTravailleurSocial_24h() {
  const sh = createRunSheet_();
  const existingUrls = ftLoadExistingOfferUrls_(sh);

  const exclusions = ftLoadExclusions_();

  const baseParams = {
    motsCles: 'travailleur social',
    publieeDepuis: 1
  };

  const resultats = ftFetchAllResults_(baseParams);
  sortOffers_(resultats);

  const rows: any[][] = [];
  const urlsForRow: string[] = [];
  const descriptionsForRow: string[] = [];
  const entrepriseDescriptionsForRow: string[] = [];

  for (const o of resultats) {
    const id = o?.id ? String(o.id) : '';
    const urlOffre = id
      ? `https://candidat.francetravail.fr/offres/recherche/detail/${id}`
      : o?.origineOffre?.urlOrigine || '';

    if (!urlOffre || existingUrls.has(urlOffre)) continue;

    const entrepriseNom = o?.entreprise?.nom || '';
    const contactNom = normalizeContactNom_(o?.contact?.nom || '', entrepriseNom);

    if (ftIsExcludedOffer_(o?.intitule || '', entrepriseNom, exclusions)) {
      continue;
    }

    const descFull = String(o?.description || '');
    const descFirstLine = descFull.split(/\r?\n/)[0] || '';

    const entDescFull = String(o?.entreprise?.description || '');
    const entDescFirstLine = entDescFull.split(/\r?\n/)[0] || '';

    rows.push([
      'Nouveau',
      '',
      formatFTDate_(o?.dateCreation),
      o?.intitule || '',
      descFirstLine,
      entrepriseNom,
      contactNom,
      ftFormatLieuLibelle_(o?.lieuTravail?.libelle || ''),
      o?.lieuTravail?.codePostal || '',
      o?.typeContrat || '',
      o?.typeContratLibelle || '',
      o?.natureContrat || '',
      o?.dureeTravailLibelle || '',
      o?.salaire?.libelle || '',
      o?.experienceLibelle || '',
      o?.qualificationLibelle || '',
      // France Travail: champ attendu = secteurActiviteLibelle (typo fréquente dans le code)
      o?.secteurActiviteLibelle || o?.secteurActiviteLibele || o?.secteurActivite || '',
      o?.romeCode || o?.codeRome || '',
      o?.appellationlibelle || '',
      safeJoin_(o?.competences, 'libelle'),
      safeJoin_(o?.formations, 'libelle'),
      safeJoin_(o?.permis, 'libelle'),
      safeJoin_(o?.langues, 'libelle'),
      normalizeBool_(o?.alternance),
      normalizeBool_(o?.accessibleTH),
      o?.contact?.courriel || o?.contact?.email || '',
      o?.contact?.telephone || '',
      entDescFirstLine,
      // Colonne technique (masquée)
      urlOffre
    ]);

    urlsForRow.push(urlOffre);
    descriptionsForRow.push(descFull);
    entrepriseDescriptionsForRow.push(entDescFull);
  }

  if (!rows.length) {
    Logger.log('✅ Aucune nouvelle offre');
    return;
  }

  const startRow = appendRows_(sh, rows);
  applyRichTexts_(sh, startRow, rows.length, urlsForRow, descriptionsForRow, entrepriseDescriptionsForRow);

  // Resize colonnes (à la fin de l'import)
  // A 100, B 200, C 75, D 300, E 200, F 150, G 150, H 100, I 75, J 50, K 75
  // puis 100 pour toutes les suivantes.
  const fixedWidths = [100, 200, 75, 300, 200, 150, 150, 100, 75, 50, 75];
  fixedWidths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const lastCol = sh.getLastColumn();
  if (lastCol > fixedWidths.length) {
    sh.setColumnWidths(fixedWidths.length + 1, lastCol - fixedWidths.length, 100);
  }

  // Verrouille la hauteur des lignes une dernière fois sur toute la zone data.
  // (Sheets peut recalculer un auto-fit pendant l'écriture; on force l'état final.)
  SpreadsheetApp.flush();
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) sh.setRowHeights(2, lastRow - 1, 21);

  Logger.log(`✅ ${rows.length} offres ajoutées`);
}
