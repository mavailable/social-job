function ftUpdateTravailleurSocial_24h() {
  const sh = createRunSheet_();
  const existingUrls = new Set<string>();

  const exclusions = ftLoadExclusions_();

  const baseParams = {
    motsCles: 'travailleur social',
    publieeDepuis: 1
  };

  const resultats = ftFetchAllResults_(baseParams);
  sortOffers_(resultats);

  const rows: any[][] = [];
  const urlsForRow: string[] = [];

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

    rows.push([
      'Nouveau',
      '',
      formatFTDate_(o?.dateCreation),
      o?.intitule || '',
      o?.description || '',
      entrepriseNom,
      contactNom,
      o?.entreprise?.description || '',
      o?.lieuTravail?.libelle || '',
      o?.lieuTravail?.codePostal || '',
      o?.typeContrat || '',
      o?.typeContratLibelle || '',
      o?.natureContrat || '',
      o?.dureeTravailLibelle || '',
      o?.salaire?.libelle || '',
      o?.experienceLibelle || '',
      o?.qualificationLibelle || '',
      o?.secteurActiviteLibele || '',
      o?.romeCode || o?.codeRome || '',
      o?.appellationlibelle || '',
      safeJoin_(o?.competences, 'libelle'),
      safeJoin_(o?.formations, 'libelle'),
      safeJoin_(o?.permis, 'libelle'),
      safeJoin_(o?.langues, 'libelle'),
      normalizeBool_(o?.alternance),
      normalizeBool_(o?.accessibleTH),
      o?.contact?.courriel || o?.contact?.email || '',
      o?.contact?.telephone || ''
    ]);

    urlsForRow.push(urlOffre);
  }

  if (!rows.length) {
    Logger.log('✅ Aucune nouvelle offre');
    return;
  }

  const startRow = appendRows_(sh, rows);
  applyRichTexts_(sh, startRow, rows.length, urlsForRow);

  Logger.log(`✅ ${rows.length} offres ajoutées`);
}
