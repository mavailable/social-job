/*************************************************
 * FRANCE TRAVAIL – MVP TRAVAILLEUR SOCIAL < 24h
 * - Secrets en Script Properties (auto-init)
 * - Date format dd/MM/yyyy HH:mm
 * - Pas de colonnes origine/url_offre/url_origine
 * - Intitulé (col D) cliquable vers l’offre
 * - Dédoublonnage via l’URL stockée en NOTE sur l’intitulé
 *************************************************/

/***************
 * SCHEMA COLONNES (source de vérité)
 * - Centralise l'ordre des colonnes, leurs en-têtes et leurs indices (1-based)
 * - Évite les "magic numbers" lors des getRange()
 ***************/
const FT_SHEET_COLUMNS = [
  { key: 'status', header: 'Status' },
  { key: 'notes', header: 'Notes' },
  { key: 'dateCreation', header: 'dateCreation' },
  { key: 'intitule', header: 'intitule' },
  { key: 'description', header: 'description' },
  { key: 'entreprise_nom', header: 'entreprise_nom' },
  { key: 'contact_nom', header: 'contact_nom' },
  { key: 'entreprise_description', header: 'entreprise_description' },
  { key: 'lieu_libelle', header: 'lieu_libelle' },
  { key: 'codePostal', header: 'codePostal' },
  { key: 'typeContrat', header: 'typeContrat' },
  { key: 'typeContratLibelle', header: 'typeContratLibelle' },
  { key: 'natureContrat', header: 'natureContrat' },
  { key: 'dureeTravailLibelle', header: 'dureeTravailLibelle' },
  { key: 'salaire_libelle', header: 'salaire_libelle' },
  { key: 'experienceLibelle', header: 'experienceLibelle' },
  { key: 'qualificationLibelle', header: 'qualificationLibelle' },
  { key: 'secteurActiviteLibelle', header: 'secteurActiviteLibelle' },
  { key: 'codeROME', header: 'codeROME' },
  { key: 'appellationlibelle', header: 'appellationlibelle' },
  { key: 'competences', header: 'competences' },
  { key: 'formations', header: 'formations' },
  { key: 'permis', header: 'permis' },
  { key: 'langues', header: 'langues' },
  { key: 'alternance', header: 'alternance' },
  { key: 'accessibleTH', header: 'accessibleTH' },
  { key: 'contact_email', header: 'contact_email' },
  { key: 'contact_tel', header: 'contact_tel' }
];

// Indices 1-based (compatibles SpreadsheetApp.getRange)
const FT_COL_INDEX = FT_SHEET_COLUMNS.reduce((acc, c, i) => {
  acc[c.key] = i + 1;
  return acc;
}, {});

function ftCol_(key) {
  const idx = FT_COL_INDEX[key];
  if (!idx) throw new Error(`❌ Colonne inconnue: ${key}`);
  return idx;
}

function ftHeaders_() {
  return FT_SHEET_COLUMNS.map(c => c.header);
}

/***************
 * DEFAULTS (à modifier UNE FOIS)
 ***************/
const FT_DEFAULTS = {
  FT_CLIENT_ID: 'TON_CLIENT_ID_ICI',
  FT_CLIENT_SECRET: 'TON_CLIENT_SECRET_ICI'
};

const FT_TOKEN_URL =
  'https://entreprise.francetravail.fr/connexion/oauth2/access_token?realm=%2Fpartenaire';

const FT_SEARCH_URL =
  'https://api.francetravail.io/partenaire/offresdemploi/v2/offres/search';

const SHEET_NAME = 'FT - Travailleurs sociaux (<24h)';

// Onglet contenant les exclusions (2 colonnes):
// A: intitule (mots-clés à exclure si présents dans l'intitulé)
// B: entreprise_nom (mots-clés à exclure si présents dans le nom de l'entreprise)
const EXCLUSIONS_SHEET_NAME = 'Exclusions';

const STATUS_VALUES = [
  'Nouveau',
  'À contacter',
  'Contacté',
  'Relance 1',
  'Relance 2',
  'Entretien',
  'Refus',
  'OK / Piste chaude',
  'Clôturé'
];

// Pagination France Travail: l'API limite la taille de range (max 150 résultats par appel)
const FT_PAGE_SIZE = 150;
// Limite de sécurité pour éviter de boucler indéfiniment (et rester dans les quotas UrlFetch)
const FT_MAX_PAGES = 20;

/***************
 * SECRETS – AUTO INIT
 ***************/
function ftEnsureSecrets_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FT_CLIENT_ID');
  const secret = props.getProperty('FT_CLIENT_SECRET');

  if (id && secret) return;

  // Si des valeurs par défaut sont fournies dans le code, on les utilise.
  if (FT_DEFAULTS.FT_CLIENT_ID && FT_DEFAULTS.FT_CLIENT_SECRET &&
    FT_DEFAULTS.FT_CLIENT_ID !== 'TON_CLIENT_ID_ICI' &&
    FT_DEFAULTS.FT_CLIENT_SECRET !== 'TON_CLIENT_SECRET_ICI') {
    props.setProperties({
      FT_CLIENT_ID: FT_DEFAULTS.FT_CLIENT_ID,
      FT_CLIENT_SECRET: FT_DEFAULTS.FT_CLIENT_SECRET
    });
    Logger.log('✅ Secrets France Travail initialisés depuis FT_DEFAULTS');
    return;
  }

  // Sinon: demander à l'utilisateur via popup (nécessite un contexte UI Spreadsheet).
  ftPromptAndSaveSecrets_();
}

/**
 * Demande à l'utilisateur de saisir FT_CLIENT_ID / FT_CLIENT_SECRET puis les sauvegarde.
 * À appeler uniquement dans un contexte où SpreadsheetApp.getUi() est disponible.
 */
function ftPromptAndSaveSecrets_() {
  // En exécution hors UI (ex: trigger), SpreadsheetApp.getUi() peut ne pas être dispo.
  // Dans ce cas on stoppe avec un message explicite.
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    const msg = '❌ FT_CLIENT_ID / FT_CLIENT_SECRET manquants. ' +
      'Le script est exécuté hors interface (trigger / exécution serveur), ' +
      'impossible d’afficher une popup. Lancez la configuration depuis le Google Sheet: ' +
      'Menu "France Travail" → "Configurer FT_CLIENT_ID / FT_CLIENT_SECRET".';
    Logger.log(msg);
    throw new Error(msg);
  }

  const props = PropertiesService.getScriptProperties();

  const idResp = ui.prompt(
    'Configuration France Travail',
    'Saisissez votre FT_CLIENT_ID :',
    ui.ButtonSet.OK_CANCEL
  );
  if (idResp.getSelectedButton() !== ui.Button.OK) {
    throw new Error('❌ Configuration annulée (FT_CLIENT_ID manquant)');
  }
  const clientId = (idResp.getResponseText() || '').trim();
  if (!clientId) {
    throw new Error('❌ FT_CLIENT_ID vide');
  }

  const secretResp = ui.prompt(
    'Configuration France Travail',
    'Saisissez votre FT_CLIENT_SECRET :',
    ui.ButtonSet.OK_CANCEL
  );
  if (secretResp.getSelectedButton() !== ui.Button.OK) {
    throw new Error('❌ Configuration annulée (FT_CLIENT_SECRET manquant)');
  }
  const clientSecret = (secretResp.getResponseText() || '').trim();
  if (!clientSecret) {
    throw new Error('❌ FT_CLIENT_SECRET vide');
  }

  props.setProperties({
    FT_CLIENT_ID: clientId,
    FT_CLIENT_SECRET: clientSecret
  });

  ui.alert('✅ Identifiants enregistrés dans Script Properties.');
}

/**
 * Permet de (re)configurer manuellement les secrets via le menu.
 */
function ftConfigureSecrets_() {
  ftPromptAndSaveSecrets_();
}

function getFTSecrets_() {
  ftEnsureSecrets_();
  const props = PropertiesService.getScriptProperties();
  return {
    clientId: props.getProperty('FT_CLIENT_ID'),
    clientSecret: props.getProperty('FT_CLIENT_SECRET')
  };
}

/***************
 * DATE FORMAT
 ***************/
function formatFTDate_(isoString) {
  if (!isoString) return '';
  const d = new Date(isoString);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

/***************
 * TOKEN (avec cache)
 ***************/
function ftGetToken_() {
  const { clientId, clientSecret } = getFTSecrets_();

  const cache = CacheService.getScriptCache();
  const cached = cache.get('FT_ACCESS_TOKEN');
  if (cached) return cached;

  const payload = {
    grant_type: 'client_credentials',
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'api_offresdemploiv2 o2dsoffre'
  };

  const res = UrlFetchApp.fetch(FT_TOKEN_URL, {
    method: 'post',
    payload,
    muteHttpExceptions: true
  });

  if (res.getResponseCode() >= 300) {
    throw new Error('❌ Token error: ' + res.getContentText());
  }

  const data = JSON.parse(res.getContentText());
  CacheService.getScriptCache().put('FT_ACCESS_TOKEN', data.access_token, 50 * 60);
  return data.access_token;
}

/***************
 * UTILS
 ***************/
function safeJoin_(arr, key) {
  if (!Array.isArray(arr)) return '';
  return arr.map(o => (key ? o?.[key] : o)).filter(Boolean).join(' | ');
}

function normalizeBool_(v) {
  if (v === true) return 'TRUE';
  if (v === false) return 'FALSE';
  return '';
}

function buildGoogleSearchUrl_(entrepriseNom, lieuLibelle) {
  if (!entrepriseNom || !lieuLibelle) return '';

  // Supprime les 4 premiers caractères du lieu (format type "XXXX - Ville")
  const lieu = lieuLibelle.length > 4
    ? lieuLibelle.substring(4)
    : lieuLibelle;

  const query = `${entrepriseNom} ${lieu}`.trim();
  return 'https://www.google.com/search?q=' + encodeURIComponent(query);
}

function buildLinkedInPeopleSearchUrl_(contactNom, entrepriseNom) {
  if (!contactNom && !entrepriseNom) return '';

  const sanitize = (s) => {
    const str = (s || '').trim();
    if (!str) return '';

    // Retire uniquement la civilité en début de chaîne
    return str
      .replace(/^(m\.\s*|mr\.?\s*|mme\.?\s*|mrs\.?\s*|mlle\.?\s*|mademoiselle\s+|monsieur\s+|madame\s+)/i, '')
      .trim();
  };

  const query = `${sanitize(contactNom)} ${sanitize(entrepriseNom)}`.trim();
  if (!query) return '';
  return 'https://www.linkedin.com/search/results/people/?keywords=' + encodeURIComponent(query);
}

function normalizeKeyword_(s) {
  return String(s || '').trim().toLowerCase();
}

// Charge les mots-clés d'exclusion depuis l'onglet "Exclusions"
function ftLoadExclusions_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(EXCLUSIONS_SHEET_NAME);
  if (!sh) return { intitule: [], entreprise: [] };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { intitule: [], entreprise: [] };

  // On lit A2:B
  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();

  const intitule = [];
  const entreprise = [];

  for (const [a, b] of values) {
    const ka = normalizeKeyword_(a);
    const kb = normalizeKeyword_(b);
    if (ka) intitule.push(ka);
    if (kb) entreprise.push(kb);
  }

  return { intitule, entreprise };
}

function ftIsExcludedOffer_(intitule, entrepriseNom, exclusions) {
  const title = normalizeKeyword_(intitule);
  const ent = normalizeKeyword_(entrepriseNom);
  const ex = exclusions || { intitule: [], entreprise: [] };

  if (title && ex.intitule?.length) {
    for (const kw of ex.intitule) {
      if (kw && title.includes(kw)) return true;
    }
  }

  if (ent && ex.entreprise?.length) {
    for (const kw of ex.entreprise) {
      if (kw && ent.includes(kw)) return true;
    }
  }

  return false;
}

/**
 * Normalise le champ contact_nom :
 * - si ça commence par "Agence France Travail" => vide
 * - si c'est du type "ENTREPRISE - Nom Prénom" => on ne garde que la partie après " - "
 */
function normalizeContactNom_(contactNom, entrepriseNom) {
  const raw = (contactNom || '').trim();
  if (!raw) return '';

  // Cas 1: contact interne FT => on ne garde rien
  if (/^Agence\s+France\s+Travail\b/i.test(raw)) return '';

  // Cas 2: "Entreprise - Nom" => on retire le préfixe
  const parts = raw.split(' - ');
  if (parts.length >= 2) {
    const left = parts[0].trim();
    const right = parts.slice(1).join(' - ').trim();

    const ent = (entrepriseNom || '').trim();
    if (!right) return '';

    if (!ent) return right;

    const norm = s => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
    if (norm(left) === norm(ent)) return right;

    // Fallback: si l'entreprise est contenue dans le préfixe (ou inversement), on retire quand même
    if (norm(left).includes(norm(ent)) || norm(ent).includes(norm(left))) return right;
  }

  return raw;
}

/***************
 * SHEET SETUP
 ***************/
// Nom d'onglet créé à chaque exécution
function ftBuildRunSheetName_() {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
  // Google Sheets interdit certains caractères dans les noms d'onglets: \ / ? * [ ]
  // ':' est autorisé et conservé ici.
  return ts.replace(/[\\/\?\*\[\]]/g, '-');
}

// Crée un nouvel onglet pour une exécution, et applique l'en-tête + validations
function createRunSheet_() {
  const ss = SpreadsheetApp.getActive();
  const name = ftBuildRunSheetName_();

  // Évite collision (ex: 2 exécutions dans la même minute)
  let finalName = name;
  let i = 2;
  while (ss.getSheetByName(finalName)) {
    finalName = `${name} (${i++})`;
  }

  const sh = ss.insertSheet(finalName);
  setupFTSheet_(sh);
  return sh;
}

function setupFTSheet_(sh) {
  // Colonnes finales (sans origine/url_offre/url_origine)
  // A Status, B Notes, C dateCreation, D intitule (cliquable), etc.
  // NOTE: contact_nom est placé juste après entreprise_nom.
  const headers = ftHeaders_();

  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_VALUES, true)
    .setAllowInvalid(false)
    .build();

  sh.getRange(2, ftCol_('status'), sh.getMaxRows(), 1).setDataValidation(rule);
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  if (sh.getLastRow() === 0) {
    setupFTSheet_(sh);
  }

  return sh;
}

/***************
 * MAIN
 ***************/
function ftUpdateTravailleurSocial_24h() {
  let token = ftGetToken_();
  // À chaque exécution, on crée un nouvel onglet daté
  const sh = createRunSheet_();
  // Dédoublonnage: sur un onglet "run" on n'a pas d'existant (onglet vide)
  const existingUrls = new Set();

  // Exclusions (onglet "Exclusions")
  const exclusions = ftLoadExclusions_();

  const baseParams = {
    motsCles: 'travailleur social',
    publieeDepuis: 1
  };

  // On agrège les résultats de plusieurs pages (0-149, 150-299, ...)
  const resultats = [];
  let retriedAfter401 = false;

  for (let page = 0; page < FT_MAX_PAGES; page++) {
    const start = page * FT_PAGE_SIZE;
    const end = start + FT_PAGE_SIZE - 1;

    const params = { ...baseParams, range: `${start}-${end}` };
    const url = FT_SEARCH_URL + '?' + Object.entries(params)
      .map(([k, v]) => `${k}=${encodeURIComponent(v)}`)
      .join('&');

    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const body = res.getContentText() || '';

    // 401: token invalide/expiré. On purge le cache et on retente une fois.
    if (code === 401 && !retriedAfter401) {
      CacheService.getScriptCache().remove('FT_ACCESS_TOKEN');
      token = ftGetToken_();
      retriedAfter401 = true;
      page--; // rejoue la même page
      continue;
    }

    // Certains 400 correspondent à "plage trop importante" quand on va trop loin ; dans ce cas on stoppe.
    if (code === 400) {
      try {
        const j = JSON.parse(body || '{}');
        const msg = String(j?.message || '');
        if (/plage\s+de\s+r[ée]sultats\s+demand[ée]e\s+est\s+trop\s+importante/i.test(msg)) break;
      } catch (e) {
        // ignore
      }
      throw new Error(`❌ Search error (HTTP ${code}): ${body}`);
    }

    if (code >= 300) {
      throw new Error(`❌ Search error (HTTP ${code}): ${body}`);
    }

    const data = JSON.parse(body || '{}');
    const pageResultats = data.resultats || [];
    if (!pageResultats.length) break;

    resultats.push(...pageResultats);

    // Dernière page si on a moins que FT_PAGE_SIZE
    if (pageResultats.length < FT_PAGE_SIZE) break;
  }

  // Tri de priorité:
  // 1) offres avec contact_nom
  // 2) puis offres avec entreprise_nom
  // (en gardant un comportement déterministe si égalité)
  resultats.sort((a, b) => {
    const aEnt = (a?.entreprise?.nom || '').trim();
    const bEnt = (b?.entreprise?.nom || '').trim();

    const aContact = normalizeContactNom_((a?.contact?.nom || '').trim(), aEnt);
    const bContact = normalizeContactNom_((b?.contact?.nom || '').trim(), bEnt);

    const score = (contact, ent) => (contact ? 2 : 0) + (ent ? 1 : 0);
    const sa = score(aContact, aEnt);
    const sb = score(bContact, bEnt);
    if (sb !== sa) return sb - sa;

    // Tie-breakers
    const da = a?.dateCreation ? new Date(a.dateCreation).getTime() : 0;
    const db = b?.dateCreation ? new Date(b.dateCreation).getTime() : 0;
    if (db !== da) return db - da;

    const ia = String(a?.id || '');
    const ib = String(b?.id || '');
    return ia.localeCompare(ib);
  });

  // On va d’abord écrire les valeurs "simples", puis appliquer les liens (RichText) + notes sur col D
  const rows = [];
  const urlsForRow = []; // URL offre, dans le même ordre que rows

  for (const o of resultats) {
    const id = o?.id ? String(o.id) : '';
    const urlOffre = id
      ? `https://candidat.francetravail.fr/offres/recherche/detail/${id}`
      : (o?.origineOffre?.urlOrigine || '');

    if (!urlOffre || existingUrls.has(urlOffre)) continue;

    const entrepriseNom = o?.entreprise?.nom || '';
    const contactNom = normalizeContactNom_(o?.contact?.nom || '', entrepriseNom);

    // Exclusions: si l'intitulé contient un mot-clé de col A
    // ou si entreprise_nom contient un mot-clé de col B => on ignore l'offre
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

  const startRow = sh.getLastRow() + 1;

  // 1) Écriture brute
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

  // Fixe la hauteur des lignes ajoutées (empêche l'adaptation au contenu)
  sh.setRowHeights(startRow, rows.length, 21);

  // 2) Validation status sur les lignes ajoutées
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_VALUES, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(startRow, ftCol_('status'), rows.length, 1).setDataValidation(rule);

  // 3) Ajouter les liens sur la colonne intitule + stocker urlOffre dans la NOTE pour dédoublonner
  const intituleRange = sh.getRange(startRow, ftCol_('intitule'), rows.length, 1);
  const intitules = intituleRange.getValues().flat();

  const richValues = intitules.map((txt, i) => {
    return SpreadsheetApp.newRichTextValue()
      .setText(txt || 'Offre')
      .setLinkUrl(urlsForRow[i])
      .build();
  });

  // 4) Ajouter lien Google sur entreprise_nom si non vide
  const entrepriseRange = sh.getRange(startRow, ftCol_('entreprise_nom'), rows.length, 1);
  const lieuRange = sh.getRange(startRow, ftCol_('lieu_libelle'), rows.length, 1);

  const entreprises = entrepriseRange.getValues().flat();
  const lieux = lieuRange.getValues().flat();

  const entrepriseRichValues = entreprises.map((name, i) => {
    if (!name) return SpreadsheetApp.newRichTextValue().setText('').build();

    const searchUrl = buildGoogleSearchUrl_(name, lieux[i]);
    if (!searchUrl) return SpreadsheetApp.newRichTextValue().setText(name).build();

    return SpreadsheetApp.newRichTextValue()
      .setText(name)
      .setLinkUrl(searchUrl)
      .build();
  });

  entrepriseRange.setRichTextValues(entrepriseRichValues.map(v => [v]));

  // 5) Ajouter lien LinkedIn sur contact_nom
  const contactNomRange = sh.getRange(startRow, ftCol_('contact_nom'), rows.length, 1);
  const contactNoms = contactNomRange.getValues().flat();

  const contactNomRichValues = contactNoms.map((contactNom, i) => {
    const txt = contactNom || '';
    if (!txt) return SpreadsheetApp.newRichTextValue().setText('').build();

    const linkedinUrl = buildLinkedInPeopleSearchUrl_(txt, entreprises[i]);
    if (!linkedinUrl) return SpreadsheetApp.newRichTextValue().setText(txt).build();

    return SpreadsheetApp.newRichTextValue()
      .setText(txt)
      .setLinkUrl(linkedinUrl)
      .build();
  });

  contactNomRange.setRichTextValues(contactNomRichValues.map(v => [v]));

  intituleRange.setRichTextValues(richValues.map(v => [v]));
  intituleRange.setNotes(urlsForRow.map(u => [u]));

  Logger.log(`✅ ${rows.length} offres ajoutées`);
}

/***************
 * MENU
 ***************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('France Travail')
    .addItem('Configurer FT_CLIENT_ID / FT_CLIENT_SECRET', 'ftConfigureSecrets_')
    .addSeparator()
    .addItem('MAJ Travailleurs sociaux (<24h)', 'ftUpdateTravailleurSocial_24h')
    .addToUi();
}