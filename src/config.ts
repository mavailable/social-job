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
] as const;

type FtColumnKey = (typeof FT_SHEET_COLUMNS)[number]['key'];

const FT_COL_INDEX: Record<FtColumnKey, number> = FT_SHEET_COLUMNS.reduce(
  (acc, c, i) => {
    (acc as any)[c.key] = i + 1;
    return acc;
  },
  {} as Record<FtColumnKey, number>
);

function ftCol_(key: FtColumnKey): number {
  const idx = FT_COL_INDEX[key];
  if (!idx) throw new Error(`❌ Colonne inconnue: ${String(key)}`);
  return idx;
}

function ftHeaders_(): string[] {
  return FT_SHEET_COLUMNS.map((c) => c.header);
}

const FT_DEFAULTS = {
  FT_CLIENT_ID: 'TON_CLIENT_ID_ICI',
  FT_CLIENT_SECRET: 'TON_CLIENT_SECRET_ICI'
} as const;

const FT_TOKEN_URL =
  'https://entreprise.francetravail.fr/connexion/oauth2/access_token?realm=%2Fpartenaire';

const FT_SEARCH_URL =
  'https://api.francetravail.io/partenaire/offresdemploi/v2/offres/search';

const SHEET_NAME = 'FT - Travailleurs sociaux (<24h)';

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
] as const;

const FT_PAGE_SIZE = 150;
const FT_MAX_PAGES = 20;
