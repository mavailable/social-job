const FT_SHEET_COLUMNS = [
  { key: 'status', header: 'Statut' },
  { key: 'notes', header: 'Notes' },
  { key: 'dateCreation', header: 'Date' },
  { key: 'intitule', header: 'Poste' },
  { key: 'description', header: 'Résumé' },
  { key: 'entreprise_nom', header: 'Entreprise' },
  { key: 'contact_nom', header: 'Contact' },
  { key: 'codePostal', header: 'CP' },
  { key: 'typeContratLibelle', header: 'Contrat' },
  { key: 'dureeTravailLibelle', header: '% ETP' },
  { key: 'contact_email', header: 'Email' },
  { key: 'contact_tel', header: 'Téléphone' },
  { key: 'entreprise_description', header: 'À propos de l’entreprise' },
  // Colonne technique (masquée): URL de l'offre (sert à la déduplication)
  { key: 'url_offre', header: 'url_offre' }
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

type FtStatus = (typeof STATUS_VALUES)[number];

type FtStatusStyle = {
  /** Background color (hex) */
  bg: string;
  /** Font color (hex) */
  fg: string;
  /** Optional richer label displayed in cells (e.g. with dot) */
  label?: string;
};

// Palette "premium" (lisible + contrastée).
const FT_STATUS_STYLES: Record<FtStatus, FtStatusStyle> = {
  'Nouveau': { label: '● Nouveau', bg: '#E3F2FD', fg: '#0D47A1' },
  'À contacter': { label: '● À contacter', bg: '#FFF8E1', fg: '#E65100' },
  'Contacté': { label: '● Contacté', bg: '#E8F5E9', fg: '#1B5E20' },
  'Relance 1': { label: '● Relance 1', bg: '#F3E5F5', fg: '#4A148C' },
  'Relance 2': { label: '● Relance 2', bg: '#EDE7F6', fg: '#311B92' },
  'Entretien': { label: '● Entretien', bg: '#E0F7FA', fg: '#006064' },
  'Refus': { label: '● Refus', bg: '#FFEBEE', fg: '#B71C1C' },
  'OK / Piste chaude': { label: '● OK / Piste chaude', bg: '#E8F5E9', fg: '#2E7D32' },
  'Clôturé': { label: '● Clôturé', bg: '#ECEFF1', fg: '#263238' }
};

function ftIsStatus_(v: unknown): v is FtStatus {
  return (STATUS_VALUES as readonly string[]).includes(String(v));
}

function ftStatusLabel_(status: FtStatus): string {
  return FT_STATUS_STYLES[status]?.label || status;
}

function ftStatusStyle_(status: FtStatus): FtStatusStyle {
  return FT_STATUS_STYLES[status] || { bg: '#FFFFFF', fg: '#000000' };
}

const FT_PAGE_SIZE = 150;
const FT_MAX_PAGES = 20;

/**
 * Ex: "25 - Sochaux" -> "Sochaux"
 */
function ftFormatLieuLibelle_(v: unknown): string {
  const s = String(v ?? '').trim();
  if (!s) return '';
  // Supprime un éventuel préfixe "<chiffres> - "
  return s.replace(/^\d+\s*-\s*/, '').trim();
}
