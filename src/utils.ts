function safeJoin_(arr: unknown, key?: string): string {
  if (!Array.isArray(arr)) return '';
  return (arr as any[])
    .map((o) => (key ? o?.[key] : o))
    .filter(Boolean)
    .join(' | ');
}

function normalizeBool_(v: unknown): string {
  if (v === true) return 'TRUE';
  if (v === false) return 'FALSE';
  return '';
}

function formatFTDate_(isoString: unknown): string {
  if (!isoString) return '';
  const d = new Date(String(isoString));
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

function buildGoogleSearchUrl_(entrepriseNom: string, lieuLibelle: string): string {
  if (!entrepriseNom || !lieuLibelle) return '';

  // Ancien comportement: substring(4) (ex: "25 - Sochaux" -> "Sochaux")
  // cassait les lieux déjà propres (ex: "Sochaux" -> "aux").
  const lieu = String(lieuLibelle)
    .trim()
    .replace(/^\d+\s*-\s*/, '')
    .trim();

  const query = `${entrepriseNom} ${lieu}`.trim();
  return 'https://www.google.com/search?q=' + encodeURIComponent(query);
}

function buildLinkedInPeopleSearchUrl_(contactNom: string, entrepriseNom: string): string {
  if (!contactNom && !entrepriseNom) return '';

  const sanitize = (s: string) => {
    const str = (s || '').trim();
    if (!str) return '';
    return str
      .replace(
        /^(m\.\s*|mr\.?\s*|mme\.?\s*|mrs\.?\s*|mlle\.?\s*|mademoiselle\s+|monsieur\s+|madame\s+)/i,
        ''
      )
      .trim();
  };

  const query = `${sanitize(contactNom)} ${sanitize(entrepriseNom)}`.trim();
  if (!query) return '';
  return 'https://www.linkedin.com/search/results/people/?keywords=' + encodeURIComponent(query);
}

function normalizeKeyword_(s: unknown): string {
  // Normalisation "exclusions":
  // - trim
  // - lowercase
  // - suppression des accents (diacritiques)
  // - espaces multiples -> 1 espace
  return String(s || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

/**
 * Compile une règle d'exclusion.
 * - Si la règle est de la forme `/.../flags` => RegExp
 * - Sinon => comparaison "contains" sur le champ normalisé
 */
function parseExclusionRule_(raw: unknown): { raw: string; kind: 'regex'; re: RegExp } | { raw: string; kind: 'contains'; needle: string } {
  const s = String(raw || '').trim();
  if (!s) return { raw: '', kind: 'contains', needle: '' };

  // Format type: /pattern/gi
  if (s.startsWith('/') && s.lastIndexOf('/') > 0) {
    const lastSlash = s.lastIndexOf('/');
    const pattern = s.slice(1, lastSlash);
    const flags = s.slice(lastSlash + 1);
    try {
      return { raw: s, kind: 'regex', re: new RegExp(pattern, flags) };
    } catch (e) {
      // Fallback "contains" si la regex est invalide
      return { raw: s, kind: 'contains', needle: normalizeKeyword_(s) };
    }
  }

  return { raw: s, kind: 'contains', needle: normalizeKeyword_(s) };
}
