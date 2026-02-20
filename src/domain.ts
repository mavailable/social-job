type Exclusions = { intitule: string[]; entreprise: string[] };

function ftIsExcludedOffer_(intitule: unknown, entrepriseNom: unknown, exclusions: Exclusions): boolean {
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

function normalizeContactNom_(contactNom: unknown, entrepriseNom: unknown): string {
  const raw = String(contactNom || '').trim();
  if (!raw) return '';

  if (/^Agence\s+France\s+Travail\b/i.test(raw)) return '';

  const parts = raw.split(' - ');
  if (parts.length >= 2) {
    const left = parts[0].trim();
    const right = parts.slice(1).join(' - ').trim();

    const ent = String(entrepriseNom || '').trim();
    if (!right) return '';

    if (!ent) return right;

    const norm = (s: string) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
    if (norm(left) === norm(ent)) return right;

    if (norm(left).includes(norm(ent)) || norm(ent).includes(norm(left))) return right;
  }

  return raw;
}

function sortOffers_(offers: any[]): any[] {
  return offers.sort((a, b) => {
    const aEnt = (a?.entreprise?.nom || '').trim();
    const bEnt = (b?.entreprise?.nom || '').trim();

    const aContact = normalizeContactNom_((a?.contact?.nom || '').trim(), aEnt);
    const bContact = normalizeContactNom_((b?.contact?.nom || '').trim(), bEnt);

    const isTempsPartiel = (o: any) => /\btemps\s+partiel\b/i.test(String(o?.dureeTravailLibelle || ''));

    const score = (o: any, contact: string, ent: string) =>
      (contact ? 2 : 0) + (ent ? 1 : 0) + (isTempsPartiel(o) ? 0.5 : 0);

    const sa = score(a, aContact, aEnt);
    const sb = score(b, bContact, bEnt);
    if (sb !== sa) return sb - sa;

    const da = a?.dateCreation ? new Date(a.dateCreation).getTime() : 0;
    const db = b?.dateCreation ? new Date(b.dateCreation).getTime() : 0;
    if (db !== da) return db - da;

    const ia = String(a?.id || '');
    const ib = String(b?.id || '');
    return ia.localeCompare(ib);
  });
}
