# Social Job – France Travail (Google Apps Script)

Ce projet automatise la **récupération d’offres France Travail** (API partenaire offres d’emploi v2) et l’**injection dans un Google Sheet** avec :
- dédoublonnage (via URL stockée en *note* de cellule),
- colonnes normalisées,
- liens cliquables (offre / recherche Google / LinkedIn),
- gestion d’exclusions (mots-clés sur intitulé / entreprise),
- menu Apps Script pour lancer la configuration et l’import.

Le code est écrit en **TypeScript** puis compilé en JS compatible Apps Script et poussé avec **clasp**.

---

## Sommaire
- [Fonctionnalités](#fonctionnalités)
- [Prérequis](#prérequis)
- [Installation](#installation)
- [Configuration](#configuration)
- [Utilisation (dans le Google Sheet)](#utilisation-dans-le-google-sheet)
- [Google Sheet : structure attendue](#google-sheet--structure-attendue)
- [Exclusions](#exclusions)
- [Développement](#développement)
- [Structure du projet](#structure-du-projet)
- [Dépannage](#dépannage)
- [Sécurité / secrets](#sécurité--secrets)
- [Notes techniques](#notes-techniques)

---

## Fonctionnalités

- **Menu personnalisé** dans Google Sheets : `France Travail`.
- **Configuration guidée** de `FT_CLIENT_ID` / `FT_CLIENT_SECRET` via popup.
- Import d’offres : **Travailleurs sociaux publiées depuis < 24h**.
- **Pagination** côté API (taille de page et limite de sécurité).
- **Tri** des offres (priorité aux offres avec contact / entreprise, puis par date).
- **Colonnes centralisées** (schéma unique dans le code).
- **Enrichissement** :
  - intitulé → lien vers l’offre France Travail,
  - entreprise → lien de recherche Google (entreprise + lieu),
  - contact → lien de recherche LinkedIn.

---

## Prérequis

- Node.js (>= 18 recommandé)
- Un projet **Google Apps Script** (idéalement *Spreadsheet-bound*, i.e. attaché à un Google Sheet)
- Accès à l’API **France Travail Partenaire** : un `client_id` et un `client_secret`

---

## Installation

1. Installer les dépendances :
   - `npm install`

2. Authentifier clasp (si nécessaire) :
   - `npm run login`

3. Associer le dépôt à un script existant (recommandé) :
   - créer/éditer `.clasp.json` à la racine
   - renseigner `scriptId`

   Exemple attendu :
   - `{"scriptId":"PASTE_SCRIPT_ID_HERE"}`

> Option : créer un script Apps Script depuis la CLI : `npm run create`.

---

## Configuration

### 1) Déployer le code dans Apps Script

- Build (compile TypeScript → `dist/`) :
  - `npm run build`

- Push vers Apps Script (build + push) :
  - `npm run push`

### 2) Autorisations Google

Au premier lancement depuis le Google Sheet, Apps Script demandera des autorisations, notamment :
- accès au tableur,
- appels réseau (`UrlFetchApp`) vers l’API France Travail.

### 3) Secrets France Travail

Les secrets sont stockés dans **Script Properties** (`PropertiesService.getScriptProperties()`).

Deux méthodes possibles :
- **Recommandée** : via le menu `France Travail` → `Configurer FT_CLIENT_ID / FT_CLIENT_SECRET`.
- Alternative : définir `FT_DEFAULTS` dans `src/config.ts` (éviter en usage partagé).

---

## Utilisation (dans le Google Sheet)

1. Ouvrir le Google Sheet lié au script.
2. Recharger la page (si besoin) pour faire apparaître le menu.
3. Menu : **France Travail**
   - `Configurer FT_CLIENT_ID / FT_CLIENT_SECRET`
   - `MAJ Travailleurs sociaux (<24h)`

L’import crée une **nouvelle feuille à chaque exécution**, nommée avec un timestamp (ex. `19-02-2026 10:42`).

---

## Google Sheet : structure attendue

Le script gère automatiquement :
- la création de la feuille d’exécution,
- l’écriture des en-têtes,
- le gel de la ligne 1,
- la validation de données sur la colonne `Status`.

Les statuts possibles sont définis dans `src/config.ts` (`STATUS_VALUES`).

Les URL des offres déjà importées sont stockées dans les **Notes** de la cellule `intitule`.

---

## Exclusions

Si un onglet `Exclusions` existe, il est lu ainsi :
- Colonne A : mots-clés à exclure si présents dans l’**intitulé**
- Colonne B : mots-clés à exclure si présents dans l’**entreprise**

Les exclusions sont comparées en *lowercase* (voir `normalizeKeyword_`).

---

## Développement

- Mode watch (re-compile en continu) :
  - `npm run watch`

- Ouvrir le projet Apps Script dans le navigateur :
  - `npm run open`

- Récupérer le code distant (attention, peut écraser local selon vos réglages clasp) :
  - `npm run pull`

---

## Structure du projet

- `src/` : source TypeScript (logique principale)
  - `main.ts` : `onOpen()` (menu)
  - `secrets.ts` : configuration + lecture des secrets (Script Properties)
  - `ftApi.ts` : token OAuth2 + fetch paginé
  - `jobs.ts` : job d’import « travailleurs sociaux <24h »
  - `sheet.ts` : création/formatage feuille, dédoublonnage, exclusions
  - `domain.ts` : règles métier (tri, normalisation contact, exclusions)
  - `config.ts` : constantes (colonnes, URLs, statuts, pagination)
  - `utils.ts` : helpers (format date, bool, liens)
- `dist/` : JS compilé poussé vers Apps Script
- `appsscript.json` : manifest Apps Script (runtime V8)
- `Code.js` : héritage historique (⚠️ ne plus modifier)

---

## Dépannage

### Le menu n’apparaît pas
- Vérifier que la fonction `onOpen()` est bien présente dans le code poussé.
- Recharger le Google Sheet.

### Erreur 401 / token
- Les tokens sont mis en cache (`CacheService`) ~50 minutes.
- En cas de 401, le script purge le cache et retante une fois.
- Vérifier que `FT_CLIENT_ID` / `FT_CLIENT_SECRET` sont corrects.

### Exécution via trigger (pas de popup)
Si les secrets ne sont pas définis et que le script s’exécute **sans UI** (trigger), la popup de configuration est impossible :
- lancer la configuration manuellement depuis le menu,
- ou pré-remplir les Script Properties.

### Quotas Apps Script
L’import peut consommer des quotas (UrlFetch / durée). Réduire :
- `FT_PAGE_SIZE` / `FT_MAX_PAGES` dans `src/config.ts`.

---

## Sécurité / secrets

- Ne versionnez pas de secrets en dur dans le dépôt.
- Préférez `Script Properties`.
- Si vous devez partager le script, documentez la procédure de configuration plutôt que d’embarquer des valeurs dans `FT_DEFAULTS`.

---

## Notes techniques

- Apps Script n’exécute pas des modules ES/Node de la même manière que Node.js : le runtime attend des **fonctions globales**. C’est pourquoi le code évite `export`/`import`.
- Le `tsconfig.json` est configuré pour compiler en JS compatible Apps Script (`module: "None"`, `outDir: dist`).
