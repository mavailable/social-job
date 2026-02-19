# Social Job – France Travail (Apps Script)

Migration vers **clasp + TypeScript**.

## Prérequis
- Node.js (>= 18 recommandé)
- Un projet Google Apps Script (Spreadsheet-bound ou standalone)

## Installation
1. Installer les dépendances:
   - `npm install`

2. Renseigner le Script ID dans `.clasp.json`:
   - Ouvrir Apps Script → Project Settings → Script ID
   - Remplacer `PASTE_SCRIPT_ID_HERE`

## Build
- `npm run build`

## Push vers Apps Script
- `npm run push`

## Notes
- Le code TypeScript est dans `src/`
- Le code compilé (envoyé par clasp) est dans `dist/`
- `appsscript.json` reste à la racine et est poussé aussi
