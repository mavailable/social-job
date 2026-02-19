// Apps Script runtime expects globals; avoid `export`/`import` syntax.

function ftEnsureSecrets_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FT_CLIENT_ID');
  const secret = props.getProperty('FT_CLIENT_SECRET');

  if (id && secret) return;

  if (
    FT_DEFAULTS.FT_CLIENT_ID &&
    FT_DEFAULTS.FT_CLIENT_SECRET &&
    FT_DEFAULTS.FT_CLIENT_ID !== 'TON_CLIENT_ID_ICI' &&
    FT_DEFAULTS.FT_CLIENT_SECRET !== 'TON_CLIENT_SECRET_ICI'
  ) {
    props.setProperties({
      FT_CLIENT_ID: FT_DEFAULTS.FT_CLIENT_ID,
      FT_CLIENT_SECRET: FT_DEFAULTS.FT_CLIENT_SECRET
    });
    Logger.log('✅ Secrets France Travail initialisés depuis FT_DEFAULTS');
    return;
  }

  ftPromptAndSaveSecrets_();
}

function ftPromptAndSaveSecrets_() {
  let ui: GoogleAppsScript.Base.Ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    const msg =
      '❌ FT_CLIENT_ID / FT_CLIENT_SECRET manquants. ' +
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
