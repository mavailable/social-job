// Apps Script runtime expects globals; avoid `export`/`import` syntax.

type FtSearchParams = Record<string, string | number | boolean>;

function ftGetToken_(): string {
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

  const data = JSON.parse(res.getContentText()) as { access_token: string };
  CacheService.getScriptCache().put('FT_ACCESS_TOKEN', data.access_token, 50 * 60);
  return data.access_token;
}

function ftFetchAllResults_(baseParams: FtSearchParams): any[] {
  let token = ftGetToken_();
  const resultats: any[] = [];
  let retriedAfter401 = false;

  for (let page = 0; page < FT_MAX_PAGES; page++) {
    const start = page * FT_PAGE_SIZE;
    const end = start + FT_PAGE_SIZE - 1;

    const params = { ...baseParams, range: `${start}-${end}` };
    const url =
      FT_SEARCH_URL +
      '?' +
      Object.entries(params)
        .map(([k, v]) => `${k}=${encodeURIComponent(String(v))}`)
        .join('&');

    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const body = res.getContentText() || '';

    if (code === 401 && !retriedAfter401) {
      retriedAfter401 = true;
      CacheService.getScriptCache().remove('FT_ACCESS_TOKEN');
      token = ftGetToken_();
      page--;
      continue;
    }

    if (code >= 300) {
      throw new Error(`❌ Search error (${code}): ${body}`);
    }

    const json = JSON.parse(body);
    const offres = (json?.resultats || []) as any[];
    resultats.push(...offres);

    // Less than a full page => last page
    if (offres.length < FT_PAGE_SIZE) break;
  }

  return resultats;
}
