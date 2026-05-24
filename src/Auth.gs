function getHealthService() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty(HEALTH_OAUTH_CLIENT_ID_KEY);
  const clientSecret = props.getProperty(HEALTH_OAUTH_CLIENT_SECRET_KEY);
  if (!clientId || !clientSecret) {
    throw new Error(
      'Health OAuth not configured. Set Script Properties ' +
      HEALTH_OAUTH_CLIENT_ID_KEY + ' and ' + HEALTH_OAUTH_CLIENT_SECRET_KEY +
      ', then run "Authorize Health API" from the Sync menu.'
    );
  }
  return OAuth2.createService(HEALTH_SERVICE_NAME)
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/v2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('healthAuthCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setCache(CacheService.getUserCache())
    .setLock(LockService.getUserLock())
    .setScope(HEALTH_OAUTH_SCOPES)
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent');
}

function healthAuthCallback(request) {
  const service = getHealthService();
  const authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput(
      '<p>Google Health authorization succeeded. You can close this tab and return to the spreadsheet.</p>'
    );
  }
  return HtmlService.createHtmlOutput(
    '<p>Google Health authorization was denied. You can close this tab and try again from the Sync menu.</p>'
  );
}

function getHealthAccessToken_() {
  const service = getHealthService();
  if (!service.hasAccess()) {
    const authUrl = service.getAuthorizationUrl();
    throw new Error(
      'Health API not authorized. Open this URL in a browser signed into the target Google account: ' + authUrl
    );
  }
  return service.getAccessToken();
}

function resetHealthAuth() {
  getHealthService().reset();
}
