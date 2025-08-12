const AUTH_BASE_URL = 'https://auth.truelayer.com';
const API_BASE_URL = 'https://api.truelayer.com';

function fetchDeployments() {
  var scriptId = ScriptApp.getScriptId();
  var token = ScriptApp.getOAuthToken();
  
  var res = UrlFetchApp.fetch(`https://script.googleapis.com/v1/projects/${scriptId}/deployments`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  
  return JSON.parse(res.getContentText()).deployments;
}

function getPublicWebAppUrl() {
  var deployments = fetchDeployments();
  var actual = null;
  for (var d of deployments) {
    var currentVersionNumber = d.deploymentConfig.versionNumber || 0;
    var webAppEntryPoint = d.entryPoints && d.entryPoints.find(e => e.entryPointType === "WEB_APP");  
    var currentUrl = webAppEntryPoint && webAppEntryPoint.webApp && webAppEntryPoint.webApp.url;
    if (!actual || (currentUrl && actual.versionNumber < currentVersionNumber)) {
      actual = {
        versionNumber: currentVersionNumber,
        url: currentUrl
      };
    }
  }
  return actual.url;
}

// ===== CONFIG SHEET HELPERS =====
function getConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('_config');
  if (!sheet) {
    sheet = ss.insertSheet('_config');
    sheet.hideSheet();
    sheet.appendRow(['key', 'value']);
  }
  return sheet;
}

function setConfigValue(key, value) {
  var sheet = getConfigSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function getConfigValue(key) {
  var sheet = getConfigSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

function deleteConfigValue(key) {
  var sheet = getConfigSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

function formatConfigKey(userEmail, providerId, key) {
  return userEmail + '_' + providerId + '_' + key;
}

function setUserConfigValue(userEmail, providerId, key, value) {
  var key = formatConfigKey(userEmail, providerId, key);
  return setConfigValue(key, value);
}

function getUserConfigValue(userEmail, providerId, key) {
  var key = formatConfigKey(userEmail, providerId, key);
  return getConfigValue(key);
}

function deleteUserConfigValue(userEmail, providerId, key) {
  var key = formatConfigKey(userEmail, providerId, key);
  return deleteConfigValue(key);
}

// ===== MENU =====
function onOpen() {
  updateMenu();
  ensureHourlyTrigger();
}

function updateMenu() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Bank Connector');
  
  var clientId = getConfigValue('CLIENT_ID');
  var clientSecret = getConfigValue('CLIENT_SECRET');
  
  var userEmail = Session.getActiveUser().getEmail();
  var accessToken = getConfigValue(userEmail + '_ACCESS_TOKEN');

  menu.addItem('Settings', 'showSettingsSidebar');
  if (clientId && clientSecret) {
    menu.addSeparator();
    menu.addItem('Connect Bank', 'connectBank');
    menu.addItem('Refresh', 'refresh');
  }
  menu.addToUi();
}

// ===== SETTINGS =====
function showSettingsSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('settings')
    .setTitle('Bank Connector Settings');
  SpreadsheetApp.getUi().showSidebar(html);
}

function saveSettings(clientId, clientSecret) {
  setConfigValue('CLIENT_ID', clientId);
  setConfigValue('CLIENT_SECRET', clientSecret);
  SpreadsheetApp.getActiveSpreadsheet().toast('Settings saved');
  updateMenu();
}

function getSettings() {
  return {
    clientId: getConfigValue('CLIENT_ID') || '',
    clientSecret: getConfigValue('CLIENT_SECRET') || '',
    redirectUri: getPublicWebAppUrl()
  };
}

// ===== BANK CONNECTION =====
function connectBank() {
  var clientId = getConfigValue('CLIENT_ID');
  var clientSecret = getConfigValue('CLIENT_SECRET');

  if (!clientId || !clientSecret) {
    SpreadsheetApp.getUi().alert('Please set Client ID and Secret first.');
    return;
  }

  var redirectUri = getPublicWebAppUrl();
  var authUrl = AUTH_BASE_URL + '/?response_type=code'
    + '&client_id=' + encodeURIComponent(clientId)
    + '&scope=' + encodeURIComponent('info accounts transactions balance cards offline_access')
    + '&redirect_uri=' + encodeURIComponent(redirectUri)
    + '&providers=' + encodeURIComponent('uk-ob-all uk-oauth-all se-ob-all at-ob-all be-ob-all ee-ob-all fi-ob-all fr-ob-all de-ob-all ie-ob-all it-ob-all lt-ob-all nl-ob-all pl-ob-all pt-ob-all es-ob-all');

  var qrUrl = "https://quickchart.io/chart?cht=qr&chs=300x300&chl=" + encodeURIComponent(authUrl);

  var html = HtmlService.createTemplateFromFile('connectDialog');
  html.data = {
    qrUrl,
    authUrl
  };
  
  var modal = html.evaluate().setWidth(350).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(modal, 'Connect Bank');
}

// OAuth callback
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.code) {
    return HtmlService.createHtmlOutput("error: no code provided.");
  }

  var code = e.parameter.code;
  var userEmail = Session.getActiveUser().getEmail();
  var clientId = getConfigValue('CLIENT_ID');
  var clientSecret = getConfigValue('CLIENT_SECRET');
  var tokenData = fetchToken(clientId, clientSecret, code);
  var metadata = fetchMetadata(tokenData.access_token);

  var providerId = metadata.provider.provider_id;

  setUserConfigValue(userEmail, providerId, 'ACCESS_TOKEN', tokenData.access_token);
  setUserConfigValue(userEmail, providerId, 'REFRESH_TOKEN', tokenData.refresh_token);
  
  // Import accounts using the dedicated function
  importAccounts(userEmail, providerId);

  return HtmlService.createHtmlOutput("Bank connected. You can close this window.");
}

function refresh() {
  var accounts = readSheet('Accounts');
  accounts.forEach(account => {
    refreshTransactions(account.account_owner, account.provider_id, account.account_id, account.source_type);
  });
}

function readSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  return data.filter((_, index) => index > 0).map(row => {
    var parsedRow = {};
    for (var j = 0; j < headers.length; j++) {
      parsedRow[headers[j]] = row[j];
    }
    return parsedRow;
  });
}

function refreshTransactions(userEmail, providerId, accountId, sourceType) {
  const today = new Date();
  const year = today.getFullYear().toString();
  const past90 = new Date();
  past90.setDate(past90.getDate() - 89);
  const startDate = Utilities.formatDate(past90, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const endDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

  var accessToken = getValidAccessToken(userEmail, providerId);
  var fetcher = sourceType == 'ACCOUNT' ? fetchTransactions : fetchCardTransactions;

  var transactions = fetcher(accessToken, accountId, startDate, endDate).map(tx => {
    tx.account_id = accountId;
    return tx;
  });

  updateSheet(year, transactions);
}

function updateSheet(year, transactions) {
  if (transactions.length <= 0) {
    return;
  }
  
  const HEADERS = [
    "date",
    "description", 
    "amount",
    "currency",
    "account_id",
    "category",
    "external_id",
    "notes"
  ];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(year);
  if (!sheet) {
    sheet = ss.insertSheet(year);
    sheet.appendRow(HEADERS);
  }

  currentTransactions = readSheet(year);

  excludeIds = currentTransactions.map(tx => tx.external_id);

  allTransactions = transactions.filter(tx => !excludeIds.includes(tx.normalised_provider_transaction_id || tx.provider_transaction_id || tx.transaction_id)).map(tx => [
    tx.timestamp.split("T")[0], // ISO date format
    tx.description,
    tx.amount, // Keep original sign (negative for expenses, positive for income)
    tx.currency, // ISO currency code
    tx.account_id, // Readable account name
    tx.transaction_category, // Category (can be filled manually or via rules in Firefly III)
    tx.normalised_provider_transaction_id || tx.provider_transaction_id || tx.transaction_id, // External ID for duplicate detection
    JSON.stringify(tx) // Additional info in notes
  ]);

  sheet.getRange(sheet.getLastRow() + 1, 1, allTransactions.length, HEADERS.length).setValues(allTransactions);
}

function updateAccountsSheet(accounts, cards, consentExpiresAt) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Accounts');
  const HEADERS = [
    "account_id",
    "display_name",
    "account_type",
    "account_subtype",
    "currency",
    "account_number",
    "iban",
    "swift_bic",
    "sort_code",
    "provider_id",
    "provider_name",
    "last_updated",
    "source_type",
    "account_owner",
    "consent_expires_at",
    "notes"
  ];
  if (!sheet) {
    sheet = ss.insertSheet('Accounts');
    // Create headers for accounts sheet
    sheet.appendRow(HEADERS);
  }
  
  var allAccountData = [];
  var currentTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var userEmail = Session.getActiveUser().getEmail();
  
  // Process regular accounts
  if (accounts && accounts.length > 0) {
    allAccountData = allAccountData.concat(accounts.map(account => {
      // Extract account details from account_number object
      var accountNumberObj = account.account_number || {};
      var accountNumber = accountNumberObj.number || '';
      var iban = accountNumberObj.iban || '';
      var swiftBic = accountNumberObj.swift_bic || '';
      var sortCode = accountNumberObj.sort_code || '';
      
      return [
        account.account_id || '',
        account.display_name || account.account_id || '',
        account.account_type || 'UNKNOWN',
        account.account_subtype || '',
        account.currency || 'EUR',
        accountNumber,
        iban,
        swiftBic,
        sortCode,
        account.provider && account.provider.provider_id || '',
        account.provider && account.provider.display_name || '',
        currentTime,
        'ACCOUNT',
        userEmail,
        consentExpiresAt,
        JSON.stringify(account)
      ];
    }));
  }
  
  // Process card accounts
  if (cards && cards.length > 0) {
    allAccountData = allAccountData.concat(cards.map(card => {
      return [
        card.account_id || '',
        card.display_name || card.account_id || '',
        'CREDIT_CARD',
        card.card_type || '',
        card.currency || 'EUR',
        (card.partial_card_number || '****') + '************',
        '', // No IBAN for cards
        '', // No SWIFT BIC for cards
        '', // No sort code for cards
        card.provider && card.provider.provider_id || '',
        card.provider && card.provider.display_name || '',
        currentTime,
        'CARD',
        userEmail,
        consentExpiresAt,
        JSON.stringify(card)
      ];
    }));
  }
  
  // Smart update logic: update existing accounts, add new ones, never delete
  var lastRow = sheet.getLastRow();
  var existingData = {};
  var existingRowMap = {};
  
  // Read existing data to build account_id lookup map
  if (lastRow > 1) {
    var existingRange = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
    var existingValues = existingRange.getValues();
    
    for (var i = 0; i < existingValues.length; i++) {
      var accountId = existingValues[i][0]; // account_id is in column A (index 0)
      if (accountId) {
        existingData[accountId] = existingValues[i];
        existingRowMap[accountId] = i + 2; // +2 because sheet is 1-indexed and we start from row 2
      }
    }
  }
  
  var newAccounts = [];

  // Process each account in allAccountData
  for (var j = 0; j < allAccountData.length; j++) {
    var accountData = allAccountData[j];
    var accountId = accountData[0];
    
    if (existingData[accountId]) {
      // Account exists - update the existing row
      var rowNumber = existingRowMap[accountId];
      sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([accountData]);
    } else {
      newAccounts.push(accountData);
    }
  }
  
  // Add new accounts at the end
  if (newAccounts.length > 0) {
    var nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, newAccounts.length, HEADERS.length).setValues(newAccounts);
  }
  
  // Auto-resize columns for better readability
  if (allAccountData.length > 0) {
    sheet.autoResizeColumns(1, HEADERS.length);
  }
}

function importAccounts(userEmail, providerId) {
  try {
    var accessToken = getValidAccessToken(userEmail, providerId);
    var metadata = fetchMetadata(accessToken);
    var accounts = fetchAccounts(accessToken);
    var cards = fetchCards(accessToken);
    
    if (!accounts && !cards) {
      Logger.log('No account data available to import');
      return;
    }
    
    // Update the accounts sheet with the fresh data
    updateAccountsSheet(accounts, cards, metadata.consent_expires_at);
  } catch (error) {
    Logger.log('Error importing accounts: ' + error.toString());
  }
}

// APIs
function fetchToken(clientId, clientSecret, code) {
  var redirectUri = getPublicWebAppUrl();

  var tokenResponse = UrlFetchApp.fetch(AUTH_BASE_URL + '/connect/token', {
      method: 'post',
      payload: {
        grant_type: 'authorization_code',
        client_id: clientId,
        client_secret: clientSecret,
        redirect_uri: redirectUri,
        code: code
      }
    });

    return JSON.parse(tokenResponse.getContentText());
}

function fetchMetadata(accessToken) {
  var res = UrlFetchApp.fetch(API_BASE_URL + '/data/v1/me', {
    headers: { Authorization: 'Bearer ' + accessToken }
  });
  var json = JSON.parse(res.getContentText());
  if (json.results && json.results[0]) {
    return json.results[0];
  }
  return null;
}

function fetchAccounts(accessToken) {
  try {
    var res = UrlFetchApp.fetch(API_BASE_URL + '/data/v1/accounts', {
      headers: { Authorization: 'Bearer ' + accessToken }
    });
    var json = JSON.parse(res.getContentText());
    return json.results || [];
  } catch {
    return [];
  }
}

function fetchCards(accessToken) {
  try {
    var res = UrlFetchApp.fetch(API_BASE_URL + '/data/v1/cards', {
      headers: { Authorization: 'Bearer ' + accessToken },
    });
    var json = JSON.parse(res.getContentText());
    return json.results || [];
  } catch {
    return [];
  }
  
}

function fetchTransactions(accessToken, accountId, fromDate, toDate) {
  var res = UrlFetchApp.fetch(API_BASE_URL + `/data/v1/accounts/${accountId}/transactions?from=${fromDate}&to=${toDate}`, {
    headers: {
      Authorization: 'Bearer ' + accessToken,
      Accept: 'application/json'
    }
  });
  var json = JSON.parse(res.getContentText());
  return json.results || [];
}

function fetchCardTransactions(accessToken, accountId, fromDate, toDate) {
  var res = UrlFetchApp.fetch(API_BASE_URL + `/data/v1/cards/${accountId}/transactions?from=${fromDate}&to=${toDate}`, {
    headers: {
      Authorization: 'Bearer ' + accessToken,
      Accept: 'application/json'
    }
  });
  var json = JSON.parse(res.getContentText());
  return json.results || [];
}

function refreshAccessToken(clientId, clientSecret, refreshToken) {
  var tokenResponse = UrlFetchApp.fetch(AUTH_BASE_URL + '/connect/token', {
    method: 'post',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret
    }
  });
  return JSON.parse(tokenResponse.getContentText());
}

function decode_jwt(value) {
  function decode(b64) {
    var buf = Utilities.base64Decode(b64);
    var data = Utilities.newBlob(buf).getDataAsString();
    return JSON.parse(data);
  }
  var parts = value.split(".");
  parts[0] = decode(parts[0]);
  parts[1] = decode(parts[1]);
  return parts;
}

function getValidAccessToken(userEmail, providerId) {
  var clientId = getConfigValue('CLIENT_ID');
  var clientSecret = getConfigValue('CLIENT_SECRET');
  var accessToken = getUserConfigValue(userEmail, providerId, 'ACCESS_TOKEN');
  var refreshToken = getUserConfigValue(userEmail, providerId, 'REFRESH_TOKEN');

  if (!accessToken || !refreshToken) {
    throw new Error('Missing tokens, please reconnect your bank.');
  }

  var jwt = decode_jwt(accessToken);

  var exp = new Date(jwt[1].exp * 1000);
  var now = new Date();
  if (now < exp) {
    return accessToken;
  }

  var newTokens = refreshAccessToken(clientId, clientSecret, refreshToken);

  setUserConfigValue(userEmail, providerId, 'ACCESS_TOKEN', newTokens.access_token);
  setUserConfigValue(userEmail, providerId, 'REFRESH_TOKEN', newTokens.refresh_token);
  return newTokens.access_token;
}



function ensureHourlyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = triggers.some(function (t) {
    return t.getHandlerFunction() === "refresh" &&
           t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK;
  });

  if (!triggerExists) {
    ScriptApp.newTrigger("refresh")
      .timeBased()
      .everyHours(1)
      .create();
    Logger.log("Hourly trigger created.");
  } else {
    Logger.log("Hourly trigger already exists.");
  }
}
