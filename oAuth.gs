
function run() {
  getService().reset(); 
  var service = getService();
  var authorizationUrl = service.getAuthorizationUrl()
  Logger.log(authorizationUrl);
  var img = "https://drive.google.com/uc?export=download&id=103lt6K977dnQpvPNzVeaIKZ_2Optn7SW"
  var html = '<html><body><a href="'+authorizationUrl+'" target="blank" onclick="google.script.host.close()"><img src="'+img+'" title="Connect to Quickbooks" width=351 height=61></a></body></html>'
  var interface = HtmlService.createHtmlOutput(html).setHeight(80).setWidth(390)
  SpreadsheetApp.getUi().showModelessDialog(interface,"Quickbooks must be authorized to continue (sandbox)");
  }

//Configures the service.
function getService() {
  return OAuth2.createService('Quickbooks Sandbox')
      .setAuthorizationBaseUrl(baseAuthURL)
      .setTokenUrl(tokenURL)
      .setClientId(clientId)
      .setClientSecret(clientSecret)
      .setScope(APIScope)
      .setCallbackFunction('authCallback')
      .setParam('response_type', responseType)
      .setParam('state', getStateToken('authCallback')) // function to generate the state token on the fly
      .setPropertyStore(PropertiesService.getUserProperties());
}

//Handles the OAuth callback
function authCallback(request) {
  Logger.log("Starting function authCallback")
  var encode = "Basic "+Utilities.base64Encode(clientId + ":" + clientSecret)
  Logger.log(encode);
  var authCode = request.parameter.code
    
  var headers = {
    'Authorization': encode
  };
  var options = {
    method: "POST",
    contentType: "application/x-www-form-urlencoded",
    headers: headers,
    muteHttpExceptions: true,
    payload : {
      'grant_type':'authorization_code',
      'code': authCode,
      'redirect_uri' : redirectURI
    }
  }
  
  var response = UrlFetchApp.fetch(tokenURL, options)
  var responseObject = JSON.parse(response)
  var accessToken = responseObject.access_token
  var refreshToken = responseObject.refresh_token
  var realmID = request.parameter.realmId
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty('authCode', authCode)
  properties.setProperty('realmID', realmID)
  properties.setProperty('accessToken', accessToken)
  properties.setProperty('refreshToken', refreshToken)
  if (responseObject.access_token) {
    var responseText = "Quickbooks is now connected"
    var img = 'https://www.j2store.org/images/extensions/apps/apps_preview_image/quickbooks_online_preview.png'
    var width = 240
    var height = 105
    var display = 'Access token:'
    var secondDisplay = accessToken
    var thirdDisplay = 'Refresh token:'
    var fourthDisplay = refreshToken
    }
  else {
    var responseText = 'An error has occurred. Please contact your system administrator'
    var img = 'https://st2.depositphotos.com/2274151/6347/i/450/depositphotos_63479487-stock-photo-epic-fail-red-grunge-seal.jpg'
    var width = 112
    var height = 91
    var display = ''
    var secondDisplay = response
    var thirdDisplay = ''
    var fourthDisplay = ''
    }
  var htmlResponse = '<html><center><img src="'+img+'" width="'+width+'" height="'+height+'"><body><p style="'+font+';"><b>'+responseText+'</p><p></p><p style="'+font+';">'+display+'</p></b><p>'+secondDisplay+'</p><p style="'+font+';"><b>'+thirdDisplay+'</b></p><p>'+fourthDisplay+'</p></body></center></html>'
  return HtmlService.createHtmlOutput(htmlResponse);
}

//Refresh access token
function refreshAccess() {
  Logger.log("Starting function refreshAccess")
  var properties = PropertiesService.getScriptProperties();
  var accessToken = properties.getProperty('accessToken')
  var refreshToken = properties.getProperty('refreshToken')
  
  var encode = "Basic "+Utilities.base64Encode(clientId + ":" + clientSecret)
  var headers = {
    'Authorization': encode
  };
  var options = {
    method: "POST",
    contentType: "application/x-www-form-urlencoded",
    headers: headers,
    muteHttpExceptions: true,
    payload : {
      'grant_type':'refresh_token',
      'refresh_token': refreshToken,
    }
  }
  var response = UrlFetchApp.fetch(tokenURL, options)
    
  var responseObject = JSON.parse(response)
  var accessToken = responseObject.access_token
  var refreshToken = responseObject.refresh_token
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty('accessToken', accessToken)
  properties.setProperty('refreshToken', refreshToken)
  
  Logger.log("response: "+response)

}
  
 
//Generate a State Token
function getStateToken(callbackFunction){
 var stateToken = ScriptApp.newStateToken()
     .withMethod(callbackFunction)
     .withTimeout(120)
     .createToken();
 return stateToken;
}

//Logs the redirect URI. Run this function to get the REDIRECT_URI to be mentioned at the top of this script.
function logRedirectUri() {
  Logger.log(getService().getRedirectUri());
}

function reset() {
  getService().reset();
}
