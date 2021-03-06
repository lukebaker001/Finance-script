var clientId = 'Q06gH4PJkQj0561JLFcvmS1inEyBbqgt3fwWFWrH5mQkRoYt5n';
var clientSecret = 'A21oPmSBzvp8lJ0aQooQLWKRoUpUVNioKzYXV7In';
var baseAuthURL = 'https://appcenter.intuit.com/connect/oauth2';
var tokenURL = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer';
var accountingSandboxURL = 'https://sandbox-quickbooks.api.intuit.com';
var accountingProductionURL = 'https://quickbooks.api.intuit.com'
var accountingURL = accountingProductionURL
var APIScope = 'com.intuit.quickbooks.accounting';
var realmID = 1419296350
var fullBaseURL = accountingURL+'/v3/company/'+realmID+'/'
var webhookToken = 'a52bfe88-8e0d-4fad-a441-fa53e03b51d4';
var webhook = "https://script.google.com/macros/s/AKfycbw9Yo3RdsvxiEpeycm2NUNnYvrWdQo1AQCjvmfnDpaKKBVjo56b/exec"
var redirectURI = 'https://script.google.com/macros/d/1hZikzgbbVMdAQqENs8I8b_J-pw19F1XmasGBjaas-C0TEHgu9QyDPU5I/usercallback';
var dateFormat = "YYYY-MM-dd"
var timeFormat = "HH:mm:ss"  ////QBO time format is weird (UTC time shown with non-UTC offset - see Build section)
var responseType = 'code';
var font = 'font-family: Verdana'
var fontSize = 'font-size: 20px'
var encode = "Basic "+Utilities.base64Encode(clientId + ":" + clientSecret)
var timezone = Session.getScriptTimeZone()

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Shared Studios')
  .addItem('Connect to Quickbooks', 'run')
  .addItem('Disconnect from Quickbooks', 'reset')
  .addSeparator()
  .addItem('Run custom query','customQuery')
  .addToUi();
} 

function postRequest(callURL,payload) { //need to add "if fail", get token from PropServ and try again
  var properties = PropertiesService.getScriptProperties
  var accessToken = properties.getProperty('accessToken')
  var encode = "Bearer "+accessToken
  var headers = {
    'Authorization': encode
  };

  var options = {
    method: 'POST',
    contentType: "application/json",
    headers: headers,
    muteHttpExceptions: true,
    payload : payload
  }
  var serverResponse = UrlFetchApp.fetch(callURL, options)
  return serverResponse
}

function getRequest(callURL) { //need to add "if fail", get token from PropServ and try again
  var properties = PropertiesService.getScriptProperties()
  var accessToken = properties.getProperty('accessToken')
  var encode = "Bearer "+accessToken
  var headers = {
    'Authorization': encode,
    'Accept': 'application/json'
  };

  var options = {
    method: 'GET',
    headers: headers,
    muteHttpExceptions: true,
  }
  var serverResponse = UrlFetchApp.fetch(callURL, options)
  return serverResponse
}

function singleRead(resourceName,entityID) {

  var callURL = fullBaseURL + resourceName + '/' +  entityID
  var sent = getRequest(callURL)
  return sent
}

function multiRead(queryBuilder) {
  var query = "query?query="+encodeURIComponent(queryBuilder)
  var callURL = fullBaseURL + query
  var sent = getRequest(callURL)
  return sent
}

function create(resourceName,details) {

  var callURL = fullBaseURL + resourceName
  var sent = postRequest(callURL,payload)
  return sent
}

function update(resourceName,details) {

  var payload = {
    sparse : "true"

  }
  var callURL = fullBaseURL + resourceName
  var sent = postRequest(callURL,payload)
  return sent
}

function deleteEntity(resourceName) {
  
  var deleteString ="?operation=delete"
  var callURL = fullBaseURL + resourceName + deleteString
  var sent = postRequest(callURL,payload)
  return sent
}

function batchRequest(){
  
}

function doPost(post) {
//  Logger.log("Post received")
//  var timezone = Session.getScriptTimeZone()
//  var spreadsheet = SpreadsheetApp.openById('1tZQfXgZ5eV2J6Fk7UCo3nK3vJ63ihdVnyhDBDZXbrMs');
//  var logSheet = spreadsheet.getSheetByName('Webhook Log');
//  var rawPost = JSON.stringify(post);
//  var eventArray = post.postData.contents.eventNotifications;
//  var eventCount = eventArray.length;
//  eventArray.forEach(function(event) {
//    entityArray = event.entities
//    entities.forEach
//      
//      
//  var inputArray = []
//  var now = Utilities.formatDate(new Date(),timezone,"YYYY-MM-dd HH:mm:ss zzzz")
//  inputArray.push(now);
//  inputArray.push(rawPost);
//  logSheet.appendRow(inputArray)
  var response = HtmlService.createHtmlOutput();
  return response;


}

function doGet(){
  Logger.log("doGet received")
  return HtmlService.createHtmlOutput('<b>GET request received successfully!</b>')
}

///////////////////////////////////////////////////////////CUSTOM QUERY///////////////////////////////////////////////////////CUSTOM QUERY
///////////////////////////////////////////////////////////CUSTOM QUERY///////////////////////////////////////////////////////CUSTOM QUERY
///////////////////////////////////////////////////////////CUSTOM QUERY///////////////////////////////////////////////////////CUSTOM QUERY

function customQuery(){
  var timezone = Session.getScriptTimeZone();
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('CustomRangeQuery');
  var queryRange = sheet.getRange('A1:J11');
  var queryValues = queryRange.getValues();
  var resourceType = queryValues[3][2];
  Logger.log("resourceType: "+resourceType)
  
  var conditions = [];
  var countOnly = queryValues[4][2];
  var conditionCount = 0;
  var holderCount = 0
  for (var i=0; i<5; i++) {
    var row = i+3;
    var holder = [];
    holderCount = 0
    if (queryValues[row][5] != '') {
      holder.push(queryValues[row][5]);
      holderCount++;
    }
    if (queryValues[row][6] != '') {
      holder.push(queryValues[row][6]);
      holderCount++;
    }
    if (queryValues[row][7] !== '') {
      holder.push(queryValues[row][7]);
      holderCount++;
    }
    if (holderCount == 3) {
      conditions.push(holder);
      conditionCount++
    }
  }
  Logger.log("conditions: "+conditions)
  
  var attributes = [];
  for (var i=0; i<9; i++) {
    var col = i+1;
    attributes.push(queryValues[9][col])
  }
  Logger.log("attributes: "+attributes)
  
  if (countOnly) var queryBuilder = 'SELECT COUNT(*)FROM '+resourceType
  else var queryBuilder = 'SELECT * FROM '+resourceType
    
  var join = " WHERE "
  var and = " AND "
  for (var i=0; i<conditions.length; i++) {
    var attribute = conditions[i][0]
    var operator = conditions[i][1]
    var condition = conditions[i][2]
    queryBuilder += join
    queryBuilder += attribute+" "
    queryBuilder += operator
    if (operator == 'IN') {
      queryBuilder += "("+condition+")";
    }
    else if (condition === true || condition === false) { 
      queryBuilder += condition;
    }
    else if (condition instanceof Date && !isNaN(condition.valueOf())) {
      var dateBuilder = Utilities.formatDate(condition,timezone,dateFormat)
      Logger.log("dateBuilder: "+dateBuilder)
      var timeBuilder = Utilities.formatDate(condition,timezone,timeFormat)
      Logger.log("timeBuilder: "+timeBuilder)
      var timezoneBuilder = Utilities.formatDate(condition,timezone,"X")
      Logger.log("timezoneBuilder: "+timezoneBuilder)
      var timeDateCondition = dateBuilder+"T"+timeBuilder+timezoneBuilder+":00"
      Logger.log("timeDateCondition: "+timeDateCondition)
      queryBuilder += "'"+timeDateCondition+"'";
    }
    else {
      queryBuilder += "'"+condition+"'";
    }
    join = " AND "
  }
  
  var orderByAttribute = queryValues[6][2];
  var orderByOrder = queryValues[6][3];
  if (orderByAttribute && orderByOrder) {
    queryBuilder += " ORDERBY "+orderByAttribute+" "+orderByOrder
  }
  
  var maxResults = queryValues[5][2];
  Logger.log("maxResults: "+maxResults)
  if (maxResults > 0 && maxResults < 1001) queryBuilder += " MAXRESULTS "+maxResults

  Logger.log("queryBuilder: "+queryBuilder)

  var response = multiRead(queryBuilder) /////////Send request
  
  Logger.log(response)
  var responseObject = JSON.parse(response);
  
  var lastRow = sheet.getLastRow();
  sheet.getRange(11,2,lastRow-9,9).clearContent()
  
  if (responseObject.Fault) {
    var errorResponseRange = sheet.getRange(11,2,4,1)
    var errorResponseValues = errorResponseRange.getValues()
    errorResponseValues[0][0] = "Error"
    errorResponseValues[1][0] = "For query <"+queryBuilder+">"
    errorResponseValues[2][0] = responseObject.Fault.Error[0].Message
    errorResponseValues[3][0] = responseObject.Fault.Error[0].Detail
    errorResponseRange.setValues(errorResponseValues)
    return
  }
  
  if (responseObject.fault) {
    var errorResponseRange = sheet.getRange(11,2,3,1)
    var errorResponseValues = errorResponseRange.getValues()
    errorResponseValues[0][0] = "Error"
    errorResponseValues[1][0] = responseObject.fault.error[0].Message
    errorResponseValues[2][0] = responseObject.Fault.Error[0].Detail
    errorResponseRange.setValues(errorResponseValues)
    return
  }
  
  if (countOnly) {
    var responseCount = responseObject.QueryResponse.totalCount
    Logger.log("responseCount: "+responseCount)
    sheet.getRange('B11').setValue('Total results: '+responseCount)
    return;
  }
  
  if (responseObject.QueryResponse[resourceType]) {
    var responseLength = responseObject.QueryResponse[resourceType].length
    Logger.log("responseLength: "+responseLength)
    var responseRange = sheet.getRange(11,2,responseLength,9)
    var responseValues = responseRange.getValues()
    var objectArray = responseObject.QueryResponse[resourceType]
    Logger.log("objectArray: "+JSON.stringify(objectArray))
    Logger.log("resourceType: "+resourceType)
    for (var i=0; i<responseLength; i++) {
      for (var j=0; j<9; j++) {
        var arrayRow = objectArray[i]
        var columnAttribute = attributes[j]
        var typeOfColumn = typeof columnAttribute
        if (!columnAttribute) continue;
        else if (columnAttribute.indexOf('.') > -1) {
          var index = columnAttribute.indexOf('.')
          var columnLength = columnAttribute.length
          var firstColumnAttribute = columnAttribute.slice(0,index)
          var secondColumnAttribute = columnAttribute.slice(index+1,columnLength)
          Logger.log("firstColumnAttribute: "+firstColumnAttribute)
          Logger.log("secondColumnAttribute: "+secondColumnAttribute)
          var firstArray = arrayRow[firstColumnAttribute]
          responseValues[i][j]=firstArray[secondColumnAttribute]
        }
        else {
          var value = arrayRow[columnAttribute];
          if (value == undefined) value = "-";
          responseValues[i][j]=value;
        }
      } //end for
    }
    responseRange.setValues(responseValues) 
  }
  else {
    sheet.getRange('B11').setValue('Query returned 0 responses')
    return;
  }
    
  //[10][1]
  
  
}

///////////////////////////////////////////////////////////IMPORT PURCHASES///////////////////////////////////////////////////////IMPORT PURCHASES
///////////////////////////////////////////////////////////IMPORT PURCHASES///////////////////////////////////////////////////////IMPORT PURCHASES
///////////////////////////////////////////////////////////IMPORT PURCHASES///////////////////////////////////////////////////////IMPORT PURCHASES


function importAllPurchases(){
  var tableStartRow = 74;
  var timezone = Session.getScriptTimeZone();
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Purchases');
  
  var queryBuilder = "SELECT * FROM Purchase ORDERBY TxnDate DESC MAXRESULTS 1000"
  
  var response = multiRead(queryBuilder) /////////Send request
  
  //Logger.log("response: "+response)
  var responseObject = JSON.parse(response);
  
  var lastRow = sheet.getLastRow();
  
  if (responseObject.Fault) {
    var errorResponseRange = sheet.getRange(tableStartRow,2,4,1)
    var errorResponseValues = errorResponseRange.getValues()
    errorResponseValues[0][0] = "Error"
    errorResponseValues[1][0] = "For query <"+queryBuilder+">"
    errorResponseValues[2][0] = responseObject.Fault.Error[0].Message
    errorResponseValues[3][0] = responseObject.Fault.Error[0].Detail
    errorResponseRange.setValues(errorResponseValues)
    return
  }
  
  if (responseObject.fault) {
    var errorResponseRange = sheet.getRange(tableStartRow,2,3,1)
    var errorResponseValues = errorResponseRange.getValues()
    errorResponseValues[0][0] = "Error"
    errorResponseValues[1][0] = responseObject.fault.error[0].Message
    errorResponseValues[2][0] = responseObject.Fault.Error[0].Detail
    errorResponseRange.setValues(errorResponseValues)
    return
  }
  
  if (responseObject.QueryResponse.Purchase) {
    var responseLength = responseObject.QueryResponse.Purchase.length
    Logger.log("responseLength: "+responseLength)
    var responseValues = []
    var purchaseArray = responseObject.QueryResponse.Purchase
    Logger.log("purchaseArray: "+JSON.stringify(purchaseArray))
    var lineNumber = 0;
    var lineArray = []
    purchaseArray.forEach(function(e) {
      var unique = 0
      var line = e.Line
      line.forEach(function(l){
        lineArray = []
        if (unique == 0) lineArray.push('Unique payment')
        else lineArray.push('Line details')
        lineArray.push(e.Id)
        var txnDate = e.TxnDate
        lineArray.push(txnDate)
        lineArray.push(Utilities.formatDate(new Date(txnDate),timezone,"YYYY-MM"))
        var account = e.AccountRef
        lineArray.push(account.name)
        lineArray.push(e.PrivateNote)
        lineArray.push(l.Id)
        lineArray.push(l.Description)
        var accountBasedDetail = l.AccountBasedExpenseLineDetail
        //if (l.AccountRef) {
        var lineAccount = accountBasedDetail.AccountRef
        var lineAccountName = lineAccount.name
        if (lineAccountName.indexOf(":") != -1) {
          var index = lineAccountName.indexOf(":")
          var parentAccount = lineAccountName.slice(0,index)
          lineArray.push (parentAccount)
        }
        else {
          lineArray.push (lineAccountName)
        }
        lineArray.push (lineAccountName)
        if (accountBasedDetail.ClassRef) {
          var class = accountBasedDetail.ClassRef
          lineArray.push(class.value)
          lineArray.push(class.name)
        }
        else {
          lineArray.push("<no class assigned>")
          lineArray.push("<no class assigned>")
        }
        lineArray.push(l.Amount)
        lineArray.push(e.PaymentType)
        lineArray.push(e.TotalAmt)
        var currency = e.CurrencyRef
        lineArray.push(currency.value)
        responseValues.push(lineArray)
        unique = 1
        lineNumber++;
    
          
      })
    })
    Logger.log("responseValues: "+ responseValues)
    var responseRange = sheet.getRange(tableStartRow,2,lineNumber,16)
    Logger.log("setting values")
    responseRange.setValues(responseValues) 
  }
  else {
    sheet.getRange(tableStartRow,2,1,1).setValue('Query returned 0 responses')
    return;
  }
  
}

function test() {
  var post = {
    "parameter":{},
    "contextPath":"",
    "contentLength":176,
    "queryString":"",
    "parameters":{},
    "postData":
    {
      "type":"application/json",
      "length":176,
      "contents": {
        "eventNotifications": [{
          "realmId":"1419296350",
          "dataChangeEvent": {
            "entities":
            [
              {
                "name":"Purchase",
                "id":"9189",
                "operation":"Update",
                "lastUpdated":"2018-09-26T17:47:53.000Z"
              }
            ]
          }
        }]
      },
      "name":"postData"
    }
  }
  }