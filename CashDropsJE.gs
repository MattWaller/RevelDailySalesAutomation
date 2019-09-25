function CashDropsJE() {
  var app = SpreadsheetApp;
  var as = app.getActiveSpreadsheet().getSheetByName("CashDropJE");
  var date = as.getRange(1, 2).getValue();
  var rows = as.getLastRow()-2;
  var i;
  var body = [];
  var header = '{"TxnDate": '+ JSON.stringify(date) + ' ,"Line":'
  var footer = ', "TxnTaxDetail": {}}'

  for (i = 0; i< rows; i++){
    var row = i + 2;
    var amount = as.getRange(row, 9).getValue();
    var pt = as.getRange(row, 8).getValue();
    var name = as.getRange(row, 3).getValue();
    var val = as.getRange(row, 10).getValue();
    var desc = as.getRange(row, 6).getValue();
    body.push({
              "Id": JSON.stringify(i),
              "Description": desc,
              "Amount": amount,
              "DetailType": "JournalEntryLineDetail",
              "JournalEntryLineDetail": {
                "PostingType": pt,
                "AccountRef": {
                  "value": JSON.stringify(val),
                  "name": name
                }
              }
            });
  }
 var b  = JSON.stringify(body)
 var payload = header+b+footer;
  Logger.log(payload)
   
  var a = SpreadsheetApp;
  var aS = a.getActiveSpreadsheet().getSheetByName("Refresh_Key");   
  var refresh_k = aS.getRange(1, 1).getValue();   
  
  var token_url = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer";
  var token_options = 
      {
        "headers": {
          "Content-Type": "application/x-www-form-urlencoded",
          "Accept": "application/json",
          "Authorization": "Basic AUTH KEY"
        },
        "payload": {
          "grant_type" : "refresh_token",
          "refresh_token" : refresh_k,
        }
      };
  
  var token_response = UrlFetchApp.fetch(token_url,token_options); 
  var access_token = JSON.parse(token_response);
  var token_key = access_token.access_token;
  var refresh_key = access_token.refresh_token;
  var a = SpreadsheetApp;
  var aS = a.getActiveSpreadsheet().getSheetByName("Refresh_Key");
  var refreshKeyStorage = aS.getRange(1, 1).setValue(refresh_key);
  
  Logger.log(token_key);
  
  var url = "https://quickbooks.api.intuit.com/v3/company/COMPANYNUMBER/journalentry";
  var Auth_Token = "bearer " + token_key;
  
  var options = 
      { 
        "method" : "POST",
        "headers": {
          "Content-Type": "application/json",
          "Accept": "application/json",
          "Authorization": Auth_Token,
        },
        "payload": payload,
      };
  var response = UrlFetchApp.fetch(url,options);
  Logger.log(response); 
  
  // Logger.log(payload);
  
  Logger.log(JSON.stringify(payload));
} 

