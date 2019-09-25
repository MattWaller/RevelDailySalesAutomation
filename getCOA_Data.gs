// get chart of account data 
function getCOA_Data() {
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
  
  
  
  var url = "https://quickbooks.api.intuit.com/v3/company/COMPANYNUMBER/query?minorversion=37";
  var Auth_Token = "bearer " + token_key;
  
  // get Asset Chart of Account accounts
  var options = 
      { 
        "method" : "POST",
        "headers": {
          "Content-Type": "application/text",
          "Accept": "application/json",
          "Authorization": Auth_Token,
        },
        "payload": "Select * from account where Classification = 'Asset'",
      };
  var response = UrlFetchApp.fetch(url,options);
  var obj = JSON.parse(response);
  var QR = obj.QueryResponse
  var QRlen = QR.Account.length;
  var b = SpreadsheetApp;
  var bS = b.getActiveSpreadsheet().getSheetByName("QBO-GL-INFO");  
  Logger.log(QRlen);
  var i;
  for (i = 0; i< QRlen; i++){
    var row = i + 2;
    var ac = QR.Account[i];
    var Name = ac.Name
    var Id = ac.Id
    var Cur = ac.CurrencyRef.name
    var Class = ac.Classification
    var acNum = ac.AcctNum
    var set_Name = bS.getRange(row, 2).setValue(Name);
    var set_Id = bS.getRange(row,3).setValue(Id);
    var set_Cur = bS.getRange(row,4).setValue(Cur);
    var set_Class = bS.getRange(row,5).setValue(Class);
    var set_acNum = bS.getRange(row,6).setValue(acNum);
  }
  // To get Equity Chart of accounts

  var Equity = 
      {
        "method" : "POST",
        "headers": {
          "Content-Type": "application/text",
          "Accept": "application/json",
          "Authorization": Auth_Token,
        },
        "payload": "Select * from account where Classification = 'Equity'",
      };    
  var eqResponse = UrlFetchApp.fetch(url,Equity);
  var eqObj = JSON.parse(eqResponse);
  var eqQR = eqObj.QueryResponse
  var eqQRlen = eqQR.Account.length;
  Logger.log(eqQRlen);
  var eq;
  for (eq = 0; eq< eqQRlen; eq++){
    var eqRow = eq + QRlen + 1;
    var eqac = eqQR.Account[eq];
    var eqName = eqac.Name
    var eqId = eqac.Id
    var eqCur = eqac.CurrencyRef.name
    var eqClass = eqac.Classification
    var eqacNum = eqac.AcctNum
    var eqset_Name = bS.getRange(eqRow, 2).setValue(eqName);
    var eqset_Id = bS.getRange(eqRow,3).setValue(eqId);
    var eqset_Cur = bS.getRange(eqRow,4).setValue(eqCur);
    var eqset_Class = bS.getRange(eqRow,5).setValue(eqClass);
    var eqset_acNum = bS.getRange(eqRow,6).setValue(eqacNum);
  } 
  
  // To get Expense Chart of accounts

  var Expense = 
      {
        "method" : "POST",
        "headers": {
          "Content-Type": "application/text",
          "Accept": "application/json",
          "Authorization": Auth_Token,
        },
        "payload": "Select * from account where Classification = 'Expense'",
      };    
  var exResponse = UrlFetchApp.fetch(url,Expense);
  var exObj = JSON.parse(exResponse);
  var exQR = exObj.QueryResponse
  var exQRlen = exQR.Account.length;
  Logger.log(exQRlen);
  var ex;
  for (ex = 0; ex< exQRlen; ex++){
    var exRow = ex + eqQRlen + 1 + QRlen;
    var exac = exQR.Account[ex];
    var exName = exac.Name
    var exId = exac.Id
    var exCur = exac.CurrencyRef.name
    var exClass = exac.Classification
    var exacNum = exac.AcctNum
    var exset_Name = bS.getRange(exRow, 2).setValue(exName);
    var exset_Id = bS.getRange(exRow,3).setValue(exId);
    var exset_Cur = bS.getRange(exRow,4).setValue(exCur);
    var exset_Class = bS.getRange(exRow,5).setValue(exClass);
    var exset_acNum = bS.getRange(exRow,6).setValue(exacNum);    
  }   
  
    // To get Liability Chart of accounts

  var Liability = 
      {
        "method" : "POST",
        "headers": {
          "Content-Type": "application/text",
          "Accept": "application/json",
          "Authorization": Auth_Token,
        },
        "payload": "Select * from account where Classification = 'Liability'",
      };    
  var liResponse = UrlFetchApp.fetch(url,Liability);
  var liObj = JSON.parse(liResponse);
  var liQR = liObj.QueryResponse
  var liQRlen = liQR.Account.length;
  Logger.log(liQRlen);
  var li;
  for (li = 0; li< liQRlen; li++){
    var liRow = li + eqQRlen + 1 + QRlen + exQRlen;
    var liac = liQR.Account[li];
    var liName = liac.Name
    var liId = liac.Id
    var liCur = liac.CurrencyRef.name
    var liClass = liac.Classification
    var liacNum = liac.AcctNum
    var liset_Name = bS.getRange(liRow, 2).setValue(liName);
    var liset_Id = bS.getRange(liRow,3).setValue(liId);
    var liset_Cur = bS.getRange(liRow,4).setValue(liCur);
    var liset_Class = bS.getRange(liRow,5).setValue(liClass);
    var liset_acNum = bS.getRange(liRow,6).setValue(liacNum);    
  }   
  
  // To get Revenue Chart of accounts

  var Revenue = 
      {
        "method" : "POST",
        "headers": {
          "Content-Type": "application/text",
          "Accept": "application/json",
          "Authorization": Auth_Token,
        },
        "payload": "Select * from account where Classification = 'Revenue'",
      };    
  var reResponse = UrlFetchApp.fetch(url,Revenue);
  var reObj = JSON.parse(reResponse);
  var reQR = reObj.QueryResponse
  var reQRlen = reQR.Account.length;
  Logger.log(reQRlen);
  var re;
  for (re = 0; re< reQRlen; re++){
    var reRow = re + eqQRlen + 1 + QRlen + exQRlen + liQRlen;
    var reac = reQR.Account[re];
    var reName = reac.Name
    var reId = reac.Id
    var reCur = reac.CurrencyRef.name
    var reClass = reac.Classification
    var reacNum = reac.AcctNum    
    var reset_Name = bS.getRange(reRow, 2).setValue(reName);
    var reset_Id = bS.getRange(reRow,3).setValue(reId);
    var reset_Cur = bS.getRange(reRow,4).setValue(reCur);
    var reset_Class = bS.getRange(reRow,5).setValue(reClass);
    var reset_acNum = bS.getRange(reRow,6).setValue(reacNum);    
  }   
  
}




