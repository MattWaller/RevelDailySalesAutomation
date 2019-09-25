function JEToQBO(){
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Journal Entry");
  
  var LineOneAmt = activeSheet.getRange(2, 9).getValue();
  var LineOneName = activeSheet.getRange(2, 3).getValue();
  var LineOnePT = activeSheet.getRange(2, 8).getValue();
  
  var LineTwoAmt = activeSheet.getRange(3, 9).getValue();
  var LineTwoName = activeSheet.getRange(3, 3).getValue();
  var LineTwoPT = activeSheet.getRange(3, 8).getValue();
  
  var LineThreeAmt = activeSheet.getRange(4, 9).getValue();
  var LineThreeName = activeSheet.getRange(4, 3).getValue();
  var LineThreePT = activeSheet.getRange(4, 8).getValue();
  
  var LineFourAmt = activeSheet.getRange(5, 9).getValue();
  var LineFourName = activeSheet.getRange(5, 3).getValue();
  var LineFourPT = activeSheet.getRange(5, 8).getValue();
  
  var LineFiveAmt = activeSheet.getRange(6, 9).getValue();
  var LineFiveName = activeSheet.getRange(6, 3).getValue();
  var LineFivePT = activeSheet.getRange(6, 8).getValue(); 
  
  var LineSixAmt = activeSheet.getRange(7, 9).getValue();
  var LineSixName = activeSheet.getRange(7, 3).getValue();   
  var LineSixPT = activeSheet.getRange(7, 8).getValue();
  
  var LineSevenAmt = activeSheet.getRange(8, 9).getValue();
  var LineSevenName = activeSheet.getRange(8, 3).getValue();   
  var LineSevenPT = activeSheet.getRange(8, 8).getValue();
  
  var LineEightAmt = activeSheet.getRange(9, 9).getValue();
  var LineEightName = activeSheet.getRange(9, 3).getValue();   
  var LineEightPT = activeSheet.getRange(9, 8).getValue();
  
  var LineNineAmt = activeSheet.getRange(10, 9).getValue();
  var LineNineName = activeSheet.getRange(10, 3).getValue();              
  var LineNinePT = activeSheet.getRange(10, 8).getValue();
  
  var LineTenAmt = activeSheet.getRange(11, 9).getValue();
  var LineTenName = activeSheet.getRange(11, 3).getValue();      
  var LineTenPT = activeSheet.getRange(11, 8).getValue();
  
  var LineElevenAmt = activeSheet.getRange(12, 9).getValue();
  var LineElevenName = activeSheet.getRange(12, 3).getValue();         
  var LineElevenPT = activeSheet.getRange(12, 8).getValue();
  
  var LineTwelveAmt = activeSheet.getRange(13, 9).getValue();
  var LineTwelveName = activeSheet.getRange(13, 3).getValue();      
  var LineTwelvePT = activeSheet.getRange(13, 8).getValue();
  
  var LineThirteenAmt = activeSheet.getRange(14, 9).getValue();
  var LineThirteenName = activeSheet.getRange(14, 3).getValue();   
  var LineThirteenPT = activeSheet.getRange(14, 8).getValue();
  
  var LineFourteenAmt = activeSheet.getRange(15, 9).getValue();
  var LineFourteenName = activeSheet.getRange(15, 3).getValue();   
  var LineFourteenPT = activeSheet.getRange(15, 8).getValue();
  
  var LineFifteenAmt = activeSheet.getRange(16, 9).getValue();
  var LineFifteenName = activeSheet.getRange(16, 3).getValue();   
  var LineFifteenPT = activeSheet.getRange(16, 8).getValue();
  
  var LineSixteenAmt = activeSheet.getRange(17, 9).getValue();
  var LineSixteenName = activeSheet.getRange(17, 3).getValue();  
  var LineSixteenPT = activeSheet.getRange(17, 8).getValue();
  
  var LineSeventeenAmt = activeSheet.getRange(18, 9).getValue();
  var LineSeventeenName = activeSheet.getRange(18, 3).getValue();  
  var LineSeventeenPT = activeSheet.getRange(18, 8).getValue();
  
  var LineEightteenAmt = activeSheet.getRange(19, 9).getValue();
  var LineEightteenName = activeSheet.getRange(19, 3).getValue();  
  var LineEightteenPT = activeSheet.getRange(19, 8).getValue(); 
  
  var LineNineteenAmt = activeSheet.getRange(20, 9).getValue();
  var LineNineteenName = activeSheet.getRange(20, 3).getValue();  
  var LineNineteenPT = activeSheet.getRange(20, 8).getValue();   
  
  var payload = 
      
      {
        "Line": [
          {
            "Id": "0",
            "Amount": LineOneAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineOnePT,
              "AccountRef": {
                "value": "192",
                "name": LineOneName
              }
            }
          },
          {
            "Id": "1",
            "Amount": LineTwoAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineTwoPT,
              "AccountRef": {
                "value": "192",
                "name": LineTwoName
              }
            }
          },
          {
            "Id": "2",
            "Amount": LineThreeAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineThreePT,
              "AccountRef": {
                "value": "242",
                "name": LineThreeName
              }
            }
          },
          {
            "Id": "3",
            "Amount": LineFourAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineFourPT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "216",
                  "name": "Moneris"
                }
              },
              "AccountRef": {
                "value": "375",
                "name": LineFourName
              }
            }
          },
          {
            "Id": "4",
            "Amount": LineFiveAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineFivePT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "216",
                  "name": "Moneris"
                }
              },
              "AccountRef": {
                "value": "375",
                "name": LineFiveName
              }
            }
          },
          {
            "Id": "5",
            "Amount": LineSixAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineSixPT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "216",
                  "name": "Moneris"
                }
              },
              "AccountRef": {
                "value": "375",
                "name": LineSixName
              }
            }
          },
          {
            "Id": "6",
            "Amount": LineSevenAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineSevenPT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "216",
                  "name": "Moneris"
                }
              },
              "AccountRef": {
                "value": "375",
                "name": LineSevenName
              }
            }
          },
          {
            "Id": "7",
            "Amount": LineEightAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineEightPT,
              "AccountRef": {
                "value": "302",
                "name": LineEightName
              }
            }
          },
          {
            "Id": "8",
            "Amount": LineNineAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineNinePT,
              "AccountRef": {
                "value": "242",
                "name": LineNineName
              }
            }
          },
          {
            "Id": "9",
            "Amount": LineTenAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineTenPT,
              "AccountRef": {
                "value": "242",
                "name": LineTenName
              }
            }
          },
          {
            "Id": "10",
            "Amount": LineElevenAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineElevenPT,
              "AccountRef": {
                "value": "303",
                "name": LineElevenName
              }
            }
          },
          {
            "Id": "11",
            "Amount": LineTwelveAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineTwelvePT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "238",
                  "name": "Revel POS"
                }
              },
              "AccountRef": {
                "value": "256",
                "name": LineTwelveName
              }
            }
          },
          {
            "Id": "12",
            "Amount": LineThirteenAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineThirteenPT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "238",
                  "name": "Revel POS"
                }
              },
              "AccountRef": {
                "value": "256",
                "name": LineThirteenName
              }
            }
          },
          {
            "Id": "13",
            "Amount": LineFourteenAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineFourteenPT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "238",
                  "name": "Revel POS"
                }
              },
              "AccountRef": {
                "value": "256",
                "name": LineFourteenName
              }
            }
          },
          {
            "Id": "14",
            "Amount": LineFifteenAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineFifteenPT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "238",
                  "name": "Revel POS"
                }
              },
              "AccountRef": {
                "value": "256",
                "name": LineFifteenName
              }
            }
          },
          {
            "Id": "15",
            "Amount": LineSixteenAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineSixteenPT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "238",
                  "name": "Revel POS"
                }
              },
              "AccountRef": {
                "value": "256",
                "name": LineSixteenName
              }
            }
          },
          {
            "Id": "16",
            "Amount": LineSeventeenAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineSeventeenPT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "238",
                  "name": "Revel POS"
                }
              },
              "AccountRef": {
                "value": "256",
                "name": LineSeventeenName
              }
            }
          },
          {
            "Id": "17",
            "Amount": LineEightteenAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineEightteenPT,
              "Entity": {
                "Type": "Customer",
                "EntityRef": {
                  "value": "238",
                  "name": "Revel POS"
                }
              },
              "AccountRef": {
                "value": "256",
                "name": LineEightteenName
              }
            }
          },
          {
            "Id": "18",
            "Amount": LineNineteenAmt,
            "DetailType": "JournalEntryLineDetail",
            "JournalEntryLineDetail": {
              "PostingType": LineNineteenPT,
              "AccountRef": {
                "value": "192",
                "name": LineNineteenName
              }
            }
          }
        ],
        "TxnTaxDetail": {}
        
      };
  
  
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
        "payload": JSON.stringify(payload),
      };
  var response = UrlFetchApp.fetch(url,options);
  Logger.log(response); 
  
  // Logger.log(payload);
  
  Logger.log(JSON.stringify(payload));
}
