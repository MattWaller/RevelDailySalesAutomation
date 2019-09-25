 function invoiceToQBO(){ 
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Invoice");
  var ProductOneQTY = activeSheet.getRange(5, 3).getValue();
  var ProductOneRate = activeSheet.getRange(5, 4).getValue();
  var ProductOneAmt = activeSheet.getRange(5, 5).getValue();
  var ProductTwoQTY = activeSheet.getRange(6, 3).getValue();
  var ProductTwoRate = activeSheet.getRange(6, 4).getValue();
  var ProductTwoAmt = activeSheet.getRange(6, 5).getValue();
  var ProductThreeQTY = activeSheet.getRange(7, 3).getValue();
  var ProductThreeRate = activeSheet.getRange(7, 4).getValue();
  var ProductThreeAmt = activeSheet.getRange(7, 5).getValue();
  var ProductFourQTY = activeSheet.getRange(8, 3).getValue();
  var ProductFourRate = activeSheet.getRange(8, 4).getValue();
  var ProductFourAmt = activeSheet.getRange(8, 5).getValue();
  var ProductFiveQTY = activeSheet.getRange(9, 3).getValue();
  var ProductFiveRate = activeSheet.getRange(9, 4).getValue();
  var ProductFiveAmt = activeSheet.getRange(9, 5).getValue();  
  var ProductSixQTY = activeSheet.getRange(10, 3).getValue();
  var ProductSixRate = activeSheet.getRange(10, 4).getValue();
  var ProductSixAmt = activeSheet.getRange(10, 5).getValue();  
  var ProductSevenQTY = activeSheet.getRange(11, 3).getValue();
  var ProductSevenRate = activeSheet.getRange(11, 4).getValue();
  var ProductSevenAmt = activeSheet.getRange(11, 5).getValue();   
  var ProductEightQTY = activeSheet.getRange(12, 3).getValue();
  var ProductEightRate = activeSheet.getRange(12, 4).getValue();
  var ProductEightAmt = activeSheet.getRange(12, 5).getValue();     
  var ProductNineQTY = activeSheet.getRange(13, 3).getValue();
  var ProductNineRate = activeSheet.getRange(13, 4).getValue();
  var ProductNineAmt = activeSheet.getRange(13, 5).getValue();       
  var ProductTenQTY = activeSheet.getRange(14, 3).getValue();
  var ProductTenRate = activeSheet.getRange(14, 4).getValue();
  var ProductTenAmt = activeSheet.getRange(14, 5).getValue();   
  var ProductElevenQTY = activeSheet.getRange(15, 3).getValue();
  var ProductElevenRate = activeSheet.getRange(15, 4).getValue();
  var ProductElevenAmt = activeSheet.getRange(15, 5).getValue();     
  var ProductTwelveQTY = activeSheet.getRange(16, 3).getValue();
  var ProductTwelveRate = activeSheet.getRange(16, 4).getValue();
  var ProductTwelveAmt = activeSheet.getRange(16, 5).getValue();     
  var ProductThirteenQTY = activeSheet.getRange(17, 3).getValue();
  var ProductThirteenRate = activeSheet.getRange(17, 4).getValue();
  var ProductThirteenAmt = activeSheet.getRange(17, 5).getValue();     
  var ProductFourteenQTY = activeSheet.getRange(18, 3).getValue();
  var ProductFourteenRate = activeSheet.getRange(18, 4).getValue();
  var ProductFourteenAmt = activeSheet.getRange(18, 5).getValue();   
  var ProductFifteenQTY = activeSheet.getRange(19, 3).getValue();
  var ProductFifteenRate = activeSheet.getRange(19, 4).getValue();
  var ProductFifteenAmt = activeSheet.getRange(19, 5).getValue();   
  var ProductSixteenQTY = activeSheet.getRange(20, 3).getValue();
  var ProductSixteenRate = activeSheet.getRange(20, 4).getValue();
  var ProductSixteenAmt = activeSheet.getRange(20, 5).getValue();   
  var ProductSeventeenQTY = activeSheet.getRange(21, 3).getValue();
  var ProductSeventeenRate = activeSheet.getRange(21, 4).getValue();
  var ProductSeventeenAmt = activeSheet.getRange(21, 5).getValue();   
  var ProductEighteenQTY = activeSheet.getRange(22, 3).getValue();
  var ProductEighteenRate = activeSheet.getRange(22, 4).getValue();
  var ProductEighteenAmt = activeSheet.getRange(22, 5).getValue();   
  var ProductNineteenQTY = activeSheet.getRange(23, 3).getValue();
  var ProductNineteenRate = activeSheet.getRange(23, 4).getValue();
  var ProductNineteenAmt = activeSheet.getRange(23, 5).getValue();  
  var ProductTwentyQTY = activeSheet.getRange(24, 3).getValue();
  var ProductTwentyRate = activeSheet.getRange(24, 4).getValue();
  var ProductTwentyAmt = activeSheet.getRange(24, 5).getValue();  
  var ProductTwentyOneQTY = activeSheet.getRange(25, 3).getValue();
  var ProductTwentyOneRate = activeSheet.getRange(25, 4).getValue();
  var ProductTwentyOneAmt = activeSheet.getRange(25, 5).getValue();  
  var ProductTwentyTwoQTY = activeSheet.getRange(27, 3).getValue();
  var ProductTwentyTwoRate = activeSheet.getRange(27, 4).getValue();
  var ProductTwentyTwoAmt = activeSheet.getRange(27, 5).getValue();  
  var ProductTwentyThreeQTY = activeSheet.getRange(28, 3).getValue();
  var ProductTwentyThreeRate = activeSheet.getRange(28, 4).getValue();
  var ProductTwentyThreeAmt = activeSheet.getRange(28, 5).getValue();  
  var ProductTwentyFourQTY = activeSheet.getRange(29, 3).getValue();
  var ProductTwentyFourRate = activeSheet.getRange(29, 4).getValue();
  var ProductTwentyFourAmt = activeSheet.getRange(29, 5).getValue();  
  var ProductTwentyFiveQTY = activeSheet.getRange(30, 3).getValue();
  var ProductTwentyFiveRate = activeSheet.getRange(30, 4).getValue();
  var ProductTwentyFiveAmt = activeSheet.getRange(30, 5).getValue();   
  var ProductTwentySixQTY = activeSheet.getRange(31, 3).getValue();
  var ProductTwentySixRate = activeSheet.getRange(31, 4).getValue();
  var ProductTwentySixAmt = activeSheet.getRange(31, 5).getValue();     
  var ProductTwentySevenQTY = activeSheet.getRange(32, 3).getValue();
  var ProductTwentySevenRate = activeSheet.getRange(32, 4).getValue();
  var ProductTwentySevenAmt = activeSheet.getRange(32, 5).getValue();       
  var ProductTwentyEightQTY = activeSheet.getRange(33, 3).getValue();
  var ProductTwentyEightRate = activeSheet.getRange(33, 4).getValue();
  var ProductTwentyEightAmt = activeSheet.getRange(33, 5).getValue();         
  var PackageOneQTY = activeSheet.getRange(26, 8).getValue();
  var PackageOneAmt = activeSheet.getRange(26, 5).getValue();   
  var Subtotal = activeSheet.getRange(34, 5).getValue();
  var GST = activeSheet.getRange(35, 5).getValue();
  var PST = activeSheet.getRange(36, 5).getValue();
  var TotalTax = activeSheet.getRange(37, 5).getValue();
  var InvoiceTotal = activeSheet.getRange (38, 5).getValue();
   var payload = 
    {
      "Line": [
      {
        "Id": "1",
        "LineNum": 1,
        "Description": "12G Buck",
        "Amount": ProductOneAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "30",
            "name": "12G BUCK x 3"
          },
          "UnitPrice": ProductOneRate,
          "Qty": ProductOneQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "2",
        "LineNum": 2,
        "Description": "12G Slug",
        "Amount": ProductTwoAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "31",
            "name": "12G SLUG x 3"
          },
          "UnitPrice": ProductTwoRate,
          "Qty": ProductTwoQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "3",
        "LineNum": 3,
        "Description": "223 Ammo",
        "Amount": ProductThreeAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "32",
            "name": "223 x 10"
          },
          "UnitPrice": ProductThreeRate,
          "Qty": ProductThreeQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "4",
        "LineNum": 4,
        "Description": "22 LR",
        "Amount": ProductFourAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "33",
            "name": ".22 LR X10"
          },
          "UnitPrice": ProductFourRate,
          "Qty": ProductFourQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "5",
        "LineNum": 5,
        "Description": "30 - 30",
        "Amount": ProductFiveAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "34",
            "name": "30 - 30 x 5"
          },
          "UnitPrice": ProductFiveRate,
          "Qty": ProductFiveQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "6",
        "LineNum": 6,
        "Description": "308",
        "Amount": ProductSixAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "35",
            "name": "308 x 8"
          },
          "UnitPrice": ProductSixRate,
          "Qty": ProductSixQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "7",
        "LineNum": 7,
        "Description": "357 Mag",
        "Amount": ProductSevenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "36",
            "name": "357 MAGNUM x 6"
          },
          "UnitPrice": ProductSevenRate,
          "Qty": ProductSevenQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "8",
        "LineNum": 8,
        "Description": "38 Special",
        "Amount": ProductEightAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "38",
            "name": ".38 Spec x 6"
          },
          "UnitPrice": ProductEightRate,
          "Qty": ProductEightQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "9",
        "LineNum": 9,
        "Description": "40 cal",
        "Amount": ProductNineAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "39",
            "name": "40 cal x 10"
          },
          "UnitPrice": ProductNineRate,
          "Qty": ProductNineQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "10",
        "LineNum": 10,
        "Description": "44 Magnum",
        "Amount": ProductTenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "40",
            "name": "44 MAGNUM x 6"
          },
          "UnitPrice": ProductTenRate,
          "Qty": ProductTenQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "11",
        "LineNum": 11,
        "Description": "45 ACP",
        "Amount": ProductElevenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "41",
            "name": "45 ACP x 10"
          },
          "UnitPrice": ProductElevenRate,
          "Qty": ProductElevenQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "12",
        "LineNum": 12,
        "Description": "45 Colt Long",
        "Amount": ProductTwelveAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "42",
            "name": "45 COLT LONG x 6"
          },
          "UnitPrice": ProductTwelveRate,
          "Qty": ProductTwelveQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "13",
        "LineNum": 13,
        "Description": "500 S&W",
        "Amount": ProductThirteenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "43",
            "name": "500 S&W x 1"
          },
          "UnitPrice": ProductThirteenRate,
          "Qty": ProductThirteenQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "14",
        "LineNum": 14,
        "Description": "50 AE",
        "Amount": ProductFourteenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "44",
            "name": "50 AE x 1"
          },
          "UnitPrice": ProductFourteenRate,
          "Qty": ProductFourteenQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "15",
        "LineNum": 15,
        "Description": "7.62 NON CRSV",
        "Amount": ProductFifteenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "45",
            "name": "7.62 NON CRSV x 10"
          },
          "UnitPrice": ProductFifteenRate,
          "Qty": ProductFifteenQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "16",
        "LineNum": 16,
        "Description": "9mm",
        "Amount": ProductSixteenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "46",
            "name": "9mm x 10"
          },
          "UnitPrice": ProductSixteenRate,
          "Qty": ProductSixteenQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "17",
        "LineNum": 17,
        "Description": "BMG .50",
        "Amount": ProductSeventeenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "47",
            "name": "BMG .50 SNIPER x 1"
          },
          "UnitPrice": ProductSeventeenRate,
          "Qty": ProductSeventeenQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "18",
        "LineNum": 18,
        "Description": "Hats",
        "Amount": ProductEighteenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "48",
            "name": "HATS"
          },
          "UnitPrice": ProductEighteenRate,
          "Qty": ProductEighteenQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "19",
        "LineNum": 19,
        "Description": "Stickers",
        "Amount": ProductNineteenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "52",
            "name": "STICKER"
          },
          "UnitPrice": ProductNineteenRate,
          "Qty": ProductNineteenQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "20",
        "LineNum": 20,
        "Description": "Targets",
        "Amount": ProductTwentyAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "51",
            "name": "Targets"
          },
          "UnitPrice": ProductTwentyRate,
          "Qty": ProductTwentyQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "21",
        "LineNum": 21,
        "Amount": PackageOneAmt,
        "DetailType": "GroupLineDetail",
        "GroupLineDetail": {
          "GroupItemRef": {
            "value": "68",
            "name": "M60 BELT x 15"
          },
          "Quantity": PackageOneQTY,
          "Line": [
            {
              "Id": "22",
              "LineNum": 22,
              "Description": "308",
              "Amount": ProductTwentyTwoAmt,
              "DetailType": "SalesItemLineDetail",
              "SalesItemLineDetail": {
                "ItemRef": {
                  "value": "35",
                  "name": "308 x 8"
                },
                "UnitPrice": ProductTwentyTwoRate,
                "Qty": ProductTwentyTwoQTY,
                "TaxCodeRef": {
                  "value": "7"
                }
              }
            }
          ]
        }
      },
      {
        "Id": "24",
        "LineNum": 24,
        "Amount": ProductTwentyThreeAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "64",
            "name": "Drop-in Fees"
          },
          "UnitPrice": ProductTwentyThreeRate,
          "Qty": ProductTwentyThreeQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "25",
        "LineNum": 25,
        "Amount": ProductTwentyFourAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "65",
            "name": "Gun Rental"
          },
          "UnitPrice": ProductTwentyFourRate,
          "Qty": ProductTwentyFourQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "26",
        "LineNum": 26,
        "Amount": ProductTwentyFiveAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "66",
            "name": "Membership"
          },
          "UnitPrice": ProductTwentyFiveRate,
          "Qty": ProductTwentyFiveQTY,
          "TaxCodeRef": {
            "value": "3"
          }
        }
      },
      {
        "Id": "27",
        "LineNum": 27,
        "Amount": ProductTwentySixAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "70",
            "name": "MISC"
          },
          "UnitPrice": ProductTwentySixRate,
          "Qty": ProductTwentySixQTY,
          "TaxCodeRef": {
            "value": "7"
          }
        }
      },
      {
        "Id": "28",
        "LineNum": 28,
        "Amount": ProductTwentySevenAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "69",
            "name": "MISC NT"
          },
          "UnitPrice": ProductTwentySevenRate,
          "Qty": ProductTwentySevenQTY,
          "TaxCodeRef": {
            "value": "2"
          }
        }
      },
      {
        "Id": "29",
        "LineNum": 29,
        "Amount": ProductTwentyEightAmt,
        "DetailType": "SalesItemLineDetail",
        "SalesItemLineDetail": {
          "ItemRef": {
            "value": "71",
            "name": "MISC GST"
          },
          "UnitPrice": ProductTwentyEightRate,
          "Qty": ProductTwentyEightQTY,
          "TaxCodeRef": {
            "value": "3"
          }
        }
      },
        {
          "Id": "33",
          "LineNum": 30,
          "Description": "T-shirts",
          "Amount": ProductTwentyOneAmt,
          "DetailType": "SalesItemLineDetail",
          "SalesItemLineDetail": {
            "ItemRef": {
              "value": "50",
              "name": "T-Shirt"
            },
            "UnitPrice": ProductTwentyOneRate,
            "Qty": ProductTwentyOneQTY,
            "TaxCodeRef": {
              "value": "7"
            }
          }
        },
        {
        "Amount": Subtotal,
        "DetailType": "SubTotalLineDetail",
        "SubTotalLineDetail": {}
      }
    ],
    "TxnTaxDetail": {
      "TotalTax": TotalTax,
      "TaxLine": [
        {
          "Amount": GST,
          "DetailType": "TaxLineDetail",
          "TaxLineDetail": {
            "TaxRateRef": {
              "value": "4"
            },
            "PercentBased": true,
            "TaxPercent": 5,
            "NetAmountTaxable": Subtotal
          }
        },
        {
          "Amount": PST,
          "DetailType": "TaxLineDetail",
          "TaxLineDetail": {
            "TaxRateRef": {
              "value": "18"
            },
            "PercentBased": true,
            "TaxPercent": 7,
            "NetAmountTaxable": Subtotal
          }
        },
        {
          "Amount": 0,
          "DetailType": "TaxLineDetail",
          "TaxLineDetail": {
            "TaxRateRef": {
              "value": "2"
            },
            "PercentBased": true,
            "TaxPercent": 0,
            "NetAmountTaxable": 0
          }
        }
      ]
    },
    "CustomerRef": {
      "value": "238",
      "name": "Revel POS"
    }

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
           "Authorization": "Basic AUTHKEY"
         },
         "payload": {
           "grant_type" : "refresh_token",
           "refresh_token" : refresh_k,
         }
       };

   var token_response = UrlFetchApp.fetch(token_url,token_options); 
   var access_token = JSON.parse(token_response);
   var token_key = access_token.access_token;
   Logger.log(token_key);
   
   var refresh_key = access_token.refresh_token;
   var a = SpreadsheetApp;
   var aS = a.getActiveSpreadsheet().getSheetByName("Refresh_Key");
   var refreshKeyStorage = aS.getRange(1, 1).setValue(refresh_key);
   
   
   
   
   
   
   
   
   var url = "https://quickbooks.api.intuit.com/v3/company/COMPANYNUMBER/invoice";
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
 };


