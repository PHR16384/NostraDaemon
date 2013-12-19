/*
 * PDF creator and emailer; adapted from http://stackoverflow.com/questions/12881547/exporting-spreadsheet-to-pdf-then-saving-the-file-in-google-drive
*/
function mailPDF(SS)
{
  var sKey = "0AticIhRKqBBedEpRVFJiUmNCRHlYbm81UU83eXFhdmc";
  
  if( isNothing(SS) ){
    SS = SpreadsheetApp.openById(sKey);
  }
  /*
  if( isNothing(sName) ){
    sName = "PHR";
  }
  if( isNothing(sEmail) ){
    sEmail = "tjk16384@gmail.com";
  }
  */
  
  var shForecast = SS.getSheetByName("Sales Forecast");
  var sName = shForecast.getRange("B26").getValue();
  var sEmail = shForecast.getRange("B28").getValue();
  
  //SS.toast("Mailing report to '" + sName + "' (" + sEmail + ")...");
  
  var shReport = SS.getSheetByName("Report");
  var shReportE = SS.getSheetByName("Email Report");
  var gID = shReportE.getSheetId();
  
  var oauthConfig = UrlFetchApp.addOAuthService("google");
  oauthConfig.setAccessTokenUrl("https://www.google.com/accounts/OAuthGetAccessToken");
  oauthConfig.setRequestTokenUrl("https://www.google.com/accounts/OAuthGetRequestToken?scope=https://spreadsheets.google.com/feeds/");
  oauthConfig.setAuthorizationUrl("https://www.google.com/accounts/OAuthAuthorizeToken");
  oauthConfig.setConsumerKey("anonymous");
  oauthConfig.setConsumerSecret("anonymous");
  
  var requestData = {
    "method": "GET",
    "oAuthServiceName": "google",
    "oAuthUseToken": "always"
  };
  
  var sURL = "https://spreadsheets.google.com/feeds/download/spreadsheets/Export?key=" + sKey + "&gid=" + gID + "&fitw=true&size=0&portrait=false&sheetnames=false&printtitle=false&exportFormat=pdf&format=pdf&gridlines=false";
  var oPDF = UrlFetchApp.fetch(sURL, requestData).getBlob().getBytes();
  
  var sSubj = shReport.getRange("B2").getValue();
  if( isNothing(sSubj) )
  {
    sSubj = sName + ", here's your Report from ICC";
  }
  var message = sName + ",\n\n" + shReport.getRange("B3").getValue();
  if( isNothing(message) )
  {
    message = "Thanks for choosing ICC!  Your report is attached below.";
  }
  
  var attach = {fileName:'Report.pdf',content:oPDF, mimeType:'application/pdf'};

  // Send the freshly constructed email 
  GmailApp.sendEmail(sEmail, sSubj, message, {
    attachments: [attach],
    cc: "info@innovativecloudcats.com"
  });
  
  //SS.toast("PDF sent.");
  
}


