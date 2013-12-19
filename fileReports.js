function fileReports(SS)
{
  if( isNothing(SS) ){
    SS = SpreadsheetApp.openById("0AticIhRKqBBedEpRVFJiUmNCRHlYbm81UU83eXFhdmc");
  }
  
  //SS.toast("Processing unsent reports...");
  
  var x=0, y=0;
  
  // REMEMBER TO UPDATE CONCURRENT WITH FORM
  var nRev = 84;
  
  var shSurveys = SS.getSheetByName("Revision: "+nRev.toString());
  //var shRef = SS.getSheetByName("Input1");
  
  var shForecast = SS.getSheetByName("Sales Forecast");
  var shRV = SS.getSheetByName("Rule Variables");
  var shRules = SS.getSheetByName("Rules");
  var shReport = SS.getSheetByName("Report");
  
  var rRows = shSurveys.getDataRange();  // grab the "entire" Sheet and turn it into a Range...
  var arRows = rRows.getValues();  //then into a 2D array
  
  var arReport = shReport.getDataRange().getValues();
  Logger.log(arReport);
  
  // In case column # of "Email", "Report filed?", etc. changes in future versions, find them manually:
  var nDate = -1, nFName = -1, nLName = -1, nEmail = -1, nPhone = -1, nFiled = -1, nProdName = -1, nMonths = -1;
  var nCompany = -1, nWebsite = -1;
  var nPriceI = -1, nPriceM = -1, nCogI = -1, nCogM = -1, nLabourI = -1, nLabourM = -1, nPercentI = -1, nPercentM = -1;
  
  for(x=0; x<arRows[0].length; x++)
  {
    //Logger.log(arRows[0][x]);
    
    switch(arRows[0][x])  // switch block is much cleaner than a chain of else-ifs
    {
      case "Report filed?":
        nFiled = x;
        break;
      case "Submitted Date":
        nDate = x;
        break;
        
      case "First Name:":
        nFName = x;
        break;
      case "Last Name:":
        nLName = x;
        break;
      case "Email:":
        nEmail = x;
        break;
      case "Phone #:":
        nPhone = x;
        break;
        
      case "Company Name:":
        nCompany = x;
        break;
      case "Website URL:":
        nWebsite = x;
        break;
        
      case "Product Name:":
        nProdName = x;
        break;
        
      case "Initial Price:":
        nPriceI = x;
        break;
      case "Monthly Price:":
        nPriceM = x;
        break;
      case "Initial COG:":
        nCogI = x;
        break;
      case "Monthly COG:":
        nCogM = x;
        break;
      case "Initial Labour:":
        nLabourI = x;
        break;
      case "Monthly Labour:":
        nLabourM = x;
        break;
      case "Initial %:":
        nPercentI = x;
        break;
      case "Monthly %:":
        nPercentM = x;
        break;
        
      case "January":  // I'll pretend that the monthly values will stay in order
        nMonths = x;
        break;
        
    }
  }
  
  //Logger.log(nFName);
  //Logger.log(nFiled);
  //Logger.log(nEmail);
  
  var arMonths = null;
  var obj = null;
  
  for(x=1; x<arRows.length; x++)
  {
    // find empty "Report filed?" rows and file their reports:
    if( isNothing(arRows[x][nFiled]) )
    {
      // If all four "Initial cost" fields are blank (blank, not zero!), copypasta the four monthly values:
      if( arRows[x][nPriceI]=="" && arRows[x][nCogI]=="" && arRows[x][nLabourI]=="" && arRows[x][nPercentI]=="" )
      {
        arRows[x][nPriceI] = arRows[x][nPriceM];
        arRows[x][nCogI] = arRows[x][nCogM];
        arRows[x][nLabourI] = arRows[x][nLabourM];
        arRows[x][nPercentI] = arRows[x][nPercentM];
      }
      
      // copypasta form results:
      shForecast.getRange("A2").setValue(arRows[x][nProdName]);
      shRV.getRange("D39:D40").setValues([ [ arRows[x][nPercentI] ], [ arRows[x][nPercentM] ] ]);
      shRules.getRange("D5").setValue(arRows[x][nPriceI]);
      shRules.getRange("E8").setValue(arRows[x][nPriceM]);
      shRules.getRange("D13:E13").setValues([ [arRows[x][nCogI],arRows[x][nCogM]] ]);
      shRules.getRange("D22:E22").setValues([ [arRows[x][nLabourI],arRows[x][nLabourM]] ]);
      arMonths = shSurveys.getRange(x+1,nMonths+1, 1,12).getValues();  // WARNING: row & column #s start counting from 1, not 0. :(
      shForecast.getRange("F2:Q2").setValues(arMonths);
      
      shForecast.getRange("B26:B32").setValues([ [arRows[x][nFName]],[arRows[x][nLName]],[arRows[x][nEmail]],[arRows[x][nPhone]],[arRows[x][nDate]],[arRows[x][nCompany]],[arRows[x][nWebsite]] ]);
      
      SpreadsheetApp.flush();
      
      // Generate a Report...
      nostraDaemon(SS);
      
      // ... then mail the Report to the person
      mailPDF(SS);
      
      // Set "Report filed?" to "Yes":
      shSurveys.getRange(x+1, nFiled+1).setValue("Y");  // WARNING: row & column #s start counting from 1, not 0. :(
      //Logger.log([x,nFiled]);
    }
  }
  
  //MailApp.sendEmail("tjk16384@gmail.com", "New Responses to Banana", "Back to work, then!  *whip crack*");
  
  //SS.toast("Finished sending reports.");
  
}


