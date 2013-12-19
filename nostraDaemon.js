/**
 * NOSTRADAEMON:
 * Generates a Report sheet from the Sales Forecast, Expenses and Rules sheets
 * Also uses the Worksheet page for (temporary?) storage
 */
function nostraDaemon(SS)
{
  if( isNothing(SS) ){
    SS = SpreadsheetApp.openById("0AticIhRKqBBedEpRVFJiUmNCRHlYbm81UU83eXFhdmc");
  }
  
  //SS.toast("Running NostraDaemon...");
  
  var x=0, y=0, z=0;
  
  // The three "input" sheets:
  var shForecast = SS.getSheetByName("Sales Forecast");
  var shExpenses = SS.getSheetByName("Expenses");
  var shRules = SS.getSheetByName("Rules");
  
  // The "output" sheet(s):
  var shReport = SS.getSheetByName("Report");
  var shWork = SS.getSheetByName("Worksheet");
  
  //Logger.log(shForecast.getDataRange().getValues());
  //Logger.log( shWork.getRange("A4:A28").getValues() );
  
  
  // if 'Sales Forecast'!E19 is 0, quit
  if( isNothing(shForecast.getRange("E19").getValue()) )
  {
    throw "ERROR: Sales Total (E19) is 0.";
    return;
  }
  
  
  /*
   * REPORT PREPARATION
   */
  
  clearReport(SS);
  
  var rExpenses = SS.getRangeByName("Expense");
  rExpenses.copyTo( SS.getRangeByName("Report_Expense"), {contentsOnly:true} );
  //Logger.log(rExpenses);
  
  
  // Grab the entire blue block from the Sales Forecast:
  var arForecast = SS.getRangeByName("SR").getValues();
  //Logger.log(arForecast);
  
  // Grab the Lookup refs from the Worksheet:
  //var arMCRC = SS.getRangeByName("MCRClookupRef").getValues();
  //var arSC = SS.getRangeByName("SClookupRef").getValues();
  //Logger.log(arMCRC);
  //Logger.log(arSC);
  
  var arMCRC_ = [ ["D","F","H","J","L","N","P","R","T","V","X","Z","AB","AD","AF"] , ["E","G","I","K","M","O","Q","S","U","W","Y","AA","AC","AE","AG"] ];
  var arSC_ = ["F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE"];
  
  // these will hold spreadsheet column letters:
  var MC = "";
  var RC = "";
  var SC = "";
  
  
  for(x=0; x<arForecast.length; x++)  // note: SR = x+2; 15 rows expected
  {
    //MC = arMCRC[x][1];  // MC now contains a column letter
    //RC = arMCRC[x][2];
    MC = arMCRC_[0][x];
    RC = arMCRC_[1][x];
    shWork.getRange("C1").setValue("MC = " + MC);  // debug
    
    for(y=0; y<arForecast[x].length; y++)  // note: SClookup = y+1; 26 columns expected
    {
      if( isNothing(arForecast[x][y]) )
      {
        continue;
      }
      
      /*
      * REPORT GL & RESOURCE SETUP
      */
      
      // copy Rules!(MC)5:29 to Worksheet:
      shRules.getRange(MC+"5:"+MC+"29").copyTo( shWork.getRange("A4:A28"), {contentsOnly:true} );
      // copy Rules!(MC)34:73 to Worksheet:
      shRules.getRange(MC+"34:"+MC+"73").copyTo( shWork.getRange("F4:F43"), {contentsOnly:true} );
      
      //SC = arSC[y][1];  // SC now contains a column letter
      SC = arSC_[y];
      shWork.getRange("D1").setValue("SC = " + SC);  // debug
      
      /*
       * UPDATE WORKSHEET SALES VALUE
       */
      shWork.getRange("A1").setValue( arForecast[x][y] );  // copy 'Sales Forecast'!(SC)(SR) to Worksheet
      
      /*
       * MONTHLY PROCESSING:
       */
      var rTemp = shReport.getRange(SC+"4:"+SC+"28");
      rTemp.copyTo( shWork.getRange("C4:C28"), {contentsOnly:true} );
      shWork.getRange("D4:D28").copyTo( rTemp, {contentsOnly:true} );
      
      rTemp = shReport.getRange(SC+"54:"+SC+"93");
      rTemp.copyTo( shWork.getRange("H4:H43"), {contentsOnly:true} );
      shWork.getRange("I4:I43").copyTo( rTemp, {contentsOnly:true} );
      
      /*
       * RECURRING PROCESSING:
       */
      shWork.getRange("E1").setValue("RC = " + RC);  // debug
      
      if( shRules.getRange(RC+"3").getValue() == "N" )
      {
        continue;  // if the "Ordinary Expense/Income" under the current column is marked "No", skip over the R.P. code
      }
      
      // Get recurring values
      rTemp = shRules.getRange(RC+"5:"+RC+"29");
      rTemp.copyTo( shWork.getRange("A4:A28"), {contentsOnly:true} );
      rTemp = shRules.getRange(RC+"34:"+RC+"73");
      rTemp.copyTo( shWork.getRange("F4:F43"), {contentsOnly:true} );
      
      for(z=y+1; z<arForecast[x].length; z++)  // starts relative to last-processed month; max 25 months
      {
        //SC = arSC[z][1];  // should be safe to reuse var SC
        SC = arSC_[z];
        shWork.getRange("D1").setValue("SC = " + SC);
        
        // copy monthly values, and += the recurring values
        rTemp = shReport.getRange(SC+"4:"+SC+"28");
        rTemp.copyTo( shWork.getRange("C4:C28"), {contentsOnly:true} );
        shWork.getRange("D4:D28").copyTo( rTemp, {contentsOnly:true} );
        
        rTemp = shReport.getRange(SC+"54:"+SC+"93");
        rTemp.copyTo( shWork.getRange("H4:H43"), {contentsOnly:true} );
        shWork.getRange("I4:I43").copyTo( rTemp, {contentsOnly:true} );
      }
    }
  }
  
  
  // Blank out any zero values in the report:
  
  var rRange = SS.getRangeByName("Report_Profit");
  deZero( rRange,rRange.getValues() );
  rRange = SS.getRangeByName("Report_Resources");
  deZero( rRange,rRange.getValues() );
  
  
  SpreadsheetApp.flush();  // force the Spreadsheet to clear out its "calculation buffer"
  //SS.toast("Predictions made.");
  
}


