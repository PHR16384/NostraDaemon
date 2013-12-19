function clearReport(SS)
{
  if( isNothing(SS) ){
    SS = SpreadsheetApp.openById("0AticIhRKqBBedEpRVFJiUmNCRHlYbm81UU83eXFhdmc");
  }
  
  //SS.toast("Clearing Report...");
  
  var sRange = "A1:I44";
  
  // empty Report values:
  var rProfit = SS.getRangeByName("Report_Profit");
  //Logger.log(rProfit.getValues());
  //Logger.log( rProfit==null );
  rProfit.clearContent();
  
  var rRes = SS.getRangeByName("Report_Resources");
  rRes.clearContent();
  
  var rWork = SS.getSheetByName("Worksheet").getRange(sRange);
  rWork.clearContent();
  
  var rWorkB = SS.getSheetByName("Worksheet.BAK").getRange(sRange);
  rWorkB.copyTo( rWork );
  
  //shWorkB.copyTo(SS).setName("Worksheet");  // needs to be chained; the return from "copyTo" points to the copy
  
  SpreadsheetApp.flush();
  //SS.toast("Cleared.");
}

