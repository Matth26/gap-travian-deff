var sheetRecapMurListRange = sheetRecap.getRange("B4:B18");

function createMessageAlly()
{
  var gdocURL = "https://docs.google.com/spreadsheets/d/1Cr7iORwAMGvxgxA4yAYL_YuCWDmzKIpS9rUoHG07lPg/edit#gid="
  var murSheetList = getMurSheetList();
  var nbr_mur = murSheetList.length;
  
  var murNotFullSheetList = [];
  
  for(var i=0; i<nbr_mur; i++)
  {
    infos = getMurInfoFromSheetName(murSheetList[i]);
    if(infos.missingDeff > 0)
      murNotFullSheetList.push(murSheetList[i]);
  }
  
  // update the nbr_mur with only the mur which are not full
  nbr_mur = murNotFullSheetList.length;
  
  var msg = "Bonsoir à tous,\\n\\n";
  if(nbr_mur == 1)
    msg = msg  + "Merci de remplir le mur suivant:\\n\\n";
  else
    msg = msg  + "Merci de remplir les " + nbr_mur + " murs suivant:\\n\\n"
    
  for(var i=0; i<nbr_mur; i++)
  {
    var murNum = i+1;
    infos = getMurInfoFromSheetName(murNotFullSheetList[i]);
    
    msg = msg + "Mur_" + murNum + ": [x|y]" + infos.x + "/" + infos.y + "[/x|y]:\\n"
    
    var dateString = Utilities.formatDate(infos.impactDate, 'CET', 'dd/MM à hh:mm:ss');
    
    msg = msg  + "Impact le [color=red]" + dateString + "[/color]\\n" 
    msg = msg  + "Deff Manquante: [b]" + infos.missingDeff + "[/b]\\n"
    msg = msg + "Mur URL: " + gdocURL + infos.sheetID + "\\n\\n"
  }
  
  msg = msg  + "[color=blue]Envoyez au plus prêt de l'impact.[/color]\\n"
  msg = msg  + "[color=blue]Si vous ne pouvez pas être là après l'impact pour récuperer la deff, merci de trouver un cogestionnaire ![/color]\\n\\n"
  
  msg = msg  + "Merci à tous, GO BYOP *love*\\n"
  
 Browser.msgBox(msg); 
}

function getMurInfoFromSheetName(name)
{
  var murInfos = {};
  var murSheet = ss.getSheetByName(name);
  murInfos.x = murSheet.getRange('C4').getValue();
  murInfos.y = murSheet.getRange('D4').getValue();
  murInfos.impactDate = murSheet.getRange('B4').getValue();
  murInfos.missingDeff = murSheet.getRange('I3').getValue();
  murInfos.missingSpy = murSheet.getRange('I6').getValue();
  murInfos.sheetID = murSheet.getSheetId().toString();
  return murInfos;
}

function resetMur()
{
  var confirm = Browser.msgBox('ATTENTION: tous les onglets de murs vont être supprimées. Voulez vraiment continuer?', Browser.Buttons.YES_NO); 
  if(confirm!='yes'){return};// if user click NO then exit the function, else move data
  
  var murSheetList = getMurSheetList();
  for(var i=0; i<murSheetList.length; i++)
  {
    Logger.log(murSheetList[i]);
    var sheet = ss.getSheetByName(murSheetList[i]);
    ss.deleteSheet(sheet);
  }
  
  // Clear in Recap list:
  sheetRecapMurListRange.clear({contentsOnly: true});
}

function createMurs()
{
  var murList = getMurList();
  
  for(var i=0; i<murList.length; i++)
  {
    var mur = murList[i];
    createMur(mur.pseudo, mur.x, mur.y, mur.hour, mur.deff, mur.spy);
  }
  
  var murSheetList = getMurSheetList();
  for(var i=0; i<murSheetList.length; i++)
  {
    sheetRecap.getRange(4+i, 2).setFormula('=HYPERLINK("#gid='+ ss.getSheetByName(murSheetList[i]).getSheetId() + '";"' + murSheetList[i] + '")');
  }
  
  /*murSheetListInverted = murSheetList.map(function(e){return [e];});
  //getRange(row, column, numRows, numColumns) 
  var range = sheetRecap.getRange(4, 2, murSheetListInverted.length, 1);
  range.setValues(murSheetListInverted);*/
}

function getMurSheetList()
{
  var murSheetList = [];
  
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++)
  {
    var name = sheets[i].getName();
    //Logger.log(name + " exist?");
    if(name.match("M_"))
    {
      murSheetList.push(name);
    }
  }
  return murSheetList;
}


function getMurList() {
  var murList = [];
  
  var recensRange = sheetRecensement.getRange("D2:Y999");
  var recensValues = recensRange.getValues();
  
  for(var i=0; i<recensValues.length; i++)
  {
    if(recensValues[i][0] === '') // no more recens
      break;
    
    if(recensValues[i][20] === "MUR") // synchro, let's add it to list if not already in it
    {
      var mur = {};
      mur.x = recensValues[i][8]; // x
      mur.y = recensValues[i][9]; // y
      mur.pseudo = recensValues[i][7];
      mur.hour = recensValues[i][11]; // Hour
      mur.deff = recensValues[i][21]; // Troupes
      mur.spy = 0;
      
      murList.push(mur);
    }
  }
  return murList;
}

function doesMurExist(x, y, hour)
{
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++)
  {
    var name = sheets[i].getName();
    Logger.log(name + " exist?");
    if(name.match("M_"))
    {
      Logger.log(name + " match M_");
      var murSheet = ss.getSheetByName(name);
      Logger.log(murSheet.getRange('C4').getValue() + " === " + x + " ?");
      Logger.log(murSheet.getRange('D4').getValue() + " === " + y + " ?");
      /*Logger.log(murSheet.getRange('B4').getValue() + " === " + hour + " ?");
      
      var date1 = new Date(hour);
      var date2 = new Date(murSheet.getRange('B4').getValue());
      
      // getTime() returns the number of milliseconds since the beginning of
      // January 1, 1970 UTC.
      // True, as the dates represent the same moment in time.
      Logger.log(date1.getTime() == date2.getTime());*/
      
      if(murSheet.getRange('C4').getValue() === x && murSheet.getRange('D4').getValue() === y)
       return true;
    }
  }
  return false;
}

function createMur(pseudo, x, y, hour, deff, spy) 
{
  var name = "M_" + pseudo + "_1";
  if(!doesMurExist(x, y, hour))
  {
     var ss = SpreadsheetApp.getActiveSpreadsheet();
     var sheet = ss.getSheetByName('MUR TYPE').copyTo(ss);
  
     /* Before cloning the sheet, delete any previous copy */
     var m = ss.getSheetByName(name);
     var i = 2;
     while(m || i==10)
     {
       name = "M_" + pseudo + "_" + i;
       m = ss.getSheetByName(name);
       i++;
     }
  
     SpreadsheetApp.flush(); // Utilities.sleep(2000);
     sheet.setName(name);
  
     /* Make the new sheet active */
     ss.setActiveSheet(sheet);
  
     //========================================================
     // Fill infos:
  
     var murSheet = ss.getSheetByName(name);
     murSheet.getRange('B2').setValue(name);
     murSheet.getRange('C2').setValue(pseudo);
     murSheet.getRange('C4').setValue(x);
     murSheet.getRange('D4').setValue(y);
     murSheet.getRange('B4').setValue(hour);
     murSheet.getRange('F3').setValue(deff);
     murSheet.getRange('F6').setValue(spy);
  }
  else
  {
    Logger.log("Mur in " + x + " " + y + " at " + hour + " already exist.");
  }
}