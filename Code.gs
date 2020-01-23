function goToActiveSheetMur1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName('Mur 1  Rowrant'));
}

function goToActiveSheetMur2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName('mur 2'));
}

//========================================================================================================
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetRecensement = ss.getSheetByName('RECENSEMENT');

//========================================================================================================
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetRecap = ss.getSheetByName('Récapitulatif');

//=======================================================
var RecensRangeToBeReset = ss.getSheetByName('RECENSEMENT').getRange("E2:Y999");
var TravallyRangeToBeReset = ss.getSheetByName('Travally').getRange("A1:Z999");
var SynchroRangeToBeReset = ss.getSheetByName('SYNCHROS').getRange("K4:U999");

function resetAfterOP()
{
  var confirm = Browser.msgBox('ATTENTION: toutes les données de l\'opé vont être supprimées. Voulez vraiment continuer?', Browser.Buttons.YES_NO); 
  if(confirm!='yes'){return};// if user click NO then exit the function, else move data
  
  RecensRangeToBeReset.clear({contentsOnly: true});
  TravallyRangeToBeReset.clear();
  SynchroRangeToBeReset.clear({contentsOnly: true});
  resetMur();
}
//=======================================================

function onOpen() {
  Logger.log("onOpen")
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Analyse des Attaques')
  .addItem('ParseTravally', 'addAttackArrayToRecen')
  .addSeparator()
  .addItem('Update Synchros', 'getSynchroList')
  .addSeparator()
  .addItem('Update Murs', 'createMurs')
  .addSeparator()
  .addItem('Get Message Ally', 'createMessageAlly')
  .addSeparator()
  .addItem('Reset Mur','resetMur')
  .addSeparator()
  .addItem('Reset OP', 'resetAfterOP')
  /*.addSubMenu(ui.createMenu('Sub-menu')
              .addItem('Test ViviType', 'fillViviType'))*/
  .addToUi();
}

var start = 0;
var finish = 0;
var yourCell = ss.getSheetByName('Travally').getRange("Q1:Q999");
var yourCell2 = ss.getSheetByName('Travally').getRange("R1:R2");

function begin(){
  start = new Date().getTime();
}

function end(){
  finish = new Date().getTime();
  writeElapsed(start, finish)
}

function writeElapsed(start, finish){
  yourCell.setValue(finish - start); //will give elapsed time in ms
}

function writeIn(lol){
  yourCell2.setValue(lol);
}

function parseTravally()
{
  
  var attackArray = [];
  
  var sheetTravally = ss.getSheetByName('Travally');
  var range = sheetTravally.getRange("A1:O999");
  
  var rangeValues = range.getValues();
  for(var i=0; i<rangeValues.length; i++)
  {
    if(rangeValues[i][0] != "" && rangeValues[i][0] != "alliancesource") // line is not empty or not header line
    {
      if(sheetTravally.getRange(i+1, 1).getBackground() != "#bdbdbd") // if line has not already been parsed in previous script run
      {
        var attack = {};
        
        Logger.log("TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT");
        Logger.log(rangeValues[i]);
        // [Error404, Carptrak, 2.0, Carpe koï, 85.0, 18.0, 1.0, ‑ 37, 23.0, Malidom, 2.0, BYOP, 122.10241602851, Mon Sep 02 19:29:02 GMT+02:00 2019, NC]
        
        if(typeof(rangeValues[i][7]) === 'string')
          attack['xAttacked'] = parseInt(rangeValues[i][7].replace('‑', '-').replace(' ', ''), 10);
        else
          attack['xAttacked'] = rangeValues[i][7];
        
        if(typeof(rangeValues[i][8]) === 'string')
          attack['yAttacked'] = parseInt(rangeValues[i][8].replace('‑', '-').replace(' ', ''), 10);
        else
          attack['yAttacked'] = rangeValues[i][8];
        
        
        attack['pseudoAttacked'] = rangeValues[i][9];
        attack['allyAttacked'] = rangeValues[i][11];
        
        if(typeof(rangeValues[i][4]) === 'string')
          attack['xAttacking'] = parseInt(rangeValues[i][4].replace('‑', '-').replace(' ', ''), 10);
        else
          attack['xAttacking'] = rangeValues[i][4];
        
        if(typeof(rangeValues[i][5]) === 'string')
          attack['yAttacking'] = parseInt(rangeValues[i][5].replace('‑', '-').replace(' ', ''), 10);
        else
          attack['yAttacking'] = rangeValues[i][5];
        
        attack['pseudoAttacking'] = rangeValues[i][1];
        attack['allyAttacking'] = rangeValues[i][0];
        
        var impactHour = new Date(rangeValues[i][13]);
        attack['impactHour'] = impactHour;
        
        attack['waveNumber'] = 1;
        
        //attack = fillViviType(attack);
        
        Logger.log(attack)
        //begin();
        attackArray = addAttackOrIncrementWaveNumber(attackArray, attack)
        //end();
        //writeElapsed(start, finish);
        
        Logger.log(attackArray);
        Logger.log("TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT");
        
        // Color the computed line:
        sheetTravally.getRange(i+1, 1, 1, 20).setBackground("#bdbdbd");
      }
      
    }
  }
  return attackArray;
}

function fillViviType(attack)
{
  Logger.log("fillViviType")
  attack['viviType'] = "unknown";
  
  var sheetVivType = ss.getSheetByName('ViviType');
  
  // CAPI
  var isCapi = false;
  var range = sheetVivType.getRange("A2:E51");
  var rangeValues = range.getValues();
  for(var i=0; i<rangeValues.length; i++)
  {
    //Logger.log(rangeValues[i]);
    if(attack['xAttacked'] == rangeValues[i][1] && attack['yAttacked'] == rangeValues[i][2])
    {
      attack['viviType'] = "CAPI";//rangeValues[i][3];
      isCapi = true;
    }
  }
  
  // ARTE
  range = sheetVivType.getRange("G2:I9");
  rangeValues = range.getValues();
  for(var i=0; i<rangeValues.length; i++)
  {
    //Logger.log(rangeValues[i]);
    if(attack['xAttacked'] == rangeValues[i][0] && attack['yAttacked'] == rangeValues[i][1])
    {
      attack['viviType'] = rangeValues[i][2];
    }
  }
  
  // VOFF
  range = sheetVivType.getRange("K2:N26");
  rangeValues = range.getValues();
  for(var i=0; i<rangeValues.length; i++)
  {
    //Logger.log(rangeValues[i]);
    if(attack['xAttacked'] == rangeValues[i][1] && attack['yAttacked'] == rangeValues[i][2])
    {
      attack['viviType'] = "VOFF";//rangeValues[i][3];
      if(isCapi == true)
        attack['viviType'] = "VOFF/CAPI";
    }
  }
  
  return attack;
}

function getDeclaringHour(row) {
  var rangeValues = AColumn.getValues();
  //Logger.log(rangeValues);
  return rangeValues[row][0];
}

function addAttackArrayToRecen()
{
  Logger.log("addAttackArrayToRecen");
  var attackArray = parseTravally();
  Logger.log(attackArray);
  
  var row = getFirstEmptyRow();
  var alreadyRecensedAttacks = [];
  if(row > 2)
    alreadyRecensedAttacks = sheetRecensement.getRange(2, 5, row-2, 11 ).getValues(); // On commence en B2, quand il n'y a aucune attaque recensée row=2 // 11 is the size of array push in recensement
  
  for(var i=0; i<attackArray.length; i++)
  {
    row = addAttack(attackArray[i], row, alreadyRecensedAttacks);
  }
}


function addAttack(atk, r, alreadyRecensedAttacks) 
{
  atk = fillViviType(atk);
  Logger.log("^^^^^^^^^^^^^^^ addAttack ^^^^^^^^^^^^^^^^^^^^^");
  
  var array = ["", atk['pseudoAttacking'], atk['allyAttacking'], atk['xAttacking'], atk['yAttacking'], "P("+atk['xAttacking']+","+atk['yAttacking']+")", 
    atk['pseudoAttacked'], atk['xAttacked'], atk['yAttacked'], "P("+atk['xAttacked']+","+atk['yAttacked']+")", atk['impactHour'], atk['waveNumber'], "", atk['viviType']]
  
  var shouldAddAtck = true;
  Logger.log("==================================")    
  for(var i=0; i<alreadyRecensedAttacks.length; i++)
  {
    Logger.log(alreadyRecensedAttacks[i])
    Logger.log("===== ATCK %s %s %s", atk.xAttacking, atk.yAttacking, atk.impactHour.valueOf())
    Logger.log("===== ALREADY %s %s %s", alreadyRecensedAttacks[i][3], alreadyRecensedAttacks[i][4], alreadyRecensedAttacks[i][8].valueOf())
    if(atk.xAttacking == alreadyRecensedAttacks[i][3] && atk.yAttacking == alreadyRecensedAttacks[i][4] && atk.impactHour.valueOf() == alreadyRecensedAttacks[i][8].valueOf())
    {
      Logger.log("ATK IS EQUAL TO ALREADY RECENSED ONE");
      shouldAddAtck = false;
    }
  }
  Logger.log("==================================")  
    
  if(shouldAddAtck) 
  {
    Logger.log("NEW ATCK LET'S ADD IT");
    var range = sheetRecensement.getRange(r, 5, 1, array.length); 
    range.setValues([array]);
    r = r+1;
  }
  
  Logger.log("^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^");
  
  return r;
}

function addAttackOrIncrementWaveNumber(attackArray, attack)
{
  Logger.log("WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW")
  Logger.log(attackArray);
  Logger.log("addAttackOrIncrementWaveNumber")
  var arrayLength = attackArray.length;
  if(arrayLength == 0)
  {
    Logger.log("attackArray empty")
    attackArray.push(attack);
  }
  else
  {
    var shouldPushAtk = true;
    for(var i = 0; i<arrayLength; i++)
    {
      if(attack.xAttacking == attackArray[i].xAttacking && attack.yAttacking == attackArray[i].yAttacking 
         && attack.impactHour.valueOf() == attackArray[i].impactHour.valueOf() 
         && attack.xAttacked == attackArray[i].xAttacked && attack.yAttacked == attackArray[i].yAttacked )
      {
        Logger.log("1) %s %s", attack.xAttacking, attackArray[i].xAttacking);
        Logger.log("2) %s %s", attack.yAttacking, attackArray[i].yAttacking);
        Logger.log("3) %s %s", attack.impactHour.valueOf(), attackArray[i].impactHour.valueOf());
        Logger.log("Attack same")
        attackArray[i].waveNumber++;
        shouldPushAtk = false;
        break; // if atk found, increment wave number en get out of the loop
      }
      else
      {
        Logger.log("Attack diff")
      }
    }
    if(shouldPushAtk)
      attackArray.push(attack); // if we are here, all attacks was diff, let's add it
  }
  Logger.log("WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW")
  return attackArray;
}

function getFirstEmptyRow() {
  var column = sheetRecensement.getRange('F:F');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .alert('You clicked the second menu item!');
}

function parsePR(copyPaste, declaringHour)
{
  Logger.log("AAAAAAAAAAAAAAAAAAAAAA");
  var attackedVillageName = "";
  var xAttacked = 0;
  var yAttacked = 0;
  var pseudoAttacked = "";
  
  var attackArray = []
  
  var t = copyPaste.toString().split("\n");
  //Logger.log(t);
  
  for(var j=0; j<t.length; j++)
  {
    t[j] = t[j].replace(/[^\x20-\xFF]+/g, '');//.replace(/[^ -~]+/g, "")
    //Logger.log(t[j]);
    
    var regExp = new RegExp(/Héros\s(.*)/);
    if(regExp.test(t[j]))
    {
      pseudoAttacked = regExp.exec(t[j])[1];
      Logger.log("======================");
      Logger.log(pseudoAttacked);
      Logger.log("======================");
    }
    
    // DELIMITER FOR ATTACKS:
    regExp = new RegExp(/Troupes en approche \((\d+)\)/);
    if(regExp.test(t[j]))
    {
      Logger.log("============================================");
      Logger.log("START PARSING ATTAQUE");
      nombreAttaque = parseInt(regExp.exec(t[j])[1], 10);
      Logger.log("Nombre d'attaques: %s", nombreAttaque);
      
      var isListeSorted = false;
      regExp = new RegExp(/.*première page page précédente.*/);
      if(regExp.test(t[j+1]))
      {
        j+=1;
      }
      
      // PARSE ATTAQUE ONE BY ONE:
      for(var a=0; a<nombreAttaque; a++)
      {
        Logger.log("ATTAQUE %s XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", a);
        var attack = {}
        attack['pseudoAttacked'] = pseudoAttacked;
        attack['waveNumber'] = 1;
        
        var firstLine = t[j+4*a+1];
        var secondLine = ""
        var thirdLine = ""
        var forthLine = ""
        
        
        
        secondLine = t[j+4*a+2].replace('−', '-').replace(/[^ -~]+/g, "");
        thirdLine = t[j+4*a+3];
        forthLine = t[j+4*a+4];
        
        
        Logger.log(firstLine);
        Logger.log(secondLine);
        Logger.log(thirdLine);
        Logger.log(forthLine);
        
        regExp = new RegExp(/(Marquer l'attaque)(.*)\s(pille|attaque)\s(.*)/);        
        if(regExp.test(firstLine))
        {
          attackedVillageName = regExp.exec(firstLine)[4];
          Logger.log("attackedVillageName %s", attackedVillageName);
          //attack['attackedVillageName'] = attackedVillageName;
          
          var attackType = regExp.exec(firstLine)[3];
          attack['attackType'] = attackType;
          Logger.log("Attaque type %s", attackType);
          
          var pseudoAttacking = regExp.exec(firstLine)[2];
          attack['pseudoAttacking'] = pseudoAttacking;
        }
        
        //  ‭(‭−‭46‬‬|‭119‬)‬ 
        regExp = new RegExp(/\((.*)\|(.*)\‬)/); //new RegExp(/\((−‭?\d+)\|(−‭?\d‭+)\)/);
        if(regExp.test(secondLine))
        {
          var xAttacking = parseInt(regExp.exec(secondLine)[1], 10);
          attack['xAttacking'] = xAttacking;
          
          var yAttacking = parseInt(regExp.exec(secondLine)[2], 10);
          attack['yAttacking'] = yAttacking;
        }
        
        // ? ? ? ? 
        
        // Arrivée:
        regExp = new RegExp(/dans (\d+):(\d{2}):(\d{2}) hà\s(\d{2}):(\d{2}):(\d{2})/); //new RegExp(/\((−‭?\d+)\|(−‭?\d‭+)\)/);
        if(regExp.test(forthLine))
        {
          var hours = parseInt(regExp.exec(forthLine)[1], 10);
          var minutes = parseInt(regExp.exec(forthLine)[2], 10);
          var secondes = parseInt(regExp.exec(forthLine)[3], 10);
          Logger.log("%s %s %s", hours, minutes, secondes);
          
          var now = new Date(Date.now());
          Logger.log(Utilities.formatDate(now, 'Etc/GMT', 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\''));
          var declaringHourObj = new Date(declaringHour);
          Logger.log(Utilities.formatDate(declaringHourObj, 'Etc/GMT', 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\''));
          var n = new Date(declaringHourObj.getTime() + 1000*60*60*hours + 1000*60*minutes + 1000*secondes)
          Logger.log(Utilities.formatDate(n, 'Etc/GMT', 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\''));
          Logger.log("Duration counter: %s", n);
          Logger.log(n.getTimezoneOffset());
          
          var impactHour = regExp.exec(forthLine)[4];
          var impactMinute = regExp.exec(forthLine)[5];
          var impactSeconde = regExp.exec(forthLine)[6];
          Logger.log("impactHour: " + impactHour);
          
          var date = new Date(n.getFullYear(), n.getMonth(), n.getDate(), impactHour, impactMinute, impactSeconde);
          //date.setMinutes(date.getMinutes() - n.getTimezoneOffset());
          Logger.log(Utilities.formatDate(date, 'Etc/GMT', 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\''));
          
          attack['impactHour'] = date;
          attack['travelHours'] = hours + minutes/60 + secondes/3600;
          
          
        }
        
        
        
        Logger.log("_________________________________");
        attack['declaringHour'] = declaringHour;
        Logger.log("MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM");
        Logger.log(attack)
        if(attack['attackType'] == "attaque")
        {
          
          attackArray = addAttackOrIncrementWaveNumber(attackArray, attack)
        }
        Logger.log(attackArray);
        Logger.log("MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM");
        //attackArray.push(attack);
        Logger.log("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
      }
      j+=4*nombreAttaque
      Logger.log("============================================");
    }; 
    
    regExp = new RegExp(/(\d+)\/\d+/);
    if(regExp.test(t[j]))
    {
      var regExpVillages = new RegExp(/Villages/);
      if(regExpVillages.test(t[j+1]))
      {
        var villageNumber = parseInt(regExp.exec(t[j])[1], 10);
        //Logger.log("dddddddd" + " " + villageNumber);
        
        for(var v=0; v<villageNumber; v++)
        {
          
          var villageName = t[j+4+3*v]
          //Logger.log(villageName);
          //Logger.log(attackedVillageName);
          if(villageName === attackedVillageName)
          {
            Logger.log(villageName);
            
            regExp = new RegExp(/\((.*)\|(.*)\‬)/); //new RegExp(/\((−‭?\d+)\|(−‭?\d‭+)\)/);
            var coor = t[j+4+3*v+1].replace('−', '-').replace(/[^ -~]+/g, "")
            //Logger.log(coor);
            if(regExp.test(coor))
            {
              xAttacked= parseInt(regExp.exec(coor)[1], 10);
              //Logger.log("------------------------");
              //Logger.log(xAttacked);
              
              yAttacked = parseInt(regExp.exec(coor)[2], 10);
              //Logger.log(yAttacked);
              //Logger.log("------------------------");
            }
          }
        }
      }
    }
    
  };
  
  // Fill each attack object with coordinates of the attackedVillage:
  for(var i=0; i<attackArray.length; i++)
  {
    attackArray[i]['xAttacked'] = xAttacked;
    attackArray[i]['yAttacked'] = yAttacked;
    attackArray[i]['pseudoAttacked'] = pseudoAttacked;
    
    var distance = Math.sqrt(Math.pow(attackArray[i]['xAttacking']-attackArray[i]['xAttacked'], 2) + Math.pow(attackArray[i]['yAttacking']-attackArray[i]['yAttacked'], 2));
    var speedMin = distance / attackArray[i]['travelHours']
    Logger.log("%s %s %s", distance, attackArray[i]['travelHours'], speedMin);
    attackArray[i]['speedMin'] = speedMin;
  }
  
  return attackArray;
}