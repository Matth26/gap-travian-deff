var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetRecensement = ss.getSheetByName('RECENSEMENT');

//=======================================================

function getSynchroList() {
  var synchroList = [];
  
  var recensRange = sheetRecensement.getRange("D2:Y999");
  var recensValues = recensRange.getValues();
  
  for(var i=0; i<recensValues.length; i++)
  {
    if(recensValues[i][0] === '') // no more recens
      break;
    
    if(recensValues[i][20] === "SYNCHRO") // synchro, let's add it to list if not already in it
    {
      var synchro = [, , , , , , , , , , ];
      synchro[0] = recensValues[i][0]; // ID
      synchro[1] = recensValues[i][8]; // x
      synchro[2] = recensValues[i][9]; // y
      synchro[3] = '';
      synchro[4] = recensValues[i][11]; // Hour
      synchro[5] = '';
      synchro[6] = recensValues[i][21]; // Troupes
      synchro[7] = 'Table Getter';
      synchro[8] = 'Voir Village';
      synchro[9] = '';  // Pseudo
      synchro[10] = ''; // Troupes quantity
      
      synchroList.push(synchro);
    }
  }
  
  if(synchroList.length > 0)
  {
    var synchroRange = ss.getSheetByName('SYNCHROS').getRange("K4:U999");
    var synchroValues = synchroRange.getValues();
    
    // Now let's push every synchro in SYNCHRO_2 tab:
    for(var i=0; i<synchroValues.length; i++)
    {
      if(synchroValues[i][0] === '') // no more synchro
        break;
      
      // Search if synchro already listed
      for(var j=0; j<synchroList.length; j++)
      {
        Logger.log(synchroList[j][0])
        Logger.log(synchroValues[i][0])
        if(synchroList[j][0] === synchroValues[i][0]) // if synchro exist, let's save it state (Pseudo and troupe already sent)
        {          
          synchroList[j][9] = synchroValues[i][9];   // Pseudo
          synchroList[j][10] = synchroValues[i][10]; // Troupes quantity
        }
      }
      
      // Logger.log(synchroValues[i])
      // [12.0, 45.0, 32.0, , Tue Dec 17 21:59:35 GMT+01:00 2019, , 4000.0, Table Getter, Voir Village, Merinos, 12222.0]
    }
  }

  ss.getSheetByName('SYNCHROS').getRange("K4:Q999").clear({contentsOnly: true});
  ss.getSheetByName('SYNCHROS').getRange("T4:U999").clear({contentsOnly: true});
  
  // Now remove all synchro and paste the synchro liste
  if(synchroList.length !== 0)
  {
    var synchroListBeforeLinks = [];
    for(var i=0; i<synchroList.length; i++)
      synchroListBeforeLinks.push(synchroList[i].slice(0, 7));
    ss.getSheetByName('SYNCHROS').getRange(4, 11, synchroList.length, 7).setValues(synchroListBeforeLinks);
    
    var synchroListAfterLinks = [];
    for(var i=0; i<synchroList.length; i++)
      synchroListAfterLinks.push(synchroList[i].slice(9, 11));
    ss.getSheetByName('SYNCHROS').getRange(4, 11+9, synchroList.length, 2).setValues(synchroListAfterLinks);
  }
}


function isSynchroInList(synchroList, synchro) {
  for(var i=0; i<synchroList.length; i++)
  {
    if(synchroList[i][0] === synchro[0]) // if ID are equal
      return true;
  }
  return false;
}
