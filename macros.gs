function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('TPM')
      .addItem('Sync open Cert sheet', 'syncCertSheet')
      .addItem('Run Clear Rules', 'clearRules')
      .addToUi();
}

function onEdit() {
  setUpdated(SpreadsheetApp.getActiveSheet());
}

function setUpdated(sheet) {
  if (!sheet.getName().startsWith('Cert-')) return;
  sheet.getRange('W1')
    .setValue('Updated')
    .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).build())
    .setHorizontalAlignment('right');
  sheet.getRange('X1')
    .setValue(new Date())
    .setTextStyle(SpreadsheetApp.newTextStyle().setBold(false).build())
    .setHorizontalAlignment('center')
    .setNumberFormat('dd" "mmm" "hh":"mm');
}

function setSynced(sheet) {
  if (!sheet.getName().startsWith('Cert-')) return;
  sheet.getRange('W2')
    .setValue('Synced')
    .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).build())
    .setHorizontalAlignment('right');
  sheet.getRange('X2')
    .setValue(new Date())
    .setTextStyle(SpreadsheetApp.newTextStyle().setBold(false).build())
    .setHorizontalAlignment('center')
    .setNumberFormat('dd" "mmm" "hh":"mm');
}

function RemoveCellHighlights() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNameArray = [];
  var firstCertIdx = 6;

  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().startsWith("Cert-")) {
      ss.setActiveSheet(sheets[i]);
      var s = ss.getActiveSheet();
      s.showColumns(3, 4);
      s.getRange('4:6').setBackground(null);
      s.getRange('B4:G').setBackground(null);
    }
  }
};

function sortSheetsAsc() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNameArray = [];
  var firstCertIdx = 6;

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.startsWith("Cert-EM-")) {
      var newName = name.replace("Cert-EM-","Cert-");
      ss.getSheetByName(name).setName(newName);
      name = newName;
    }
    if (name.startsWith("Cert-")) {
      sheetNameArray.push(name);
    }
  }

  sheetNameArray.sort();

  for( var j = 0; j < sheets.length; j++ ) {
    ss.setActiveSheet(ss.getSheetByName(sheetNameArray[j]));
    ss.moveActiveSheet(j + firstCertIdx);
  }
}

function RemoveSingleHighlights() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().showColumns(3, 5);
  spreadsheet.getRange('4:6').setBackground(null);
  spreadsheet.getRange('B4:G').setBackground(null);
  spreadsheet.getActiveSheet().setTabColor('red');
};


function copyFormatting() {
  const spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getSheets()
  .filter(sheet => sheet.getName().startsWith("Cert-") && sheet.getName() != "Cert-A&I")
  .forEach(sheet => {
    console.log(sheet.getName());
    spreadsheet.setActiveSheet(sheet, true);
    spreadsheet.getRange('A4').activate();
    var currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    currentCell.activateAsCurrentCell();
    spreadsheet.getRange("'Cert-A&I\'!A4:A4").copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING, false);
  })
}

function sheetnames() { 
  var out = new Array()
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) out.push( [ sheets[i].getName() ] )
  return out  
}

function syncCertSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetName = sheet.getName();
  
  if(sheetName.indexOf("Cert-")>-1){
    var data = {
      'certType': 'EM',
      'gSheetId': '1xLf55Pv8OkYidRGQvc7CaGx62xqBX1wBSHdDcbWaF8g',
      'sheetName': sheetName
    };
    var options = {
      'method' : 'post',
      'contentType': 'application/json',
      'payload' : JSON.stringify(data)
    };
    var url = 'https://tpm.devflows.devfactory.com/10547_in1_prod_063_eng_tempo_rules';
    url = 'https://tpm.devflows.devfactory.com/12425_in1_prod_063_eng_tempo_rules_action';
    UrlFetchApp.fetch(url, options);

    var message = 'Certification Sync has been invoked for "' + sheetName + '".  This sync may take up to 10 minutes to complete, please be patient.';
    var title = '⏰ Tempo Certification ⏰';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 5);
    setSynced(sheet);
  } else {
    var message = 'Sheet name: ' + sheetName + ' is not a Certification Sheet!  Sync not invoked.';
    var title = '⚠️ Tempo Certification ⚠️';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 5);
  }
}

function clearRules() {
  
    var data = {
      'certType': 'EM',
      'gSheetId': '1xLf55Pv8OkYidRGQvc7CaGx62xqBX1wBSHdDcbWaF8g',
      'clearRulesOnly': true
    };
    var options = {
      'method' : 'post',
      'contentType': 'application/json',
      'payload' : JSON.stringify(data)
    };
    UrlFetchApp.fetch('https://tpm.devflows.devfactory.com/11400_in1_prod_067_invoke_tempo_cert_sync_ii', options);

    var message = 'Clear Rules automation has been invoked.  This process may take up to 5 minutes to complete.';
    var title = '⚠️ Clear Rules ⚠️';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 5);
}