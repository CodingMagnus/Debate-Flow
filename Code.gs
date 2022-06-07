function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Debate')
    .addItem('Kick Current Flow', 'KICK')
    .addSeparator()
    .addItem('Add Case (Aff) Flows', 'ADDCASE')
    .addSeparator()
    .addItem('Add Off-Case (Neg) Flows', 'ADDOFF')
    .addSeparator()
    .addItem('Set Name', 'PROMPTINFO')
    .addToUi();
  SpreadsheetApp.getActiveSpreadsheet().getDeveloperMetadata()
}

// Functions used in script

function getNumOff() {
  return SpreadsheetApp.getActiveSpreadsheet().getRangeByName('_NUM_OFF_CASE');
}

function getNumAdvantages() {
  return SpreadsheetApp.getActiveSpreadsheet().getRangeByName('_NUM_ON_CASE');
}

function hide() {
  const sheet = SpreadsheetApp.getActiveSheet()
  sheet.hideSheet();
}


// Functions triggered by Debate menu

function ADDOFF() {
  const ui = SpreadsheetApp.getUi();

  const prevNumOff = getNumOff().getValue();

  const flowsToAdd = ui.prompt('# of flows');
  const neg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Neg');
  for (let i = 0, num = parseInt(flowsToAdd.getResponseText()); i < num; i++) {
    const newSheet = neg.copyTo(SpreadsheetApp.getActiveSpreadsheet());
    try {
      newSheet.setName(`N${prevNumOff + i + 1}`); // Sets name to N + its identity number i.e N1, N2, etc.
    } catch (err) {
      Logger.log(err);
    };
    newSheet.activate();
  };
  
  getNumOff().setValue(parseInt(prevNumOff) + parseInt(flowsToAdd.getResponseText()))
};

function ADDCASE() {
  const ui = SpreadsheetApp.getUi();

  const prevNumAdvs = getNumAdvantages().getValue();

  const flowsToAdd = ui.prompt('# of flows');
  const aff = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aff')
  for (let i = 0, num = parseInt(flowsToAdd.getResponseText()); i < num; i++) {
    const newSheet = aff.copyTo(SpreadsheetApp.getActiveSpreadsheet());
    try {
      newSheet.setName(`A${prevNumAdvs + i + 1}`); // Sets name to A + its identity number i.e N1, N2, etc.
    } catch (err) {
      Logger.log(err);
    };
    newSheet.activate();
  };
  getNumAdvantages().setValue(parseInt(prevNumAdvs) + parseInt(flowsToAdd.getResponseText()))
};

function SETNAMEFROMINFO(info) {
  //Logger.log(info)
  const nameFromInfo = `@${info.tournament} Round ${info.round} ${info.side} v. ${info.opponent}`
  Logger.log('renamed to ' + nameFromInfo)
  SpreadsheetApp.getActiveSpreadsheet().rename(nameFromInfo);
};

function PROMPTINFO() {
  const promptUiFile = HtmlService.createHtmlOutputFromFile('RenameUI');
  const ui = SpreadsheetApp.getUi();
  ui.showDialog(promptUiFile);
};


// Functions triggered by users

function ext(ref) {
  return '--> ' + ref;
}


// Functions triggered by buttons or other UI

function KICK() {
  const sheet = SpreadsheetApp.getActiveSheet(); // Current page i.e DA etc. etc.
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet(); // Full spreadsheet (with multiple pages)
  const oldName = sheet.getName();
  sheet.setName('>' + sheet.getName() + '<');
  spreadSheet.moveActiveSheet(spreadSheet.getNumSheets()) // Moves page to back
  sheet.setTabColor(null);
  return oldName;
}
