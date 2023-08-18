function onOpen() {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Show Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Sidebar');
  DocumentApp.getUi().showSidebar(ui);
}

function getCharacterCount() {
  var body = DocumentApp.getActiveDocument().getBody();
  var text = body.getText();
  return text.length;
}

function getLastEditTime() {
  var lastEdit = DocumentApp.getActiveDocument().getLastEditTime();
  return lastEdit.toLocaleString();
}

function createOrUpdateGoogleSheet(sheetName, content) {
  var ss = getOrCreateSpreadsheet(sheetName);
  var sheet = ss.getActiveSheet();
  sheet.clearContents();
  var lines = content.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var rowData = lines[i].split('\t');
    sheet.appendRow(rowData);
  }
  var url = ss.getUrl();
  return url;
}

function getOrCreateSpreadsheet(sheetName) {
  var spreadsheet = null;
  try {
    spreadsheet = SpreadsheetApp.openByName(sheetName);
  } catch (e) {
    spreadsheet = SpreadsheetApp.create(sheetName);
  }
  return spreadsheet;
}

function createLocalTextFile(fileName, content) {
  var folder = DriveApp.getRootFolder();
  var files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    var file = files.next();
    file.setContent(content);
    return 'File updated: ' + file.getUrl();
  } else {
    var newFile = folder.createFile(fileName, content);
    return 'File created: ' + newFile.getUrl();
  }
}

function getSidebarContent() {
  var characterCount = getCharacterCount();
  var lastEditTime = getLastEditTime();
  var content = "Character Count: " + characterCount + "\nLast Edit Time: " + lastEditTime;
  return content;
}
