// ── JJ Loyalty Points — Google Apps Script (Updated with Staff column) ──

var SHEET_NAME = 'Sheet1';

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var result;
  try {
    var action = e.parameter.action;
    if (action === 'getAll') {
      result = getAllCustomers();
    } else if (action === 'save') {
      var data = JSON.parse(e.parameter.data);
      result = saveAllCustomers(data);
    } else {
      result = { success: false, error: 'Unknown action' };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAllCustomers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, customers: [] };

  // Now reading 10 columns (added registeredBy in column J)
  var values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var customers = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    var customer = {
      id:               String(row[0]),
      name:             String(row[1]),
      phone:            String(row[2]),
      points:           Number(row[3]) || 0,
      totalEarned:      Number(row[4]) || 0,
      totalRedeemed:    Number(row[5]) || 0,
      membershipJoined: String(row[6]),
      membershipExpiry: String(row[7]),
      created:          String(row[8]),
      registeredBy:     String(row[9]) || '',
      history:          []
    };
    customer.history = getHistory(customer.id);
    customers.push(customer);
  }
  return { success: true, customers: customers };
}

function saveAllCustomers(customers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  // Clear all data except header
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).clearContent();
  }

  for (var i = 0; i < customers.length; i++) {
    var c = customers[i];
    // Write 10 columns including registeredBy
    sheet.getRange(i + 2, 1, 1, 10).setValues([[
      c.id,
      c.name,
      c.phone          || '',
      c.points         || 0,
      c.totalEarned    || 0,
      c.totalRedeemed  || 0,
      c.membershipJoined || '',
      c.membershipExpiry || '',
      c.created        || '',
      c.registeredBy   || ''
    ]]);
    saveHistory(c.id, c.history || []);
  }

  formatSheet(sheet);
  return { success: true };
}

function getHistory(customerId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName('History');
  if (!histSheet) return [];
  var lastRow = histSheet.getLastRow();
  if (lastRow < 2) return [];
  var values = histSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var history = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (String(row[0]) === customerId) {
      history.push({
        type:   String(row[1]),
        amount: Number(row[2]) || 0,
        note:   String(row[3]),
        date:   String(row[4])
      });
    }
  }
  return history;
}

function saveHistory(customerId, historyArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName('History');
  if (!histSheet) {
    histSheet = ss.insertSheet('History');
    histSheet.getRange(1, 1, 1, 5).setValues([['customerId','type','amount','note','date']]);
    histSheet.hideSheet();
  }
  var lastRow = histSheet.getLastRow();
  if (lastRow >= 2) {
    var values = histSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = values.length - 1; i >= 0; i--) {
      if (String(values[i][0]) === customerId) histSheet.deleteRow(i + 2);
    }
  }
  for (var j = 0; j < historyArray.length; j++) {
    var h = historyArray[j];
    histSheet.appendRow([customerId, h.type, h.amount || 0, h.note || '', h.date || '']);
  }
}

function formatSheet(sheet) {
  try {
    // Update header row to include registeredBy column
    sheet.getRange(1, 1, 1, 10).setValues([[
      'id', 'name', 'phone', 'points', 'totalEarned',
      'totalRedeemed', 'membershipJoined', 'membershipExpiry', 'created', 'registeredBy'
    ]]);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1a1f2e').setFontColor('#d69e2e');
    for (var col = 1; col <= 10; col++) sheet.autoResizeColumn(col);
  } catch(e) {}
}
