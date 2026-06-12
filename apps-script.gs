// AquaCRM — Google Apps Script Backend v3
// Обновлено: очередь клиентов (Статус / КтоВзял), защита от одновременного звонка двух операторов

var CLIENTS_SHEET = 'Mijozlar';
var CALLS_SHEET   = 'Qongiroqlar';
var ORDERS_SHEET  = 'Buyurtmalar';

function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  var result;
  try {
    var data = {};
    if (e.postData) {
      try { data = JSON.parse(e.postData.contents); } catch(ex) { data = {}; }
    }
    var action = (data.action) || (e.parameter && e.parameter.action) || '';

    if      (action === 'getClients')     result = getClients();
    else if (action === 'updateClient')   result = updateClient(data);
    else if (action === 'logCall')        result = logCall(data);
    else if (action === 'getCalls')       result = getCalls(data.clientId);
    else if (action === 'getOrders')      result = getOrders(data.clientId, data.isPredoplata);
    else if (action === 'getStats')       result = getStats();
    else if (action === 'initSheets')     result = initSheets();
    else if (action === 'takeNextClient') result = takeNextClient(data.operatorName);
    else if (action === 'setClientStatus')result = setClientStatus(data.clientId, data.status, data.operatorName);
    else result = {error: 'Unknown action: ' + action};
  } catch(err) {
    result = {error: err.toString()};
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── INIT SHEETS ───────────────────────────────────────────────
function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Mijozlar
  var cs = ss.getSheetByName(CLIENTS_SHEET);
  if (!cs) {
    cs = ss.insertSheet(CLIENTS_SHEET);
    var hc = ['id','name','phone','phone2','addr','addrNorm','addrMaps','manager','debt','orders','totalSum','lastOrder','days','note','priority','wasCalled','predoplata','courier','lastCallDate','lastCallStatus'];
    cs.getRange(1,1,1,hc.length).setValues([hc]).setFontWeight('bold').setBackground('#1565C0').setFontColor('#fff');
    cs.setFrozenRows(1);
  }
  ensureStatusCols(cs);

  // Qongiroqlar
  var qs = ss.getSheetByName(CALLS_SHEET);
  if (!qs) {
    qs = ss.insertSheet(CALLS_SHEET);
    var hq = ['id','clientId','clientName','phone','operator','status','note','callDate','callTime'];
    qs.getRange(1,1,1,hq.length).setValues([hq]).setFontWeight('bold').setBackground('#1565C0').setFontColor('#fff');
    qs.setFrozenRows(1);
  }

  // Buyurtmalar — история заказов
  var os = ss.getSheetByName(ORDERS_SHEET);
  if (!os) {
    os = ss.insertSheet(ORDERS_SHEET);
    var ho = ['id','clientId','clientName','date','qty','price','sum','payMethod','operator','note'];
    os.getRange(1,1,1,ho.length).setValues([ho]).setFontWeight('bold').setBackground('#6A1B9A').setFontColor('#fff');
    os.setFrozenRows(1);
  }

  return {success: true, message: 'Sheets initialized!'};
}

// ── QUEUE COLUMNS (Статус / КтоВзял) ──────────────────────────
// Adds the two columns to Mijozlar if they don't exist yet and
// returns their 1-based column indexes.
function ensureStatusCols(sh) {
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var statusIdx = headers.indexOf('Статус');
  var takenByIdx = headers.indexOf('КтоВзял');

  if (statusIdx < 0) {
    lastCol++;
    sh.getRange(1, lastCol).setValue('Статус').setFontWeight('bold').setBackground('#1565C0').setFontColor('#fff');
    statusIdx = lastCol - 1;
  }
  if (takenByIdx < 0) {
    lastCol++;
    sh.getRange(1, lastCol).setValue('КтоВзял').setFontWeight('bold').setBackground('#1565C0').setFontColor('#fff');
    takenByIdx = lastCol - 1;
  }

  return {statusCol: statusIdx + 1, takenByCol: takenByIdx + 1};
}

// ── GET CLIENTS ───────────────────────────────────────────────
function getClients() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CLIENTS_SHEET);
  if (!sh) return {error: 'Sheet not found. Run initSheets() first.'};

  ensureStatusCols(sh);

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {clients: []};

  var headers = data[0];
  var clients = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    obj.status = obj['Статус'] || '';
    obj.takenBy = obj['КтоВзял'] || '';
    clients.push(obj);
  }
  return {clients: clients, total: clients.length};
}

// ── UPDATE CLIENT ─────────────────────────────────────────────
function updateClient(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CLIENTS_SHEET);
  if (!sh) return {error: 'Sheet not found'};

  var allData = sh.getDataRange().getValues();
  var headers = allData[0];
  var idIdx = headers.indexOf('id');

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idIdx]) === String(data.id)) {
      var fields = ['name','phone','phone2','addr','addrNorm','manager','note','predoplata','courier'];
      fields.forEach(function(field) {
        var idx = headers.indexOf(field);
        if (idx >= 0 && data[field] !== undefined) {
          sh.getRange(i+1, idx+1).setValue(data[field]);
        }
      });
      return {success: true, id: data.id};
    }
  }
  return {error: 'Client not found: ' + data.id};
}

// ── QUEUE: TAKE NEXT CLIENT ─────────────────────────────────────
// Atomically finds the first client with empty "Статус", marks it
// "Взят" with the operator name + timestamp in "КтоВзял", and
// returns that client. Uses a script lock so two operators can
// never receive the same client.
function takeNextClient(operatorName) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return {success: false, error: 'Server band, biroz kutib qayta urinib koring'};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CLIENTS_SHEET);
    if (!sh) return {error: 'Sheet not found'};

    var cols = ensureStatusCols(sh);
    var data = sh.getDataRange().getValues();
    var headers = data[0];

    var tz = 'Asia/Tashkent';
    var stamp = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm:ss');
    var takenByValue = (operatorName || 'Operator') + ' (' + stamp + ')';

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      if (row[cols.statusCol - 1]) continue; // already taken/processed

      // Re-read this exact cell right before writing — make sure no
      // other operator grabbed this row in the last few milliseconds.
      var freshStatus = sh.getRange(i+1, cols.statusCol).getValue();
      if (freshStatus) continue;

      sh.getRange(i+1, cols.statusCol).setValue('Взят');
      sh.getRange(i+1, cols.takenByCol).setValue(takenByValue);

      var obj = {};
      for (var j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
      obj.status = 'Взят';
      obj.takenBy = takenByValue;
      return {success: true, client: obj};
    }

    return {success: false, message: 'Navbatda bosh mijoz qolmadi'};
  } finally {
    lock.releaseLock();
  }
}

// ── QUEUE: SET CLIENT STATUS ────────────────────────────────────
// Changes a client's queue status after a call
// (Обзвонен / Заказ / Перезвонить).
function setClientStatus(clientId, status, operatorName) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return {error: 'Server band, biroz kutib qayta urinib koring'};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CLIENTS_SHEET);
    if (!sh) return {error: 'Sheet not found'};

    var cols = ensureStatusCols(sh);
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('id');

    var tz = 'Asia/Tashkent';
    var stamp = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm:ss');
    var takenByValue = (operatorName || 'Operator') + ' (' + stamp + ')';

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(clientId)) {
        sh.getRange(i+1, cols.statusCol).setValue(status || '');
        sh.getRange(i+1, cols.takenByCol).setValue(takenByValue);
        return {success: true, id: clientId, status: status, takenBy: takenByValue};
      }
    }
    return {error: 'Client not found: ' + clientId};
  } finally {
    lock.releaseLock();
  }
}

// ── LOG CALL ─────────────────────────────────────────────────
function logCall(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CALLS_SHEET);
  if (!sh) return {error: 'Calls sheet not found'};

  var now = new Date();
  var tz = 'Asia/Tashkent';
  var dateStr = Utilities.formatDate(now, tz, 'dd.MM.yyyy');
  var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
  var callId = now.getTime();

  sh.appendRow([callId, data.clientId, data.clientName||'', data.phone||'',
    data.operator||'Operator', data.status||'', data.note||'', dateStr, timeStr]);

  // Update lastCallDate in Mijozlar
  var cs = ss.getSheetByName(CLIENTS_SHEET);
  if (cs) {
    var cd = cs.getDataRange().getValues();
    var ch = cd[0];
    var idIdx = ch.indexOf('id');
    var lcdIdx = ch.indexOf('lastCallDate');
    var lcsIdx = ch.indexOf('lastCallStatus');

    // Add columns if missing
    if (lcdIdx < 0) {
      var nc = ch.length + 1;
      cs.getRange(1, nc).setValue('lastCallDate');
      cs.getRange(1, nc+1).setValue('lastCallStatus');
      lcdIdx = nc - 1; lcsIdx = nc;
    }

    for (var i = 1; i < cd.length; i++) {
      if (String(cd[i][idIdx]) === String(data.clientId)) {
        cs.getRange(i+1, lcdIdx+1).setValue(dateStr);
        cs.getRange(i+1, lcsIdx+1).setValue(data.status||'');
        break;
      }
    }
  }

  // If delivery — also log to Buyurtmalar
  if (data.status === 'delivery' && data.qty) {
    var os = ss.getSheetByName(ORDERS_SHEET);
    if (os) {
      os.appendRow([callId, data.clientId, data.clientName||'', dateStr,
        data.qty||1, data.price||13000, data.sum||0,
        data.payMethod||'cash', data.operator||'Operator', data.note||'']);
    }
  }

  return {success: true, callId: callId, date: dateStr, time: timeStr};
}

// ── GET CALLS ─────────────────────────────────────────────────
function getCalls(clientId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CALLS_SHEET);
  if (!sh) return {calls: []};

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {calls: []};

  var headers = data[0];
  var cidIdx = headers.indexOf('clientId');
  var calls = [];

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][cidIdx]) === String(clientId)) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
      calls.push(obj);
      if (calls.length >= 20) break;
    }
  }
  return {calls: calls};
}

// ── GET ORDERS ────────────────────────────────────────────────
function getOrders(clientId, isPredoplata) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ORDERS_SHEET);
  if (!sh) return {orders: []};

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {orders: []};

  var headers = data[0];
  var cidIdx = headers.indexOf('clientId');
  var orders = [];

  // Get all orders for this client (newest first)
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][cidIdx]) === String(clientId)) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
      orders.push(obj);
    }
  }

  // Limit: 30 for regular, all for predoplata
  var limit = isPredoplata ? orders.length : 30;
  return {orders: orders.slice(0, limit), total: orders.length};
}

// ── GET STATS ─────────────────────────────────────────────────
function getStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cs = ss.getSheetByName(CLIENTS_SHEET);
  var qs = ss.getSheetByName(CALLS_SHEET);

  var clients = cs ? Math.max(0, cs.getLastRow() - 1) : 0;
  var calls   = qs ? Math.max(0, qs.getLastRow() - 1) : 0;

  var now = new Date();
  var today = Utilities.formatDate(now, 'Asia/Tashkent', 'dd.MM.yyyy');
  var todayCalls = 0;

  if (qs && calls > 0) {
    var qData = qs.getDataRange().getValues();
    var headers = qData[0];
    var dateIdx = headers.indexOf('callDate');
    for (var i = 1; i < qData.length; i++) {
      if (qData[i][dateIdx] === today) todayCalls++;
    }
  }

  return {totalClients: clients, totalCalls: calls, todayCalls: todayCalls, date: today};
}
