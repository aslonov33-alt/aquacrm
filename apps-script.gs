// AquaCRM — Google Apps Script Backend v4
// Обновлено: заявки на доставку (Buyurtmalar: deliveryDate / status) хранятся на сервере

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
    else if (action === 'addDelivery')      result = addDelivery(data);
    else if (action === 'getDeliveries')    result = getDeliveries(data.deliveryDate, data.courier);
    else if (action === 'completeDelivery') result = completeDelivery(data.deliveryId, data.qty, data.price, data.sum, data.payMethod);
    else if (action === 'addCourierOrder')    result = addCourierOrder(data);
    else if (action === 'lookupClientByPhone') result = lookupClientByPhone(data.phone);
    else if (action === 'getAnalytics')        result = getAnalytics(data.month, data.year, data.courier);
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
  ensureDeliveryCols(os);

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
  var statusIdx = headers.indexOf('status');
  var orders = [];

  // Get all orders for this client (newest first), skip pending delivery requests
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][cidIdx]) === String(clientId)) {
      if (statusIdx >= 0 && data[i][statusIdx] === 'pending') continue;
      var obj = {};
      for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
      orders.push(obj);
    }
  }

  // Limit: 30 for regular, all for predoplata
  var limit = isPredoplata ? orders.length : 30;
  return {orders: orders.slice(0, limit), total: orders.length};
}

// ── DELIVERY COLUMNS (deliveryDate / status) ───────────────────
// Adds the two columns to Buyurtmalar if they don't exist yet and
// returns their 1-based column indexes.
function ensureDeliveryCols(sh) {
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var ddIdx = headers.indexOf('deliveryDate');
  var stIdx = headers.indexOf('status');
  var crIdx = headers.indexOf('courier');
  var srcIdx = headers.indexOf('source');
  var addrIdx = headers.indexOf('address');
  var phIdx = headers.indexOf('phone');

  if (ddIdx < 0) {
    lastCol++;
    sh.getRange(1, lastCol).setValue('deliveryDate').setFontWeight('bold').setBackground('#6A1B9A').setFontColor('#fff');
    ddIdx = lastCol - 1;
  }
  if (stIdx < 0) {
    lastCol++;
    sh.getRange(1, lastCol).setValue('status').setFontWeight('bold').setBackground('#6A1B9A').setFontColor('#fff');
    stIdx = lastCol - 1;
  }
  if (crIdx < 0) {
    lastCol++;
    sh.getRange(1, lastCol).setValue('courier').setFontWeight('bold').setBackground('#6A1B9A').setFontColor('#fff');
    crIdx = lastCol - 1;
  }
  if (srcIdx < 0) {
    lastCol++;
    sh.getRange(1, lastCol).setValue('source').setFontWeight('bold').setBackground('#6A1B9A').setFontColor('#fff');
    srcIdx = lastCol - 1;
  }
  if (addrIdx < 0) {
    lastCol++;
    sh.getRange(1, lastCol).setValue('address').setFontWeight('bold').setBackground('#6A1B9A').setFontColor('#fff');
    addrIdx = lastCol - 1;
  }
  if (phIdx < 0) {
    lastCol++;
    sh.getRange(1, lastCol).setValue('phone').setFontWeight('bold').setBackground('#6A1B9A').setFontColor('#fff');
    phIdx = lastCol - 1;
  }

  return {deliveryDateCol: ddIdx + 1, statusCol: stIdx + 1, courierCol: crIdx + 1,
          sourceCol: srcIdx + 1, addressCol: addrIdx + 1, phoneCol: phIdx + 1};
}

// Normalizes a courier-zone value for comparison: trims whitespace and
// uppercases, so "k1"/"K1"/" K1 " all compare equal as "K1".
function normCourier(v) {
  return String(v == null ? '' : v).trim().toUpperCase();
}

// Normalizes a Buyurtmalar deliveryDate cell to a 'dd.MM.yyyy' string.
// Sheets silently converts date-like strings written via appendRow into
// Date values, so getValues() may return a Date object instead of the
// original 'dd.MM.yyyy' string — convert it back for comparison/output.
function normDeliveryDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Tashkent', 'dd.MM.yyyy');
  }
  return String(v == null ? '' : v).trim();
}

// Looks up the "courier" zone of a client in Mijozlar by id.
function lookupClientCourier(clientId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CLIENTS_SHEET);
  if (!sh) return '';
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('id');
  var crIdx = headers.indexOf('courier');
  if (idIdx < 0 || crIdx < 0) return '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(clientId)) {
      return normCourier(data[i][crIdx]);
    }
  }
  return '';
}

// Extracts the last 9 digits of a phone number for comparison
// (handles "+998901234567", "998901234567", "901234567", spaces/dashes).
function last9Digits(phone) {
  var digits = String(phone == null ? '' : phone).replace(/\D/g, '');
  return digits.slice(-9);
}

// Looks up a client in Mijozlar by phone (last 9 digits, matches phone or phone2).
// Returns {found:true, clientId, name, addr} or {found:false}.
function lookupClientByPhone(phone) {
  var target = last9Digits(phone);
  if (target.length < 9) return {found: false};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CLIENTS_SHEET);
  if (!sh) return {found: false};

  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('id');
  var nameIdx = headers.indexOf('name');
  var phoneIdx = headers.indexOf('phone');
  var phone2Idx = headers.indexOf('phone2');
  var addrIdx = headers.indexOf('addr');
  var addrNormIdx = headers.indexOf('addrNorm');

  for (var i = 1; i < data.length; i++) {
    if (!data[i][idIdx]) continue;
    if (last9Digits(data[i][phoneIdx]) === target ||
        (phone2Idx >= 0 && last9Digits(data[i][phone2Idx]) === target)) {
      var addr = (addrNormIdx >= 0 && data[i][addrNormIdx]) ? data[i][addrNormIdx] : (addrIdx >= 0 ? data[i][addrIdx] : '');
      return {found: true, clientId: data[i][idIdx], name: data[i][nameIdx] || '', addr: addr || ''};
    }
  }
  return {found: false};
}

// ── DELIVERY: ADD ────────────────────────────────────────────
// Adds a new pending delivery request to Buyurtmalar.
function addDelivery(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return {error: 'Server band, biroz kutib qayta urinib koring'};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ORDERS_SHEET);
    if (!sh) return {error: 'Buyurtmalar sheet not found'};

    ensureDeliveryCols(sh);
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    var tz = 'Asia/Tashkent';
    var dateStr = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy');
    var deliveryId = new Date().getTime();

    var row = new Array(headers.length).fill('');
    var set = function(name, val) {
      var idx = headers.indexOf(name);
      if (idx >= 0) row[idx] = val;
    };
    set('id', deliveryId);
    set('clientId', data.clientId);
    set('clientName', data.clientName || '');
    set('date', dateStr);
    set('qty', data.qty || 1);
    set('operator', data.operator || 'Operator');
    set('deliveryDate', data.deliveryDate || dateStr);
    set('status', 'pending');
    set('courier', lookupClientCourier(data.clientId));

    sh.appendRow(row);
    return {success: true, id: deliveryId};
  } finally {
    lock.releaseLock();
  }
}

// ── COURIER: ADD OWN ORDER ────────────────────────────────────
// Lets a courier log an order a client called in to them directly,
// bypassing the operator. Same Buyurtmalar row shape as addDelivery,
// plus address/phone (for clients not yet in Mijozlar) and source='courier'.
function addCourierOrder(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return {error: 'Server band, biroz kutib qayta urinib koring'};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ORDERS_SHEET);
    if (!sh) return {error: 'Buyurtmalar sheet not found'};

    ensureDeliveryCols(sh);
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    var tz = 'Asia/Tashkent';
    var dateStr = data.deliveryDate || Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy');
    var deliveryId = new Date().getTime();
    var qty = parseInt(data.qty, 10) || 1;
    var price = parseInt(data.price, 10) || 0;
    var lookup = lookupClientByPhone(data.phone);

    var row = new Array(headers.length).fill('');
    var set = function(name, val) {
      var idx = headers.indexOf(name);
      if (idx >= 0) row[idx] = val;
    };
    set('id', deliveryId);
    set('clientId', lookup.found ? lookup.clientId : '');
    set('clientName', data.clientName || '');
    set('date', dateStr);
    set('qty', qty);
    set('price', price);
    set('sum', qty * price);
    set('payMethod', data.payMethod || 'cash');
    set('operator', data.operator || data.courier || 'Kuryer');
    set('note', data.note || '');
    set('deliveryDate', dateStr);
    set('status', 'pending');
    set('courier', normCourier(data.courier));
    set('source', 'courier');
    set('address', data.address || '');
    set('phone', data.phone || '');

    sh.appendRow(row);
    return {success: true, id: deliveryId, clientId: lookup.found ? lookup.clientId : '', clientFound: lookup.found};
  } finally {
    lock.releaseLock();
  }
}

// ── DELIVERY: GET PENDING ────────────────────────────────────
// Returns pending deliveries, optionally filtered by deliveryDate and/or
// courier zone. If courier is empty (operator/admin), all pending
// deliveries are returned. Deliveries with no courier zone set (legacy
// rows) are visible to every courier.
function getDeliveries(deliveryDate, courier) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ORDERS_SHEET);
  if (!sh) return {deliveries: []};

  ensureDeliveryCols(sh);

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {deliveries: []};

  var headers = data[0];
  var statusIdx = headers.indexOf('status');
  var ddIdx = headers.indexOf('deliveryDate');
  var crIdx = headers.indexOf('courier');

  var courierFilter = normCourier(courier);

  var deliveries = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    if (row[statusIdx] !== 'pending') continue;
    var rowDeliveryDate = normDeliveryDate(row[ddIdx]);
    if (deliveryDate && rowDeliveryDate !== String(deliveryDate)) continue;
    if (courierFilter) {
      var rowCourier = normCourier(row[crIdx]);
      if (rowCourier && rowCourier !== courierFilter) continue;
    }
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
    obj.deliveryDate = rowDeliveryDate;
    deliveries.push(obj);
  }
  return {deliveries: deliveries, total: deliveries.length};
}

// ── DELIVERY: COMPLETE ────────────────────────────────────────
// Marks a delivery as done, optionally updating the final qty/price/sum/payMethod.
function completeDelivery(deliveryId, qty, price, sum, payMethod) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return {error: 'Server band, biroz kutib qayta urinib koring'};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ORDERS_SHEET);
    if (!sh) return {error: 'Buyurtmalar sheet not found'};

    ensureDeliveryCols(sh);
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('id');
    var qtyIdx = headers.indexOf('qty');
    var priceIdx = headers.indexOf('price');
    var sumIdx = headers.indexOf('sum');
    var payIdx = headers.indexOf('payMethod');
    var statusIdx = headers.indexOf('status');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(deliveryId)) {
        if (qty !== undefined)       sh.getRange(i+1, qtyIdx+1).setValue(qty);
        if (price !== undefined)     sh.getRange(i+1, priceIdx+1).setValue(price);
        if (sum !== undefined)       sh.getRange(i+1, sumIdx+1).setValue(sum);
        if (payMethod !== undefined) sh.getRange(i+1, payIdx+1).setValue(payMethod);
        sh.getRange(i+1, statusIdx+1).setValue('done');
        return {success: true, id: deliveryId};
      }
    }
    return {error: 'Delivery not found: ' + deliveryId};
  } finally {
    lock.releaseLock();
  }
}

// ── DELIVERY: FIX COURIER ZONES (one-off утилита) ──────────────
// Перезаписывает колонку "courier" у всех pending заявок в Buyurtmalar
// значением из Mijozlar (по clientId). Запускать вручную из редактора
// Apps Script: выбрать fixDeliveryCouriers в выпадающем списке и
// нажать Run — деплой не нужен. Результат смотреть в "Выполнения"/Logger.
function fixDeliveryCouriers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ORDERS_SHEET);
  if (!sh) return;

  ensureDeliveryCols(sh);
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var cidIdx = headers.indexOf('clientId');
  var crIdx = headers.indexOf('courier');
  var statusIdx = headers.indexOf('status');

  var fixed = 0, noZone = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][statusIdx] !== 'pending') continue;
    var zone = lookupClientCourier(data[i][cidIdx]);
    if (!zone) { noZone++; continue; }
    if (normCourier(data[i][crIdx]) !== zone) {
      sh.getRange(i+1, crIdx+1).setValue(zone);
      fixed++;
    }
  }
  Logger.log('Tuzatildi: ' + fixed + ', Mijozda zona yoq: ' + noZone);
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

// ── ANALYTICS (boss) ──────────────────────────────────────────
// Aggregates Buyurtmalar for the given month/year, optionally filtered by
// courier zone ('K1'/'K2'/...). Dates are stored as 'dd.MM.yyyy' text
// (normDeliveryDate handles rows Sheets converted to Date objects).
// Rows with empty 'status' are legacy entries — counted as 'done'.
// Rows with empty 'source' are legacy entries — counted as 'operator'.
function getAnalytics(month, year, courier) {
  var empty = {totalOrders: 0, totalRevenue: 0, totalBottles: 0, doneCount: 0, pendingCount: 0,
                byDay: [], byCourier: [], bySource: {operator: 0, courier: 0}};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ORDERS_SHEET);
  if (!sh) return empty;

  ensureDeliveryCols(sh);
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return empty;

  var headers = data[0];
  var dateIdx = headers.indexOf('date');
  var qtyIdx = headers.indexOf('qty');
  var sumIdx = headers.indexOf('sum');
  var statusIdx = headers.indexOf('status');
  var courierIdx = headers.indexOf('courier');
  var sourceIdx = headers.indexOf('source');

  var mm = month ? (Number(month) < 10 ? '0' + Number(month) : '' + Number(month)) : '';
  var yy = year ? String(year) : '';
  var courierFilter = courier ? normCourier(courier) : '';

  var totalOrders = 0, totalRevenue = 0, totalBottles = 0, doneCount = 0, pendingCount = 0;
  var dayMap = {}, courierMap = {}, sourceCount = {operator: 0, courier: 0};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var dateStr = normDeliveryDate(row[dateIdx]);
    var parts = dateStr.split('.');
    if (parts.length !== 3) continue;
    if (mm && parts[1] !== mm) continue;
    if (yy && parts[2] !== yy) continue;

    var rowCourier = normCourier(row[courierIdx]);
    if (courierFilter && rowCourier !== courierFilter) continue;

    var qty = parseInt(row[qtyIdx], 10) || 0;
    var sum = parseInt(row[sumIdx], 10) || 0;
    var status = row[statusIdx] || 'done';
    var source = String(row[sourceIdx] || '').trim().toLowerCase() || 'operator';

    totalOrders++;
    totalRevenue += sum;
    totalBottles += qty;
    if (status === 'pending') pendingCount++; else doneCount++;
    if (source === 'courier') sourceCount.courier++; else sourceCount.operator++;

    if (!dayMap[dateStr]) dayMap[dateStr] = {date: dateStr, orders: 0, bottles: 0, revenue: 0, done: 0, pending: 0};
    dayMap[dateStr].orders++;
    dayMap[dateStr].bottles += qty;
    dayMap[dateStr].revenue += sum;
    if (status === 'pending') dayMap[dateStr].pending++; else dayMap[dateStr].done++;

    if (!courierMap[rowCourier]) courierMap[rowCourier] = {courier: rowCourier, orders: 0, bottles: 0, revenue: 0};
    courierMap[rowCourier].orders++;
    courierMap[rowCourier].bottles += qty;
    courierMap[rowCourier].revenue += sum;
  }

  var byDay = [];
  for (var dKey in dayMap) byDay.push(dayMap[dKey]);
  byDay.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  var daysCount = byDay.length || 1;
  var byCourier = [];
  for (var cKey in courierMap) {
    var cm = courierMap[cKey];
    cm.avgPerDay = Math.round(cm.revenue / daysCount);
    byCourier.push(cm);
  }
  byCourier.sort(function(a, b) { return b.revenue - a.revenue; });

  return {totalOrders: totalOrders, totalRevenue: totalRevenue, totalBottles: totalBottles,
          doneCount: doneCount, pendingCount: pendingCount,
          byDay: byDay, byCourier: byCourier, bySource: sourceCount};
}

// ── ONE-OFF: REMOVE DUPLICATE PENDING DELIVERIES ───────────────
// Run MANUALLY from the Apps Script editor (select removeDuplicateDeliveries
// in the function dropdown next to "Run" — no deploy needed, not an API action).
// Finds groups of Buyurtmalar rows with the same clientId + deliveryDate and
// status='pending', keeps the row with the smallest id in each group and
// deletes the rest. Rows with status='done' or empty clientId are never
// touched. Logs a preview (groups found / rows to delete) before deleting,
// then logs how many rows were actually removed — check "Выполнения"
// (Executions) for the output. Google Sheets keeps version history
// (Файл → История версий) if you need to undo.
function removeDuplicateDeliveries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ORDERS_SHEET);
  if (!sh) { Logger.log('Buyurtmalar sheet topilmadi'); return; }

  ensureDeliveryCols(sh);
  var data = sh.getDataRange().getValues();
  if (data.length < 2) { Logger.log('Buyurtmalar bosh'); return; }

  var headers = data[0];
  var idIdx = headers.indexOf('id');
  var clientIdIdx = headers.indexOf('clientId');
  var statusIdx = headers.indexOf('status');
  var dateIdx = headers.indexOf('deliveryDate');

  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[idIdx]) continue;
    if (row[statusIdx] !== 'pending') continue;
    var clientId = row[clientIdIdx];
    if (clientId === '' || clientId === null || clientId === undefined) continue;
    var dateStr = normDeliveryDate(row[dateIdx]);
    var key = String(clientId) + '|' + dateStr;
    if (!groups[key]) groups[key] = [];
    groups[key].push({rowIndex: i + 1, id: row[idIdx]});
  }

  var toDelete = [];
  var dupGroups = 0;
  for (var key in groups) {
    var items = groups[key];
    if (items.length < 2) continue;
    dupGroups++;
    items.sort(function(a, b) { return a.id - b.id; });
    for (var k = 1; k < items.length; k++) toDelete.push(items[k]);
  }

  Logger.log('Dublikat guruhlari: ' + dupGroups + ', ochirish uchun qatorlar: ' + toDelete.length);
  if (!toDelete.length) { Logger.log('Ochirish kerak emas.'); return; }

  toDelete.sort(function(a, b) { return b.rowIndex - a.rowIndex; });
  for (var m = 0; m < toDelete.length; m++) {
    sh.deleteRow(toDelete[m].rowIndex);
  }

  Logger.log('Ochirildi: ' + toDelete.length + ' qator');
}
