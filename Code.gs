// ════════════════════════════════════════════════════════════════
//  StockNote — Google Apps Script (전체 완성본)
//  시트 구조:
//    "stocks"   : id, name, code, market, sector, rating, holding,
//                 targetPrice, stopLoss, tags, memo, createdAt
//    "trades"   : id, stockName, acctName, type, date, price,
//                 quantity, avgBuyPrice, reason, emotion, memo, createdAt
//    "notes"    : id, stockName, title, content, rating, createdAt
//    "accounts" : id, name, createdAt
// ════════════════════════════════════════════════════════════════

const SS  = SpreadsheetApp.getActiveSpreadsheet();

// ── 시트 헤더 정의 ────────────────────────────────────────────
const HEADERS = {
  stocks:   ['id','name','code','market','sector','rating','holding',
             'targetPrice','stopLoss','tags','memo','createdAt'],
  trades:   ['id','stockName','acctName','type','date','price',
             'quantity','avgBuyPrice','reason','emotion','memo','createdAt'],
  notes:    ['id','stockName','title','content','rating','createdAt'],
  accounts: ['id','name','createdAt'],
};

// ── 시트 가져오기 (없으면 자동 생성) ────────────────────────────
function getSheet(name) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.appendRow(HEADERS[name]);
    sheet.getRange(1, 1, 1, HEADERS[name].length)
      .setFontWeight('bold')
      .setBackground('#2563eb')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── 시트 → JSON 배열 변환 ────────────────────────────────────
function sheetToJson(name) {
  const sheet  = getSheet(name);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] === '' ? '' : String(row[i]); });
    return obj;
  }).filter(r => r.id); // id 없는 빈 행 제외
}

// ── ID로 행 번호 찾기 (1-based, 헤더 포함) ────────────────────
function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

// ── 새 ID 생성 ────────────────────────────────────────────────
function newId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
}

// ── CORS 응답 헬퍼 ────────────────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
//  GET 핸들러
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';
    switch (action) {
      case 'getStocks':   return jsonResponse(sheetToJson('stocks'));
      case 'getTrades':   return jsonResponse(sheetToJson('trades'));
      case 'getNotes':    return jsonResponse(sheetToJson('notes'));
      case 'getAccounts': return jsonResponse(sheetToJson('accounts'));
      default:
        return jsonResponse({ ok: true, version: '2.0', message: 'StockNote API' });
    }
  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

// ════════════════════════════════════════════════════════════════
//  POST 핸들러
// ════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    const data   = body.data || {};
    const id     = body.id;

    switch (action) {

      // ── 관심종목 ───────────────────────────────────────────
      case 'addStock': {
        const sheet = getSheet('stocks');
        const newId_ = data.id || newId();
        const row = HEADERS.stocks.map(h =>
          h === 'id' ? newId_ :
          h === 'createdAt' ? new Date().toISOString() :
          (data[h] !== undefined ? data[h] : '')
        );
        sheet.appendRow(row);
        return jsonResponse({ ok: true, id: newId_ });
      }
      case 'updateStock': {
        const sheet = getSheet('stocks');
        const rowNum = findRowById(sheet, id);
        if (rowNum < 0) return jsonResponse({ error: 'not found' });
        HEADERS.stocks.forEach((h, i) => {
          if (h !== 'id' && h !== 'createdAt' && data[h] !== undefined) {
            sheet.getRange(rowNum, i + 1).setValue(data[h]);
          }
        });
        return jsonResponse({ ok: true });
      }
      case 'deleteStock': {
        const sheet = getSheet('stocks');
        const rowNum = findRowById(sheet, id);
        if (rowNum > 0) sheet.deleteRow(rowNum);
        return jsonResponse({ ok: true });
      }

      // ── 매매일지 ───────────────────────────────────────────
      case 'addTrade': {
        const sheet = getSheet('trades');
        const newId_ = data.id || newId();
        const row = HEADERS.trades.map(h =>
          h === 'id' ? newId_ :
          h === 'createdAt' ? new Date().toISOString() :
          (data[h] !== undefined ? data[h] : '')
        );
        sheet.appendRow(row);
        return jsonResponse({ ok: true, id: newId_ });
      }
      case 'updateTrade': {
        const sheet = getSheet('trades');
        const rowNum = findRowById(sheet, id);
        if (rowNum < 0) return jsonResponse({ error: 'not found' });
        HEADERS.trades.forEach((h, i) => {
          if (h !== 'id' && h !== 'createdAt' && data[h] !== undefined) {
            sheet.getRange(rowNum, i + 1).setValue(data[h]);
          }
        });
        return jsonResponse({ ok: true });
      }
      case 'deleteTrade': {
        const sheet = getSheet('trades');
        const rowNum = findRowById(sheet, id);
        if (rowNum > 0) sheet.deleteRow(rowNum);
        return jsonResponse({ ok: true });
      }

      // ── 종목노트 ───────────────────────────────────────────
      case 'addNote': {
        const sheet = getSheet('notes');
        const newId_ = data.id || newId();
        const row = HEADERS.notes.map(h =>
          h === 'id' ? newId_ :
          h === 'createdAt' ? new Date().toISOString() :
          (data[h] !== undefined ? data[h] : '')
        );
        sheet.appendRow(row);
        return jsonResponse({ ok: true, id: newId_ });
      }
      case 'updateNote': {
        const sheet = getSheet('notes');
        const rowNum = findRowById(sheet, id);
        if (rowNum < 0) return jsonResponse({ error: 'not found' });
        HEADERS.notes.forEach((h, i) => {
          if (h !== 'id' && h !== 'createdAt' && data[h] !== undefined) {
            sheet.getRange(rowNum, i + 1).setValue(data[h]);
          }
        });
        return jsonResponse({ ok: true });
      }
      case 'deleteNote': {
        const sheet = getSheet('notes');
        const rowNum = findRowById(sheet, id);
        if (rowNum > 0) sheet.deleteRow(rowNum);
        return jsonResponse({ ok: true });
      }

      // ── 계좌 관리 ──────────────────────────────────────────
      case 'addAccount': {
        const sheet = getSheet('accounts');
        const newId_ = data.id || newId();
        const row = HEADERS.accounts.map(h =>
          h === 'id' ? newId_ :
          h === 'createdAt' ? new Date().toISOString() :
          (data[h] !== undefined ? data[h] : '')
        );
        sheet.appendRow(row);
        return jsonResponse({ ok: true, id: newId_ });
      }
      case 'updateAccount': {
        const sheet = getSheet('accounts');
        const rowNum = findRowById(sheet, id);
        if (rowNum < 0) return jsonResponse({ error: 'not found' });
        HEADERS.accounts.forEach((h, i) => {
          if (h !== 'id' && h !== 'createdAt' && data[h] !== undefined) {
            sheet.getRange(rowNum, i + 1).setValue(data[h]);
          }
        });
        return jsonResponse({ ok: true });
      }
      case 'deleteAccount': {
        const sheet = getSheet('accounts');
        const rowNum = findRowById(sheet, id);
        if (rowNum > 0) sheet.deleteRow(rowNum);
        return jsonResponse({ ok: true });
      }

      default:
        return jsonResponse({ error: 'Unknown action: ' + action });
    }
  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}
