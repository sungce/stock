// ═══════════════════════════════════════════════════════════
//  StockNote — Google Apps Script 백엔드  (Code.gs)
//  배포: 확장 프로그램 > Apps Script > 배포 > 웹 앱
//       실행 계정: 나 / 액세스 권한: 모든 사용자
// ═══════════════════════════════════════════════════════════

// ① 스프레드시트 ID를 여기에 붙여넣으세요
const SPREADSHEET_ID = '1x68B7wIl4yCu-56V-inmxQGUkOFveCn3xXiIie2TX98';

// 시트별 헤더 정의
const SHEET_HEADERS = {
  STOCKS:   ['id','code','name','sector','tags','targetPrice','stopLoss','rating','holding','memo','createdAt'],
  TRADES:   ['id','stockName','acctName','type','date','price','quantity','avgBuyPrice','reason','emotion','memo','createdAt'],
  NOTES:    ['id','stockId','stockName','content','risks','links','createdAt','updatedAt'],
  ACCOUNTS: ['id','name','createdAt'],
};

// 시트 색상
const SHEET_COLORS = {
  STOCKS: '#1e40af', TRADES: '#065f46', NOTES: '#7e22ce', ACCOUNTS: '#92400e'
};

// ───────────────────────────────────────────────
//  라우터 — GET (읽기)
// ───────────────────────────────────────────────
function doGet(e) {
  try {
    const p = e.parameter;
    let result;
    switch (p.action) {
      case 'getStocks':   result = sheetToObjects('STOCKS');   break;
      case 'getTrades':   result = sheetToObjects('TRADES');   break;
      case 'getNotes':    result = sheetToObjects('NOTES');    break;
      case 'getAccounts': result = sheetToObjects('ACCOUNTS'); break;
      case 'ping':        result = { ok: true, ts: new Date().toISOString() }; break;
      default:            result = { error: '알 수 없는 액션: ' + p.action };
    }
    return jsonRes(result);
  } catch (err) {
    return jsonRes({ error: err.toString() });
  }
}

// ───────────────────────────────────────────────
//  라우터 — POST (쓰기)
// ───────────────────────────────────────────────
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonRes({ error: '요청 파싱 오류: ' + err.toString() });
  }
  try {
    let result;
    switch (body.action) {
      // 관심종목
      case 'addStock':      result = addRow('STOCKS',   body.data);          break;
      case 'updateStock':   result = updateRow('STOCKS',   body.id, body.data); break;
      case 'deleteStock':   result = deleteRow('STOCKS',   body.id);         break;
      // 종목노트
      case 'addNote':       result = addRow('NOTES',    body.data);          break;
      case 'updateNote':    result = updateRow('NOTES',    body.id, body.data); break;
      case 'deleteNote':    result = deleteRow('NOTES',    body.id);         break;
      // 매매일지
      case 'addTrade':      result = addRow('TRADES',   body.data);          break;
      case 'updateTrade':   result = updateRow('TRADES',   body.id, body.data); break;
      case 'deleteTrade':   result = deleteRow('TRADES',   body.id);         break;
      // 계좌 관리
      case 'addAccount':    result = addRow('ACCOUNTS', body.data);          break;
      case 'updateAccount': result = updateRow('ACCOUNTS', body.id, body.data); break;
      case 'deleteAccount': result = deleteRow('ACCOUNTS', body.id);         break;
      default: result = { error: '알 수 없는 액션: ' + body.action };
    }
    return jsonRes(result);
  } catch (err) {
    return jsonRes({ error: err.toString() });
  }
}

// ───────────────────────────────────────────────
//  공통 헬퍼
// ───────────────────────────────────────────────
function jsonRes(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateId() { return Utilities.getUuid(); }
function now()        { return new Date().toISOString(); }

// 시트 가져오기 — 없으면 자동 생성
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = SHEET_HEADERS[name];
    if (headers) {
      const range = sheet.getRange(1, 1, 1, headers.length);
      range.setValues([headers]);
      range.setFontWeight('bold')
           .setBackground(SHEET_COLORS[name] || '#0d1b2a')
           .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// 시트 → 객체 배열 변환
function sheetToObjects(sheetName) {
  const sheet  = getSheet(sheetName);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0].map(String);
  return data.slice(1)
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = (row[i] == null) ? '' : String(row[i]); });
      return obj;
    });
}

// ID로 행 번호 검색 (1-based)
function findRowIndex(sheetName, id) {
  const sheet   = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

// ───────────────────────────────────────────────
//  범용 CRUD
// ───────────────────────────────────────────────
function addRow(sheetName, d) {
  const sheet   = getSheet(sheetName);
  const headers = SHEET_HEADERS[sheetName];
  const id      = d.id || generateId();
  const row     = headers.map(h => {
    if (h === 'id')        return id;
    if (h === 'createdAt') return now();
    if (h === 'updatedAt') return now();
    const v = d[h];
    if (v === undefined || v === null || v === '') return '';
    // 숫자형 컬럼
    if (['targetPrice','stopLoss','rating','price','quantity','avgBuyPrice'].includes(h)) {
      return Number(v) || 0;
    }
    return String(v);
  });
  sheet.appendRow(row);
  return { success: true, id };
}

function updateRow(sheetName, id, d) {
  const rowNum  = findRowIndex(sheetName, id);
  if (rowNum < 0) return { error: '항목을 찾을 수 없습니다 (id: ' + id + ')' };
  const sheet   = getSheet(sheetName);
  const headers = SHEET_HEADERS[sheetName];
  headers.forEach((h, i) => {
    if (h === 'id' || h === 'createdAt') return; // 변경 불가
    if (h === 'updatedAt') {
      sheet.getRange(rowNum, i + 1).setValue(now());
      return;
    }
    if (d[h] === undefined) return; // 전달 안 된 필드는 그대로
    const v = d[h];
    if (['targetPrice','stopLoss','rating','price','quantity','avgBuyPrice'].includes(h)) {
      sheet.getRange(rowNum, i + 1).setValue(Number(v) || 0);
    } else {
      sheet.getRange(rowNum, i + 1).setValue(v == null ? '' : String(v));
    }
  });
  return { success: true };
}

function deleteRow(sheetName, id) {
  const rowNum = findRowIndex(sheetName, id);
  if (rowNum < 0) return { error: '항목을 찾을 수 없습니다 (id: ' + id + ')' };
  getSheet(sheetName).deleteRow(rowNum);
  return { success: true };
}

// ───────────────────────────────────────────────
//  SETUP — 최초 1회 실행 (편집기에서 직접 실행)
// ───────────────────────────────────────────────
function setupSheets() {
  Object.keys(SHEET_HEADERS).forEach(name => getSheet(name));
  SpreadsheetApp.getUi().alert(
    '✅ 시트 설정 완료!\n\n' +
    'STOCKS / TRADES / NOTES / ACCOUNTS 시트가 생성되었습니다.\n\n' +
    '배포 → 배포 관리 → 새 버전 → 배포를 클릭하세요.'
  );
}

// ───────────────────────────────────────────────
//  TEST — 편집기에서 직접 실행
// ───────────────────────────────────────────────
function testPing() {
  Logger.log('ping: ' + JSON.stringify({ ok: true, ts: new Date().toISOString() }));
}

function testGetAccounts() {
  Logger.log('accounts: ' + JSON.stringify(sheetToObjects('ACCOUNTS')));
}

function testAddAccount() {
  const result = addRow('ACCOUNTS', { name: '키움증권 테스트' });
  Logger.log('addAccount: ' + JSON.stringify(result));
}

function testAddTrade() {
  const result = addRow('TRADES', {
    stockName: '삼성전자', acctName: '키움증권',
    type: 'BUY', date: '2026-05-01',
    price: 73000, quantity: 10, avgBuyPrice: 0,
    reason: '신규진입', emotion: '😊', memo: '테스트'
  });
  Logger.log('addTrade: ' + JSON.stringify(result));
}

function testGetTrades() {
  Logger.log('trades: ' + JSON.stringify(sheetToObjects('TRADES')));
}
