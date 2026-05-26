// ═══════════════════════════════════════════════════════════
//  주식노트 — Google Apps Script 백엔드  (Code.gs)
//  배포: 확장 프로그램 > Apps Script > 배포 > 웹 앱
//       실행 계정: 나 / 액세스 권한: 모든 사용자
// ═══════════════════════════════════════════════════════════

// ① 스프레드시트 ID를 여기에 붙여넣으세요
//    URL: https://docs.google.com/spreadsheets/d/[여기]/edit
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';


// ───────────────────────────────────────────────
//  라우터 — GET (읽기)
// ───────────────────────────────────────────────
function doGet(e) {
  try {
    const p = e.parameter;
    let result;
    switch (p.action) {
      case 'getStocks':  result = getStocks();             break;
      case 'getNotes':   result = getNotes(p.stockId);     break;
      case 'getTrades':  result = getTrades(p.stockId);    break;
      case 'ping':       result = { ok: true, ts: new Date().toISOString() }; break;
      default:           result = { error: '알 수 없는 액션: ' + p.action };
    }
    return jsonRes(result);
  } catch (err) {
    return jsonRes({ error: err.toString() });
  }
}


// ───────────────────────────────────────────────
//  라우터 — POST (쓰기)
//  Content-Type: text/plain 으로 호출해야 CORS 에러 없음
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
      // ── 관심종목 ──
      case 'addStock':    result = addStock(body.data);           break;
      case 'updateStock': result = updateStock(body.id, body.data); break;
      case 'deleteStock': result = deleteById('STOCKS', body.id); break;
      // ── 종목 노트 ──
      case 'addNote':     result = addNote(body.data);            break;
      case 'updateNote':  result = updateNote(body.id, body.data); break;
      case 'deleteNote':  result = deleteById('NOTES', body.id);  break;
      // ── 매매일지 ──
      case 'addTrade':    result = addTrade(body.data);           break;
      case 'updateTrade': result = updateTrade(body.id, body.data); break;
      case 'deleteTrade': result = deleteById('TRADES', body.id); break;
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

function generateId() {
  return Utilities.getUuid();
}

function now() {
  return new Date().toISOString();
}

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`시트 "${name}" 를 찾을 수 없습니다. setupSheets() 를 먼저 실행하세요.`);
  return sheet;
}

// 시트 전체를 객체 배열로 변환 (1행 = 헤더)
function sheetToObjects(sheetName) {
  const sheet = getSheet(sheetName);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(String);
  return data.slice(1)
    .filter(row => row[0] !== '' && row[0] !== null) // 빈 행 제외
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

// ID로 행 번호(1-based) 검색
function findRowIndex(sheetName, id) {
  const sheet = getSheet(sheetName);
  const ids   = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // +2: 헤더 행 + 0-index
  }
  return -1;
}

function deleteById(sheetName, id) {
  const row = findRowIndex(sheetName, id);
  if (row < 0) return { error: '항목을 찾을 수 없습니다 (id: ' + id + ')' };
  getSheet(sheetName).deleteRow(row);
  return { success: true };
}


// ───────────────────────────────────────────────
//  STOCKS  관심종목
//  컬럼: id | code | name | sector | tags |
//        targetPrice | stopLoss | rating | holding | memo | createdAt
// ───────────────────────────────────────────────
function getStocks() {
  return sheetToObjects('STOCKS');
}

function addStock(d) {
  const sheet = getSheet('STOCKS');
  const id = generateId();
  sheet.appendRow([
    id,
    d.code        || '',
    d.name        || '',
    d.sector      || '기타',
    d.tags        || '',
    Number(d.targetPrice) || 0,
    Number(d.stopLoss)    || 0,
    Number(d.rating)      || 3,
    d.holding     || '관심',
    d.memo        || '',
    now()
  ]);
  return { success: true, id };
}

function updateStock(id, d) {
  const row = findRowIndex('STOCKS', id);
  if (row < 0) return { error: '항목 없음' };
  const sheet = getSheet('STOCKS');
  // col 2~10 업데이트 (id, createdAt 제외)
  sheet.getRange(row, 2, 1, 9).setValues([[
    d.code        || '',
    d.name        || '',
    d.sector      || '기타',
    d.tags        || '',
    Number(d.targetPrice) || 0,
    Number(d.stopLoss)    || 0,
    Number(d.rating)      || 3,
    d.holding     || '관심',
    d.memo        || ''
  ]]);
  return { success: true };
}


// ───────────────────────────────────────────────
//  NOTES  종목 노트
//  컬럼: id | stockId | stockName | content |
//        risks | links | createdAt | updatedAt
// ───────────────────────────────────────────────
function getNotes(stockId) {
  const notes = sheetToObjects('NOTES');
  return stockId ? notes.filter(n => String(n.stockId) === String(stockId)) : notes;
}

function addNote(d) {
  const sheet = getSheet('NOTES');
  const id = generateId();
  const t  = now();
  sheet.appendRow([
    id,
    d.stockId   || '',
    d.stockName || '',
    d.content   || '',
    d.risks     || '',
    d.links     || '',
    t, t
  ]);
  return { success: true, id };
}

function updateNote(id, d) {
  const row = findRowIndex('NOTES', id);
  if (row < 0) return { error: '항목 없음' };
  const sheet = getSheet('NOTES');
  sheet.getRange(row, 4, 1, 3).setValues([[
    d.content || '',
    d.risks   || '',
    d.links   || ''
  ]]);
  sheet.getRange(row, 8).setValue(now()); // updatedAt
  return { success: true };
}


// ───────────────────────────────────────────────
//  TRADES  매매일지
//  컬럼: id | stockId | stockName | type | date |
//        price | quantity | reason | emotion | memo | createdAt
// ───────────────────────────────────────────────
function getTrades(stockId) {
  const trades = sheetToObjects('TRADES');
  return stockId ? trades.filter(t => String(t.stockId) === String(stockId)) : trades;
}

function addTrade(d) {
  const sheet = getSheet('TRADES');
  const id = generateId();
  sheet.appendRow([
    id,
    d.stockId   || '',
    d.stockName || '',
    d.type      || 'BUY',
    d.date      || '',
    Number(d.price)    || 0,
    Number(d.quantity) || 0,
    d.reason    || '',
    d.emotion   || '',
    d.memo      || '',
    now()
  ]);
  return { success: true, id };
}

function updateTrade(id, d) {
  const row = findRowIndex('TRADES', id);
  if (row < 0) return { error: '항목 없음' };
  const sheet = getSheet('TRADES');
  sheet.getRange(row, 4, 1, 7).setValues([[
    d.type      || 'BUY',
    d.date      || '',
    Number(d.price)    || 0,
    Number(d.quantity) || 0,
    d.reason    || '',
    d.emotion   || '',
    d.memo      || ''
  ]]);
  return { success: true };
}


// ───────────────────────────────────────────────
//  SETUP  최초 1회 실행 → 시트 구조 생성
//  Apps Script 편집기에서 직접 실행하세요
// ───────────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function makeSheet(name, headers) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      sheet.clearContents();
    }
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#0d1b2a');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    // 열 너비 자동 조정
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    return sheet;
  }

  makeSheet('STOCKS', [
    'id','code','name','sector','tags',
    'targetPrice','stopLoss','rating','holding','memo','createdAt'
  ]);
  makeSheet('NOTES', [
    'id','stockId','stockName','content',
    'risks','links','createdAt','updatedAt'
  ]);
  makeSheet('TRADES', [
    'id','stockId','stockName','type','date',
    'price','quantity','reason','emotion','memo','createdAt'
  ]);

  SpreadsheetApp.getUi().alert(
    '✅ 시트 설정 완료!\n\n' +
    'STOCKS / NOTES / TRADES 시트가 생성되었습니다.\n' +
    '이제 웹 앱으로 배포하세요.'
  );
}


// ───────────────────────────────────────────────
//  TEST  편집기에서 직접 실행해 정상 동작 확인
// ───────────────────────────────────────────────
function testAddStock() {
  const result = addStock({
    code: '005930', name: '삼성전자', sector: '반도체',
    tags: 'HBM,AI서버', targetPrice: 85000, stopLoss: 68000,
    rating: 4, holding: '관심', memo: 'HBM3E 점유율 기대'
  });
  Logger.log(JSON.stringify(result));
}

function testGetStocks() {
  Logger.log(JSON.stringify(getStocks()));
}
