// ═══════════════════════════════════════════════════════════
//  StockNote — Google Apps Script 백엔드  (Code.gs)
//  배포: 확장 프로그램 > Apps Script > 배포 > 웹 앱
//       실행 계정: 나 / 액세스 권한: 모든 사용자
//
//  ※ 최초 1회: 편집기에서 setupSheets() 함수를 직접 실행하세요
// ═══════════════════════════════════════════════════════════

// ① 스프레드시트 ID를 여기에 붙여넣으세요
//    URL: https://docs.google.com/spreadsheets/d/[여기]/edit
const SPREADSHEET_ID = '1x68B7wIl4yCu-56V-inmxQGUkOFveCn3xXiIie2TX98';


// ───────────────────────────────────────────────
//  라우터 — GET (읽기)
// ───────────────────────────────────────────────
function doGet(e) {
  try {
    const p = e.parameter;
    let result;
    switch (p.action) {
      case 'getStocks':   result = getStocks();            break;
      case 'getTrades':   result = getTrades(p.stockId);   break;
      case 'getNotes':    result = getNotes(p.stockId);    break;
      case 'getAccounts': result = getAccounts();          break;  // ✅ 수정: 계좌 전용 함수 호출
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
      // ── 관심종목 ──────────────────────────────
      case 'addStock':      result = addStock(body.data);              break;
      case 'updateStock':   result = updateStock(body.id, body.data);  break;
      case 'deleteStock':   result = deleteById('STOCKS', body.id);    break;
      // ── 종목노트 ──────────────────────────────
      case 'addNote':       result = addNote(body.data);               break;
      case 'updateNote':    result = updateNote(body.id, body.data);   break;
      case 'deleteNote':    result = deleteById('NOTES', body.id);     break;
      // ── 매매일지 ──────────────────────────────
      case 'addTrade':      result = addTrade(body.data);              break;
      case 'updateTrade':   result = updateTrade(body.id, body.data);  break;
      case 'deleteTrade':   result = deleteById('TRADES', body.id);    break;
      // ── 계좌 관리 ─────────────────────────────  ✅ 수정: 계좌 전용 함수 호출
      case 'addAccount':    result = addAccount(body.data);            break;
      case 'updateAccount': result = updateAccount(body.id, body.data); break;
      case 'deleteAccount': result = deleteById('ACCOUNTS', body.id);  break;
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
  if (!sheet) throw new Error(
    `시트 "${name}"를 찾을 수 없습니다. 편집기에서 setupSheets()를 먼저 실행하세요.`
  );
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
      headers.forEach((h, i) => {
        // 숫자형 컬럼은 문자열로 변환하여 일관성 유지
        obj[h] = (row[i] === null || row[i] === undefined) ? '' : String(row[i]);
      });
      return obj;
    });
}

// ID로 행 번호(1-based) 검색
function findRowIndex(sheetName, id) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
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
    d.code              || '',
    d.name              || '',
    d.sector            || '기타',
    d.tags              || '',
    Number(d.targetPrice) || 0,
    Number(d.stopLoss)    || 0,
    Number(d.rating)      || 3,
    d.holding           || '관심',
    d.memo              || '',
    now()
  ]);
  return { success: true, id };
}

function updateStock(id, d) {
  const row = findRowIndex('STOCKS', id);
  if (row < 0) return { error: '항목 없음' };
  const sheet = getSheet('STOCKS');
  // col 2~10 업데이트 (id=1, createdAt=11 제외)
  sheet.getRange(row, 2, 1, 9).setValues([[
    d.code              || '',
    d.name              || '',
    d.sector            || '기타',
    d.tags              || '',
    Number(d.targetPrice) || 0,
    Number(d.stopLoss)    || 0,
    Number(d.rating)      || 3,
    d.holding           || '관심',
    d.memo              || ''
  ]]);
  return { success: true };
}


// ───────────────────────────────────────────────
//  NOTES  종목노트
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
//  TRADES  매매일지                              ✅ 수정: acctName, avgBuyPrice 추가
//  컬럼: id | stockName | acctName | type | date |
//        price | quantity | avgBuyPrice | reason | emotion | memo | createdAt
// ───────────────────────────────────────────────
function getTrades(stockId) {
  const trades = sheetToObjects('TRADES');
  return stockId ? trades.filter(t => String(t.stockName) === String(stockId)) : trades;
}

function addTrade(d) {
  const sheet = getSheet('TRADES');
  const id = generateId();
  sheet.appendRow([
    id,
    d.stockName         || '',
    d.acctName          || '',          // ✅ 계좌명
    d.type              || 'BUY',
    d.date              || '',
    Number(d.price)     || 0,
    Number(d.quantity)  || 0,
    Number(d.avgBuyPrice) || 0,         // ✅ 평균매수가
    d.reason            || '',
    d.emotion           || '',
    d.memo              || '',
    now()
  ]);
  return { success: true, id };
}

function updateTrade(id, d) {
  const row = findRowIndex('TRADES', id);
  if (row < 0) return { error: '항목 없음' };
  const sheet = getSheet('TRADES');
  // col 2~11 업데이트 (id=1, createdAt=12 제외)
  sheet.getRange(row, 2, 1, 10).setValues([[
    d.stockName           || '',
    d.acctName            || '',        // ✅ 계좌명
    d.type                || 'BUY',
    d.date                || '',
    Number(d.price)       || 0,
    Number(d.quantity)    || 0,
    Number(d.avgBuyPrice) || 0,         // ✅ 평균매수가
    d.reason              || '',
    d.emotion             || '',
    d.memo                || ''
  ]]);
  return { success: true };
}


// ───────────────────────────────────────────────
//  ACCOUNTS  계좌 관리                           ✅ 신규 추가
//  컬럼: id | name | createdAt
// ───────────────────────────────────────────────
function getAccounts() {
  return sheetToObjects('ACCOUNTS');
}

function addAccount(d) {
  const sheet = getSheet('ACCOUNTS');
  const id = d.id || generateId();
  sheet.appendRow([
    id,
    d.name || '',
    now()
  ]);
  return { success: true, id };
}

function updateAccount(id, d) {
  const row = findRowIndex('ACCOUNTS', id);
  if (row < 0) return { error: '항목 없음' };
  const sheet = getSheet('ACCOUNTS');
  sheet.getRange(row, 2).setValue(d.name || ''); // name 컬럼만 업데이트
  return { success: true };
}


// ───────────────────────────────────────────────
//  SETUP  최초 1회 실행 → 시트 구조 생성        ✅ ACCOUNTS, TRADES 컬럼 수정
//  Apps Script 편집기에서 직접 실행하세요
// ───────────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function makeSheet(name, headers, color) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      sheet.clearContents();
    }
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground(color || '#0d1b2a');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);
    return sheet;
  }

  makeSheet('STOCKS', [
    'id','code','name','sector','tags',
    'targetPrice','stopLoss','rating','holding','memo','createdAt'
  ], '#1e40af');

  makeSheet('TRADES', [
    'id','stockName','acctName','type','date',       // ✅ acctName 추가
    'price','quantity','avgBuyPrice',                 // ✅ avgBuyPrice 추가
    'reason','emotion','memo','createdAt'
  ], '#065f46');

  makeSheet('NOTES', [
    'id','stockId','stockName','content',
    'risks','links','createdAt','updatedAt'
  ], '#7e22ce');

  makeSheet('ACCOUNTS', [                             // ✅ 신규 시트
    'id','name','createdAt'
  ], '#92400e');

  SpreadsheetApp.getUi().alert(
    '✅ 시트 설정 완료!\n\n' +
    'STOCKS / TRADES / NOTES / ACCOUNTS 시트가 생성되었습니다.\n' +
    '이제 배포 > 배포 관리 > 새 버전으로 배포하세요.'
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

function testAddAccount() {
  const result = addAccount({ name: '키움증권' });
  Logger.log(JSON.stringify(result));
}

function testGetAccounts() {
  Logger.log(JSON.stringify(getAccounts()));
}

function testAddTrade() {
  const result = addTrade({
    stockName: '삼성전자', acctName: '키움증권',
    type: 'BUY', date: '2026-05-01',
    price: 73000, quantity: 10, avgBuyPrice: 0,
    reason: '신규진입', emotion: '😊', memo: '테스트'
  });
  Logger.log(JSON.stringify(result));
}
