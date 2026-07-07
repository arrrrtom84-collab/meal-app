/**
 * 도니도니용돈기입장 - Google Apps Script 백엔드 (code.gs)
 * '시트1'을 DB로 삼아 입출금 데이터를 조회, 저장, 수정, 삭제(CRUD)합니다.
 */

const SHEET_NAME = '시트1';

// 1. 웹앱 초기 실행 (HTML 파일 렌더링)
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('도니도니용돈기입장')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. 스프레드시트 및 시트 가져오기 (없으면 '시트1' 자동 생성 및 헤더 구성)
function getTargetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // 기본 헤더 행 추가
    sheet.appendRow(['ID', '날짜', '구분', '내용', '금액']);
  }
  return sheet;
}

// 3. 목표 금액 가져오기/저장하기 전용 시트 (또는 스크립트 속성 이용)
// 여기서는 간편하고 확실하게 데이터 관리를 위해 ScriptProperties를 활용해 목표 금액을 저장합니다.
function getMonthlyGoal() {
  const userProperties = PropertiesService.getScriptProperties();
  const goal = userProperties.getProperty('monthly_goal');
  return goal ? parseInt(goal, 10) : 50000; // 기본 목표액 50,000원
}

function setMonthlyGoal(goalAmount) {
  const userProperties = PropertiesService.getScriptProperties();
  userProperties.setProperty('monthly_goal', goalAmount.toString());
  return true;
}

// 4. 데이터 조회 (Read)
function getTransactions() {
  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return []; // 헤더만 있는 경우 빈 배열 반환
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const transactions = data.map(function(row) {
    return {
      id: row[0],
      date: Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      type: row[2], // '입금' 또는 '출금'
      content: row[3],
      amount: Number(row[4])
    };
  });
  
  // 날짜 기준 내림차순 정렬 (최신 정보가 위로 오도록)
  transactions.sort((a, b) => new Date(b.date) - new Date(a.date));
  
  return transactions;
}

// 5. 데이터 추가 (Create)
function addTransaction(item) {
  const sheet = getTargetSheet();
  const id = 'ID_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
  const dateObj = new Date(item.date);
  
  sheet.appendRow([
    id,
    dateObj,
    item.type,
    item.content,
    Number(item.amount)
  ]);
  
  return getTransactions(); // 업데이트된 목록 반환
}

// 6. 데이터 수정 (Update)
function updateTransaction(id, updatedItem) {
  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return getTransactions();
  
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flatMap(row => row[0]);
  const rowIndex = ids.indexOf(id);
  
  if (rowIndex !== -1) {
    const sheetRow = rowIndex + 2; // 1-based index & header row offset
    sheet.getRange(sheetRow, 2).setValue(new Date(updatedItem.date));
    sheet.getRange(sheetRow, 3).setValue(updatedItem.type);
    sheet.getRange(sheetRow, 4).setValue(updatedItem.content);
    sheet.getRange(sheetRow, 5).setValue(Number(updatedItem.amount));
  }
  
  return getTransactions();
}

// 7. 데이터 삭제 (Delete)
function deleteTransaction(id) {
  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return getTransactions();
  
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flatMap(row => row[0]);
  const rowIndex = ids.indexOf(id);
  
  if (rowIndex !== -1) {
    const sheetRow = rowIndex + 2;
    sheet.deleteRow(sheetRow);
  }
  
  return getTransactions();
}