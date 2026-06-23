function doGet() {
  updateVisitCount();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('오범석 링크 모음')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Links');
  if (!sheet) sheet = ss.insertSheet('Links');
  const data = sheet.getDataRange().getValues();
  data.shift();

  const links = data.map((row, i) => {
    const [cat, title, url, emoji, pinned, position, ...rest] = row;
    const subs = [];
    for (let j = 0; j < rest.length; j += 2) {
      if (rest[j] || rest[j + 1]) subs.push({ title: rest[j], url: rest[j + 1] });
    }
    return {
      rowIndex: i + 2,
      category: cat || '기타',
      title: title || '',
      url: url || '',
      emoji: emoji || '🔗',
      pinned: pinned === true || String(pinned).toUpperCase() === 'TRUE',
      position: position === '오른쪽' ? 'right' : 'left',
      subs
    };
  }).filter(link => link.title !== '');

  return { links };
}

function addLink(payload) {
  if (payload.pw !== "9600") return "비밀번호 오류";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Links');
  const posText = payload.position === 'right' ? '오른쪽' : '왼쪽';
  sheet.appendRow([payload.cat, payload.title, payload.url, payload.emoji, payload.pinned, posText, ...payload.subs]);
  return "추가 완료!";
}

function editLink(payload) {
  if (payload.pw !== "9600") return "비밀번호 오류";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Links');
  const posText = payload.position === 'right' ? '오른쪽' : '왼쪽';
  const rowData = [payload.cat, payload.title, payload.url, payload.emoji, payload.pinned, posText, ...payload.subs];
  const maxCols = Math.max(rowData.length, sheet.getLastColumn() || 1);
  while (rowData.length < maxCols) rowData.push("");
  sheet.getRange(payload.rowIndex, 1, 1, rowData.length).setValues([rowData]);
  return "수정 완료!";
}

function deleteLink(pw, rowIndex) {
  if (pw !== "9600") return "비밀번호 오류";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Links');
  sheet.deleteRow(rowIndex);
  return "삭제 완료!";
}

function syncLinksOrder(pw, newOrderData) {
  if (pw !== "9600") return "비밀번호 오류";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Links');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (newOrderData && newOrderData.length > 0) {
    const maxCols = Math.max(lastCol || 1, Math.max(...newOrderData.map(r => r.length)));
    const paddedData = newOrderData.map(row => {
      while (row.length < maxCols) row.push("");
      return row;
    });
    sheet.getRange(2, 1, paddedData.length, maxCols).setValues(paddedData);
  }
  return "순서 저장 완료";
}

function updateVisitCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsSheet = ss.getSheetByName('Stats') || ss.insertSheet('Stats');
  const range = statsSheet.getRange('A1');
  const count = range.getValue() || 0;
  range.setValue(Number(count) + 1);
}
