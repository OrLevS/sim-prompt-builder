// שמירה משותפת של תרגילים בגיליון Google.
// הוראות התקנה: README.md בתיקייה הראשית.

const SHEET_NAME = 'exercises';
const KEY = '2221';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id', 'createdAt', 'name', 'author', 'students', 'mode', 'data']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const rows = sheet_().getDataRange().getValues();
    rows.shift();
    const items = rows.filter(r => r[0]).map(r => ({
      id: String(r[0]), createdAt: r[1], name: r[2], author: r[3], students: r[4], mode: r[5], data: r[6]
    })).reverse();
    return json_({ ok: true, items: items });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const body = JSON.parse(e.postData.contents);
    if (body.key !== KEY) return json_({ ok: false, error: 'אין הרשאה' });
    if (body.action !== 'save') return json_({ ok: false, error: 'פעולה לא מוכרת' });
    const x = body.entry || {};
    const data = JSON.stringify(x.data || {});
    if (data.length > 45000) return json_({ ok: false, error: 'התרגיל ארוך מדי לשמירה' });
    sheet_().appendRow([
      String(x.id || ''), x.createdAt || new Date().toISOString(),
      String(x.name || '').slice(0, 200), String(x.author || '').slice(0, 100),
      String(x.students || '').slice(0, 200), x.mode === 'auto' ? 'auto' : 'manual', data
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
