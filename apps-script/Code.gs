// BorrowBloc backend: paste into Extensions → Apps Script on your Google Sheet.
// Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
// Copy the /exec URL into SHEET_ENDPOINT in index.html.
//
// Expects a tab named "Inventory" with headers in row 1:
// ID | Tool Name | Category | Status | Borrower | Due Date

const SHEET_NAME = 'Inventory';

function sheet_() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === '—' ? '' : String(v || '');
}

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const tools = rows
    .filter(r => r[0] !== '')
    .map(r => ({
      id: Number(r[0]),
      name: r[1],
      category: r[2],
      status: r[3],
      borrower: r[4] === '—' ? '' : r[4],
      due: fmtDate_(r[5]),
    }));
  return json_({ tools });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sh = sheet_();

    if (body.action === 'add') {
      sh.appendRow([body.id, clean_(body.name), clean_(body.category), 'Available', '—', '—']);
      return json_({ ok: true });
    }

    const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => Number(r[0]));
    const idx = ids.indexOf(Number(body.id));
    if (idx === -1) return json_({ ok: false, error: 'Tool not found' });
    const row = idx + 2;

    if (body.action === 'checkout') {
      if (sh.getRange(row, 4).getValue() === 'Checked Out') return json_({ ok: false, error: 'Already checked out' });
      // Store the due date as text so Sheets doesn't reformat it.
      sh.getRange(row, 4, 1, 3).setValues([['Checked Out', clean_(body.borrower), "'" + clean_(body.due)]]);
    } else if (body.action === 'return') {
      sh.getRange(row, 4, 1, 3).setValues([['Available', '—', '—']]);
    } else {
      return json_({ ok: false, error: 'Unknown action' });
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Strip leading formula characters so a borrower name can't become a sheet formula.
function clean_(s) {
  return String(s || '').slice(0, 60).replace(/^[=+\-@]+/, '');
}
