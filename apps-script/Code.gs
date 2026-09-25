// BorrowBloc backend: paste into Extensions → Apps Script on your Google Sheet.
// Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
// Copy the /exec URL into SHEET_ENDPOINT in index.html.
//
// Expects a tab named "Inventory" with headers in row 1:
// ID | Tool Name | Category | Status | Borrower | Due Date | Photo | Owner | Waitlist
// Photo holds a small base64 JPEG (the page keeps it under the 50k cell limit), an https URL,
// or a bundled path like assets/tools/ladder.jpg. Waitlist is a comma-separated list of names.
// A "Loans" tab (Date | Tool ID | Tool | Category | Borrower | Owner) is created automatically.

const SHEET_NAME = 'Inventory';
const LOANS_NAME = 'Loans';
const COL = { status: 4, borrower: 5, due: 6, photo: 7, owner: 8, waitlist: 9 };

function sheet_() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
}

function loansSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(LOANS_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOANS_NAME);
    sh.appendRow(['Date', 'Tool ID', 'Tool', 'Category', 'Borrower', 'Owner']);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === '—' ? '' : String(v || '');
}

function waitlist_(v) {
  return String(v || '').split(',').map(s => s.trim()).filter(s => s && s !== '—');
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
      photo: r[6] || '',
      owner: r[7] || '',
      waitlist: waitlist_(r[8]),
    }));
  const loans = loansSheet_().getDataRange().getValues().slice(1)
    .filter(r => r[1] !== '')
    .map(r => ({ date: fmtDate_(r[0]), toolId: Number(r[1]), tool: r[2], category: r[3], borrower: r[4], owner: r[5] }));
  return json_({ tools, loans });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sh = sheet_();

    if (body.action === 'add') {
      const id = Math.max(0, ...sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => Number(r[0]) || 0)) + 1;
      sh.appendRow([id, clean_(body.name), clean_(body.category), 'Available', '—', '—', photo_(body.photo), clean_(body.owner), '']);
      return json_({ ok: true });
    }

    const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => Number(r[0]));
    const idx = ids.indexOf(Number(body.id));
    if (idx === -1) return json_({ ok: false, error: 'Tool not found' });
    const row = idx + 2;
    const [name, category, status, , , , owner, wl] = sh.getRange(row, 2, 1, 8).getValues()[0];
    const queue = waitlist_(wl);
    const same = (a, b) => String(a).trim().toLowerCase() === String(b).trim().toLowerCase();

    if (body.action === 'checkout') {
      const borrower = clean_(body.borrower);
      if (status === 'Checked Out') return json_({ ok: false, error: 'Already checked out' });
      if (queue.length && !same(queue[0], borrower)) return json_({ ok: false, error: 'Held for ' + queue[0] });
      // Store the due date as text so Sheets doesn't reformat it.
      sh.getRange(row, COL.status, 1, 3).setValues([['Checked Out', borrower, "'" + clean_(body.due)]]);
      sh.getRange(row, COL.waitlist).setValue(queue.filter(n => !same(n, borrower)).join(', '));
      loansSheet_().appendRow(["'" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), Number(body.id), name, category, borrower, owner]);
    } else if (body.action === 'return') {
      sh.getRange(row, COL.status, 1, 3).setValues([['Available', '—', '—']]);
    } else if (body.action === 'photo') {
      sh.getRange(row, COL.photo).setValue(photo_(body.photo));
    } else if (body.action === 'join') {
      const who = clean_(body.name).replace(/,/g, '');
      if (!queue.some(n => same(n, who))) queue.push(who);
      sh.getRange(row, COL.waitlist).setValue(queue.join(', '));
    } else if (body.action === 'leave') {
      sh.getRange(row, COL.waitlist).setValue(queue.filter(n => !same(n, body.name)).join(', '));
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

function photo_(p) {
  p = String(p || '');
  return /^(data:image\/(png|jpe?g|webp|gif);base64,|https:\/\/|assets\/[\w\/.-]+$)/.test(p) && p.length <= 50000 ? p : '';
}

// Strip leading formula characters so a name can't become a sheet formula.
function clean_(s) {
  return String(s || '').slice(0, 60).replace(/^[=+\-@]+/, '');
}
