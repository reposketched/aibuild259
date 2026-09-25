// BorrowBloc backend: paste into Extensions → Apps Script on your Google Sheet.
// Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
// Copy the /exec URL into SHEET_ENDPOINT in index.html.
//
// Expects a tab named "Inventory" with headers in row 1:
// ID | Tool Name | Category | Status | Borrower | Due Date | Photo | Owner | Waitlist | Reservations | Condition | Condition Note
//   Photo:        small base64 JPEG (kept under the 50k cell limit), an https URL, or a path like assets/tools/ladder.jpg
//   Waitlist:     comma-separated names
//   Reservations: JSON, e.g. [{"name":"Ria Nimbkar","start":"2026-10-03","end":"2026-10-04"}]
//   Condition:    Good | Needs attention | Broken
// "Loans" and "Requests" tabs are created automatically.

const SHEET_NAME = 'Inventory';
const COL = { status: 4, borrower: 5, due: 6, photo: 7, owner: 8, waitlist: 9, reservations: 10, condition: 11, conditionNote: 12 };
const TABS = {
  Loans: ['Date', 'Tool ID', 'Tool', 'Category', 'Borrower', 'Owner', 'Due', 'Returned', 'Condition', 'Note'],
  Requests: ['ID', 'Tool', 'Note', 'By', 'Date', 'Status', 'Fulfilled By', 'Fulfilled On'],
};

function sheet_() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
}

function tab_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(TABS[name]); }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === '—' ? '' : String(v || '');
}

function waitlist_(v) {
  return String(v || '').split(',').map(s => s.trim()).filter(s => s && s !== '—');
}

function reservations_(v) {
  try {
    const list = JSON.parse(v || '[]');
    return Array.isArray(list) ? list.filter(r => r && r.name && r.start && r.end && r.end >= today_()) : [];
  } catch (e) {
    return [];
  }
}

// Dates are written with a leading apostrophe so Sheets keeps them as plain text.
const text_ = s => "'" + s;

function doGet() {
  const tools = sheet_().getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '')
    .map(r => ({
      id: Number(r[0]), name: r[1], category: r[2], status: r[3],
      borrower: r[4] === '—' ? '' : r[4], due: fmtDate_(r[5]), photo: r[6] || '', owner: r[7] || '',
      waitlist: waitlist_(r[8]), reservations: reservations_(r[9]),
      condition: r[10] || 'Good', conditionNote: r[11] || '',
    }));
  const loans = tab_('Loans').getDataRange().getValues().slice(1)
    .filter(r => r[1] !== '')
    .map(r => ({
      date: fmtDate_(r[0]), toolId: Number(r[1]), tool: r[2], category: r[3], borrower: r[4], owner: r[5],
      due: fmtDate_(r[6]), returned: fmtDate_(r[7]), condition: r[8] || '', note: r[9] || '',
    }));
  const requests = tab_('Requests').getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '')
    .map(r => ({
      id: Number(r[0]), tool: r[1], note: r[2], by: r[3], date: fmtDate_(r[4]),
      status: r[5] || 'open', fulfilledBy: r[6] || '', fulfilledOn: fmtDate_(r[7]),
    }));
  return json_({ tools, loans, requests });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sh = sheet_();
    const same = (a, b) => String(a).trim().toLowerCase() === String(b).trim().toLowerCase();

    // ── Requests (wish list) ──
    if (body.action === 'request') {
      const rq = tab_('Requests');
      const id = Math.max(0, ...rq.getDataRange().getValues().slice(1).map(r => Number(r[0]) || 0)) + 1;
      rq.appendRow([id, clean_(body.tool), clean_(body.note), clean_(body.by), text_(today_()), 'open', '', '']);
      return json_({ ok: true, id });
    }
    if (body.action === 'closeRequest') {
      const row = findRow_(tab_('Requests'), 1, body.id);
      if (!row) return json_({ ok: false, error: 'Request not found' });
      tab_('Requests').getRange(row, 6).setValue('closed');
      return json_({ ok: true });
    }

    if (body.action === 'add') {
      const id = Math.max(0, ...sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => Number(r[0]) || 0)) + 1;
      sh.appendRow([id, clean_(body.name), clean_(body.category), 'Available', '—', '—', photo_(body.photo), clean_(body.owner), '', '[]', 'Good', '']);
      if (body.requestId) {
        const rq = tab_('Requests'), row = findRow_(rq, 1, body.requestId);
        if (row) rq.getRange(row, 6, 1, 3).setValues([['fulfilled', clean_(body.owner), text_(today_())]]);
      }
      return json_({ ok: true, id });
    }

    // ── Everything else acts on one tool row ──
    const row = findRow_(sh, 1, body.id);
    if (!row) return json_({ ok: false, error: 'Tool not found' });
    const [, name, category, status, borrower, due, , owner, wl, resv, condition] = sh.getRange(row, 1, 1, 12).getValues()[0];
    const queue = waitlist_(wl);
    const bookings = reservations_(resv);
    const today = today_();

    if (body.action === 'checkout') {
      const who = clean_(body.borrower);
      const active = bookings.find(r => r.start <= today && today <= r.end);
      const held = active ? active.name : queue[0];
      if (status === 'Checked Out') return json_({ ok: false, error: 'Already checked out' });
      if (condition === 'Broken') return json_({ ok: false, error: 'Out for repair' });
      if (held && !same(held, who)) return json_({ ok: false, error: 'Held for ' + held });
      const next = bookings.filter(r => r.start > today && !same(r.name, who)).sort((a, b) => a.start < b.start ? -1 : 1)[0];
      if (next && String(body.due) >= next.start) return json_({ ok: false, error: 'Due date runs into a booking on ' + next.start });
      sh.getRange(row, COL.status, 1, 3).setValues([['Checked Out', who, text_(clean_(body.due))]]);
      sh.getRange(row, COL.waitlist, 1, 2).setValues([[
        queue.filter(n => !same(n, who)).join(', '),
        JSON.stringify(bookings.filter(r => !(same(r.name, who) && r.start <= today))),
      ]]);
      tab_('Loans').appendRow([text_(today), Number(body.id), name, category, who, owner, text_(clean_(body.due)), '', '', '']);
    } else if (body.action === 'return') {
      const cond = ['Good', 'Needs attention', 'Broken'].includes(body.condition) ? body.condition : 'Good';
      const note = clean_(body.note);
      sh.getRange(row, COL.status, 1, 3).setValues([['Available', '—', '—']]);
      sh.getRange(row, COL.condition, 1, 2).setValues([[cond, cond === 'Good' ? '' : note]]);
      // Close the most recent open loan for this tool.
      const loansSh = tab_('Loans');
      const data = loansSh.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (Number(data[i][1]) === Number(body.id) && !data[i][7]) {
          loansSh.getRange(i + 1, 8, 1, 3).setValues([[text_(today), cond, note]]);
          break;
        }
      }
    } else if (body.action === 'fixed') {
      sh.getRange(row, COL.condition, 1, 2).setValues([['Good', '']]);
    } else if (body.action === 'photo') {
      sh.getRange(row, COL.photo).setValue(photo_(body.photo));
    } else if (body.action === 'join') {
      const who = clean_(body.name).replace(/,/g, '');
      if (!queue.some(n => same(n, who))) queue.push(who);
      sh.getRange(row, COL.waitlist).setValue(queue.join(', '));
    } else if (body.action === 'leave') {
      sh.getRange(row, COL.waitlist).setValue(queue.filter(n => !same(n, body.name)).join(', '));
    } else if (body.action === 'reserve') {
      const start = String(body.start), end = String(body.end);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end) || end < start || start < today)
        return json_({ ok: false, error: 'Invalid dates' });
      if (status === 'Checked Out' && fmtDate_(due) && start <= fmtDate_(due)) return json_({ ok: false, error: 'Tool is out until ' + fmtDate_(due) });
      const clash = bookings.find(r => start <= r.end && r.start <= end);
      if (clash) return json_({ ok: false, error: 'Clashes with a booking by ' + clash.name });
      bookings.push({ name: clean_(body.name), start, end });
      bookings.sort((a, b) => a.start < b.start ? -1 : 1);
      sh.getRange(row, COL.reservations).setValue(JSON.stringify(bookings));
    } else if (body.action === 'unreserve') {
      sh.getRange(row, COL.reservations).setValue(JSON.stringify(bookings.filter(r => !(same(r.name, body.name) && r.start === body.start))));
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

function findRow_(sh, col, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const idx = sh.getRange(2, col, n, 1).getValues().map(r => Number(r[0])).indexOf(Number(id));
  return idx === -1 ? 0 : idx + 2;
}

function photo_(p) {
  p = String(p || '');
  return /^(data:image\/(png|jpe?g|webp|gif);base64,|https:\/\/|assets\/[\w\/.-]+$)/.test(p) && p.length <= 50000 ? p : '';
}

// Strip leading formula characters so user text can't become a sheet formula.
function clean_(s) {
  return String(s || '').slice(0, 140).replace(/^[=+\-@]+/, '');
}
