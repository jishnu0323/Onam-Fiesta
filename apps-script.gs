/**********************************************************************
 *  CAFS Onam 2026 — Participant Registration backend (Google Apps Script)
 *
 *  WHAT THIS DOES
 *   - Saves participant registrations into a "Participants" tab of your
 *     Google Sheet (the SAME sheet the app reads from).
 *   - Checks the admin username/password SERVER-SIDE so the password is
 *     never exposed in the public web page.
 *
 *  SETUP (one time)
 *   1. Open your Google Sheet.
 *   2. Extensions -> Apps Script.  Delete any sample code.
 *   3. Paste ALL of this file in.  Set SHEET_ID below (from the sheet URL:
 *        docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit ).
 *   4. (Optional) change the admin credentials in ADMINS.
 *   5. Deploy -> New deployment -> type "Web app".
 *        - Description: Onam registration
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Click Deploy, authorise when asked, and COPY the Web app URL
 *      (looks like https://script.google.com/macros/s/XXXX/exec ).
 *   6. Paste that URL into register.html -> CONFIG.SCRIPT_URL.
 *
 *  If you change this code later: Deploy -> Manage deployments ->
 *  edit the existing deployment -> Version: New version -> Deploy
 *  (keeps the SAME URL).
 **********************************************************************/

const SHEET_ID = 'PASTE_YOUR_SHEET_ID_HERE';
const TAB      = 'Participants';

// username : password   (server-side only, never sent to the public page)
const ADMINS = {
  'admin': 'cloudOnam26!'
};

function HEADERS() {
  return ['ID','Timestamp','Event','Name','Department','Phone','Team','Gender','Notes','AddedBy'];
}

function doGet(e)  { return handle({ action: 'list' }); }
function doPost(e) {
  var data = {};
  try { data = JSON.parse(e.postData.contents); } catch (err) {}
  return handle(data);
}

function handle(data) {
  var out;
  try {
    var action = (data && data.action) || 'list';
    if (action === 'list') {
      out = { ok: true, rows: listRows() };
    } else if (action === 'auth') {
      out = { ok: valid(data.user, data.pass) };
      if (!out.ok) out.error = 'Invalid username or password';
    } else if (action === 'add') {
      if (!valid(data.user, data.pass)) { out = { ok: false, error: 'Invalid login' }; }
      else { addRow(data.p || {}, data.user); out = { ok: true, rows: listRows() }; }
    } else if (action === 'delete') {
      if (!valid(data.user, data.pass)) { out = { ok: false, error: 'Invalid login' }; }
      else { deleteRow(data.rowId); out = { ok: true, rows: listRows() }; }
    } else {
      out = { ok: false, error: 'Unknown action' };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
                       .setMimeType(ContentService.MimeType.JSON);
}

function valid(u, p) { return !!u && ADMINS.hasOwnProperty(u) && ADMINS[u] === p; }

function sheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB);
  if (!sh) sh = ss.insertSheet(TAB);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS());
  return sh;
}

function listRows() {
  var sh = sheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, HEADERS().length).getValues();
  return vals.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { id: r[0], ts: r[1], event: r[2], name: r[3], dept: r[4],
             phone: r[5], team: r[6], gender: r[7], notes: r[8], addedBy: r[9] };
  });
}

function addRow(p, user) {
  var sh = sheet();
  var id = 'P' + (new Date().getTime());
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  sh.appendRow([id, ts, p.event || '', p.name || '', p.dept || '', p.phone || '',
                p.team || '', p.gender || '', p.notes || '', user || '']);
}

function deleteRow(rowId) {
  var sh = sheet();
  var last = sh.getLastRow();
  if (last < 2) return;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === rowId) { sh.deleteRow(i + 2); return; }
  }
}
