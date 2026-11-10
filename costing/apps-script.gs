/* Managish Costing — Google Sheet connector
   1. Go to script.google.com (restaurant Google account) > New project.
      (Or from a Google Sheet: Extensions > Apps Script.)
   2. Delete what is there and paste ALL of this code. Save.
      The Google Sheet "Managish Costing" is created automatically in Drive.
   3. Deploy > New deployment > type "Web app".
      Execute as: Me.   Who has access: Anyone.   Deploy, allow access.
   4. Copy the Web app URL (ends with /exec) and send it to Claude, or paste it in the app's Settings.
   The edit password starts as 0000. Change it in the app (Settings) after connecting. */

var SHEET = 'Data';

function book_() {
  var a = null;
  try { a = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (a) return a;
  var p = props_(), id = p.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    id = p.getProperty('SHEET_ID');
    if (id) return SpreadsheetApp.openById(id);
    var s = SpreadsheetApp.create('Managish Costing');
    p.setProperty('SHEET_ID', s.getId());
    return s;
  } finally { lock.releaseLock(); }
}
/* Optional: run once from the editor to create the Sheet and give permissions */
function setup() { sheet_(); MailApp.getRemainingDailyQuota(); Logger.log(book_().getUrl()); }

function sheet_() {
  var ss = book_();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, sh.getMaxRows(), 2).setNumberFormat('@');
    sh.getRange(1, 1, 1, 4).setValues([['collection', 'id', 'data', 'updated']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function props_() { return PropertiesService.getScriptProperties(); }
function version_() { return Number(props_().getProperty('VERSION') || 0); }

/* Everyone can read */
function doGet(e) {
  var v = version_();
  var since = e && e.parameter ? e.parameter.since : undefined;
  if (since !== undefined && since !== '' && Number(since) === v) return out_({ version: v, unchanged: true });
  var sh = sheet_(), last = sh.getLastRow();
  var data = { items: {}, productions: {}, dishes: {}, settings: {} };
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, 3).getValues();
    for (var i = 0; i < vals.length; i++) {
      var c = vals[i][0], id = String(vals[i][1]), j = vals[i][2];
      if (!c || !j) continue;
      try { (data[c] = data[c] || {})[id] = JSON.parse(j); } catch (err) {}
    }
  }
  return out_({ version: v, data: data, defaultPassword: !props_().getProperty('PASSWORD') });
}

/* Only with the edit password */
function doPost(e) {
  var b = JSON.parse(e.postData.contents);
  var pw = props_().getProperty('PASSWORD') || '0000';
  if (String(b.password) !== pw) return out_({ ok: false, error: 'password' });
  if (b.op === 'check') return out_({ ok: true });
  var sh = sheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (b.op === 'password') {
      props_().setProperty('PASSWORD', String(b.newPassword));
      return out_({ ok: true });
    }
    if (b.op === 'email') {
      var pdf = Utilities.newBlob(String(b.html), MimeType.HTML, 'report.html').getAs(MimeType.PDF).setName(String(b.filename || 'Managish_Costing.pdf'));
      MailApp.sendEmail({
        to: String(b.to),
        subject: String(b.subject || 'Managish costing update'),
        htmlBody: String(b.body || 'Costing update attached.'),
        attachments: [pdf],
        name: 'Managish Costing'
      });
      return out_({ ok: true });
    }
    var last = sh.getLastRow();
    var keys = last > 1 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
    var index = {};
    for (var i = 0; i < keys.length; i++) index[keys[i][0] + '/' + keys[i][1]] = i + 2;
    var now = new Date(), adds = [], dels = [];
    (b.writes || []).forEach(function (w) {
      var k = w.collection + '/' + w.id, row = index[k];
      if (w['delete']) { if (row > 0) dels.push(row); return; }
      var json = JSON.stringify(w.doc);
      if (row > 0) sh.getRange(row, 3, 1, 2).setValues([[json, now]]);
      else if (row < 0) adds[-row - 1][2] = json;
      else { adds.push([w.collection, String(w.id), json, now]); index[k] = -adds.length; }
    });
    if (adds.length) sh.getRange(sh.getLastRow() + 1, 1, adds.length, 4).setValues(adds);
    dels.sort(function (a, c) { return c - a; }).forEach(function (r) { sh.deleteRow(r); });
    var v = version_() + 1;
    props_().setProperty('VERSION', String(v));
    return out_({ ok: true, version: v });
  } finally {
    lock.releaseLock();
  }
}
