// Extensiones → Apps Script, pegar este archivo, Implementar → Nueva implementación
// → Aplicación web, ejecutar como yo, acceso: cualquier persona. Autorizar y
// copiar la URL que termina en /exec dentro de SCRIPT_URL en index.html.
var SHEET_ID = "1l-hyNvIWsPIPollb8dWCnVlvbZff3svX2XV6DKxMRwg";
var OUTPUT_SHEET = "output";
var STATEMENTS_SHEET = "statements";

function doGet(e) {
  return handle_(e);
}

function doPost(e) {
  return handle_(e);
}

function handle_(e) {
  var params = (e && e.parameter) || {};
  var callback = String(params.callback || "");
  if (!/^espectro[A-Za-z0-9_]+$/.test(callback)) callback = "";
  var result;
  try {
    if (String(params.ping || "") === "1") result = { ok: true, ping: true };
    else result = save_(params);
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  var body = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + body + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

function save_(params) {
  var email = String(params.email || "").trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("email");
  var scores = String(params.scores || "").split(",");
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var statements = ss.getSheetByName(STATEMENTS_SHEET);
    if (!statements) throw new Error("statements");
    var column = statements.getRange("A:A").getValues();
    var count = 0;
    for (var i = 0; i < column.length; i++) {
      if (String(column[i][0] || "").trim()) count++;
      else if (count) break;
    }
    if (!count) throw new Error("empty");
    var headers = ["email respondent"];
    for (var n = 1; n <= count; n++) headers.push("S" + n);
    var out = ss.getSheetByName(OUTPUT_SHEET);
    if (!out) out = ss.insertSheet(OUTPUT_SHEET);
    var width = headers.length;
    var existing = out.getRange(1, 1, 1, width).getValues()[0];
    var headerOk = true;
    for (var h = 0; h < width; h++) {
      if (String(existing[h] || "") !== headers[h]) headerOk = false;
    }
    if (!headerOk) out.getRange(1, 1, 1, width).setValues([headers]);
    var row = [email];
    for (var s = 0; s < count; s++) {
      var raw = scores[s] == null ? "" : String(scores[s]).trim();
      if (!raw) {
        row.push("");
        continue;
      }
      var num = Number(raw);
      if (!isFinite(num)) row.push("");
      else row.push(Math.max(-100, Math.min(100, Math.round(num))));
    }
    var last = Math.max(out.getLastRow(), 1);
    var emails = last > 1 ? out.getRange(2, 1, last - 1, 1).getValues() : [];
    var found = 0;
    for (var r = 0; r < emails.length; r++) {
      if (String(emails[r][0] || "").trim().toLowerCase() === email) found = r + 2;
    }
    if (found) out.getRange(found, 1, 1, width).setValues([row]);
    else out.getRange(last + 1, 1, 1, width).setValues([row]);
    return { ok: true, email: email, updated: found > 0 };
  } finally {
    lock.releaseLock();
  }
}
