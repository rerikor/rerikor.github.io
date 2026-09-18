/**
 * Сбор оценок с ai_art_scorer.html в Google Таблицу + отдача статистики для /stats.
 *
 * doPost — принимает результат прохождения квиза с сайта и дописывает строку на лист "Scores".
 * doGet  — отдаёт все строки в JSON (или JSONP, если передан ?callback=) для страницы rerikor.ru/stats.
 *
 * ВАЖНО: после любой правки этого файла нужно не просто сохранить его, а обновить развертывание:
 * «Развернуть» → «Управление развертываниями» → карандаш у нужного развертывания →
 * Версия: «Новая версия» → «Развернуть». Ссылка /exec при этом не меняется.
 */

var SHEET_NAME = 'Scores';
var HEADERS = ['Timestamp', 'Judge', 'Scores (JSON)', 'Sum', 'Count', 'Average'];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var res = writeRow_(data);
    console.log('OK: ' + JSON.stringify(res));
    return json_(res, null);
  } catch (err) {
    console.error('doPost error: ' + err);
    return json_({ ok: false, error: String(err) }, null);
  }
}

function doGet(e) {
  var cb = (e && e.parameter && e.parameter.callback) || null;
  try {
    return json_({ ok: true, rows: readRows_(), generated: new Date().toISOString() }, cb);
  } catch (err) {
    return json_({ ok: false, error: String(err) }, cb);
  }
}

function readRows_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0] && !r[2]) continue;
    var scores = [];
    try { scores = JSON.parse(r[2]); } catch (ignore) {}
    if (!Array.isArray(scores) || !scores.length) continue;
    rows.push({
      ts: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      judge: String(r[1] || ''),
      scores: scores
    });
  }
  return rows;
}

function writeRow_(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS.concat(['Средняя по всем:']));
      sheet.getRange('H1').setFormula('=IFERROR(AVERAGE(F2:F);0)');
      sheet.setFrozenRows(1);
    }
    var scores = Array.isArray(data.scores) ? data.scores : [];
    var sum = scores.reduce(function (a, b) { return a + Number(b); }, 0);
    var count = scores.length;
    var avg = count ? Math.round(sum / count * 100) / 100 : 0;
    sheet.appendRow([new Date(), data.judgeName || '(без имени)', JSON.stringify(scores), sum, count, avg]);
    SpreadsheetApp.flush();
    return { ok: true, spreadsheet: ss.getUrl(), sheet: sheet.getName(), rows: sheet.getLastRow() };
  } finally {
    lock.releaseLock();
  }
}

function json_(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
