/**
 * Google Apps Script для сбора оценок с ai_art_scorer.html в Google Таблицу.
 *
 * Как подключить (один раз, ~5 минут):
 * 1. Создай новую Google Таблицу (sheets.new).
 * 2. В ней открой Расширения → Apps Script.
 * 3. Удали содержимое редактора и вставь целиком этот файл.
 * 4. Нажми "Развернуть" (Deploy) → "Новое развертывание" (New deployment).
 *    - Тип: "Веб-приложение" (Web app).
 *    - Execute as: "Me" (твой аккаунт).
 *    - Who has access: "Anyone" (иначе браузеры посетителей не смогут отправить POST).
 * 5. Скопируй выданный URL вида
 *    https://script.google.com/macros/s/XXXXXXXXXXXX/exec
 * 6. Вставь этот URL в ai_art_scorer.html вместо
 *    'https://script.google.com/macros/s/AKfycby-REPLACE-WITH-YOUR-DEPLOYMENT-ID/exec'
 *    (переменная SCORES_ENDPOINT в начале <script>).
 * 7. Готово — при каждом прохождении квиза строка с оценками будет добавляться
 *    на лист "Scores". Столбец Average — средняя оценка одного прохождения,
 *    а в ячейке H1 автоматически считается средняя по ВСЕМ прохождениям.
 *
 * Если позже понадобится поменять таблицу/логику — правь только этот файл
 * в редакторе Apps Script и делай "Manage deployments" → редактировать
 * существующее развертывание (тогда URL не меняется).
 */

function doPost(e) {
  var sheet = getOrCreateSheet_();
  var data = JSON.parse(e.postData.contents);

  var scores = Array.isArray(data.scores) ? data.scores : [];
  var sum = typeof data.sum === 'number' ? data.sum : scores.reduce(function (a, b) { return a + b; }, 0);
  var count = typeof data.count === 'number' ? data.count : scores.length;
  var average = typeof data.average === 'number' ? data.average : (count ? Math.round((sum / count) * 100) / 100 : 0);

  sheet.appendRow([
    new Date(),
    data.judgeName || '(без имени)',
    JSON.stringify(scores),
    sum,
    count,
    average
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Scores');
  if (!sheet) {
    sheet = ss.insertSheet('Scores');
    sheet.appendRow(['Timestamp', 'Judge', 'Scores (JSON)', 'Sum', 'Count', 'Average']);
    // Средняя оценка по всем прохождениям — всегда актуальна благодаря формуле.
    sheet.getRange('H1').setFormula('=IFERROR(AVERAGE(F2:F), 0)');
    sheet.getRange('G1').setValue('Средняя по всем:');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
