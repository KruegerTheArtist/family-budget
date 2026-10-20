/**
 * Семейный бюджет — бэкенд синхронизации PWA с Google-таблицей.
 *
 * Установка:
 *  1. Создай новую Google-таблицу (sheets.new).
 *  2. Расширения → Apps Script. Удали весь код, вставь этот.
 *  3. Замени ТОКЕН ниже на свой (любая длинная случайная строка).
 *  4. Deploy → New deployment → тип "Web app":
 *       Execute as: Me,  Who has access: Anyone.
 *     Разреши доступ (Authorize), когда попросит.
 *  5. Скопируй "Web app URL" (…/exec).
 *  6. В приложении: Данные → Синхронизация → Подключить таблицу →
 *     вставь URL и этот же токен. Нажми "Проверить подключение", потом "Сохранить".
 *
 * При изменении кода делай Deploy → Manage deployments → Edit → New version.
 */

var SECRET = 'ЗАМЕНИ_НА_СВОЙ_ТОКЕН';   // <-- придумай свою строку, напр. "budget-7xQ2v9mK"

var SH_TX = 'Операции', SH_EC = 'Категории расходов', SH_IC = 'Категории доходов', SH_NM = 'Люди';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) {}
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.token !== SECRET) return out_({ ok: false, error: 'bad token' });
    if (req.action === 'ping') return out_({ ok: true, rev: getRev_() });
    if (req.action === 'pull') return out_({ ok: true, rev: getRev_(), data: readAll_() });
    if (req.action === 'push') {
      if (typeof req.baseRev === 'number' && req.baseRev !== getRev_() && !req.force)
        return out_({ ok: false, error: 'conflict', rev: getRev_() });
      writeAll_(req.data || {});
      return out_({ ok: true, rev: bumpRev_() });
    }
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doGet() { return out_({ ok: true, info: 'budget sync alive', rev: getRev_() }); }

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) { var s = ss_().getSheetByName(name); return s || ss_().insertSheet(name); }
function getRev_() { var p = PropertiesService.getDocumentProperties().getProperty('rev'); return p ? Number(p) : 0; }
function bumpRev_() { var r = getRev_() + 1; PropertiesService.getDocumentProperties().setProperty('rev', String(r)); return r; }

function fmtDate_(v) {
  if (v instanceof Date) {
    var z = function (n) { return (n < 10 ? '0' : '') + n; };
    return v.getFullYear() + '-' + z(v.getMonth() + 1) + '-' + z(v.getDate());
  }
  return String(v || '');
}
function typeRu_(t) { return t === 'income' ? 'Доход' : 'Расход'; }
function typeCode_(s) {
  s = String(s || '').toLowerCase();
  return (s.indexOf('дох') === 0 || s.indexOf('+') === 0 || s === 'income') ? 'income' : 'expense';
}

function readAll_() {
  var tx = [], s = ss_().getSheetByName(SH_TX);
  if (s && s.getLastRow() > 1) {
    s.getRange(2, 1, s.getLastRow() - 1, 8).getValues().forEach(function (r) {
      if (!r[0] && !r[2] && !r[3]) return;
      tx.push({
        date: fmtDate_(r[0]), type: typeCode_(r[1]), amount: Number(r[2]) || 0,
        category: String(r[3] || ''), group: String(r[4] || ''), name: String(r[5] || ''),
        comment: String(r[6] || ''), id: String(r[7] || ('x' + Math.random().toString(36).slice(2)))
      });
    });
  }
  return { v: 1, tx: tx, expenseCats: readPairs_(SH_EC), incomeCats: readPairs_(SH_IC), names: readNames_() };
}
function readPairs_(name) {
  var out = [], s = ss_().getSheetByName(name);
  if (s && s.getLastRow() > 1)
    s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0]) out.push({ name: String(r[0]), group: String(r[1] || '') });
    });
  return out;
}
function readNames_() {
  var out = [], s = ss_().getSheetByName(SH_NM);
  if (s && s.getLastRow() > 1)
    s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) out.push(String(r[0])); });
  return out;
}

function writeAll_(db) {
  var tx = db.tx || [], ec = db.expenseCats || [], ic = db.incomeCats || [], nm = db.names || [];
  var s = sheet_(SH_TX);
  s.clear();
  s.getRange('A:A').setNumberFormat('@'); // даты как текст, чтобы Sheets не ломал формат
  s.getRange(1, 1, 1, 8).setValues([['Дата', 'Тип', 'Сумма', 'Категория', 'Группа', 'Кто', 'Комментарий', 'id']]).setFontWeight('bold');
  if (tx.length) {
    var vals = tx.map(function (t) {
      return [t.date, typeRu_(t.type), Number(t.amount) || 0, t.category || '', t.group || '', t.name || '', t.comment || '', t.id || ''];
    });
    s.getRange(2, 1, vals.length, 8).setValues(vals);
  }
  writePairs_(SH_EC, ec);
  writePairs_(SH_IC, ic);
  var sn = sheet_(SH_NM);
  sn.clear();
  sn.getRange(1, 1, 1, 1).setValues([['Имя']]).setFontWeight('bold');
  if (nm.length) sn.getRange(2, 1, nm.length, 1).setValues(nm.map(function (n) { return [n]; }));
}
function writePairs_(name, arr) {
  var s = sheet_(name);
  s.clear();
  s.getRange(1, 1, 1, 2).setValues([['Категория', 'Группа']]).setFontWeight('bold');
  if (arr.length) s.getRange(2, 1, arr.length, 2).setValues(arr.map(function (c) { return [c.name, c.group]; }));
}
