/**
 * Семейный бюджет — бэкенд синхронизации PWA с ТВОЕЙ Google-таблицей
 * (структура как в исходном Excel: листы «Таблица расходов», «Таблица доходов», «Данные»).
 *
 * Приложение пишет операции в A:I этих листов и справочники в лист «Данные»
 * (A — группы расходов, B — категории расходов, C — имена, J:K — категория→группа расходов,
 * P:Q — категория→группа доходов). Остальные колонки/листы (сводки, диаграммы, месяцы) не трогает.
 *
 * Установка:
 *  1. Загрузи Raskhody_list.xlsx на Google Диск → открой → Файл → «Сохранить как таблицу Google».
 *  2. Скопируй ID таблицы из адреса: docs.google.com/spreadsheets/d/<ВОТ_ЭТО>/edit
 *  3. Открой свой Apps Script проект (Расширения → Apps Script в любой таблице или script.google.com),
 *     замени весь код этим, впиши SHEET_ID и свой SECRET (тот же токен, что в приложении).
 *  4. Deploy → New deployment → Web app: Execute as Me, Who has access Anyone → Deploy → Authorize.
 *     (Если проект уже задеплоен: Deploy → Manage deployments → ✎ → Version: New version → Deploy.
 *      URL останется прежним. Если попросит заново разрешить доступ — разреши.)
 *  5. В приложении: Данные → Синхронизация → «Взять из таблицы».
 */

var SECRET   = 'ЗАМЕНИ_НА_СВОЙ_ТОКЕН';   // тот же токен, что введён в приложении
var SHEET_ID = 'ВСТАВЬ_ID_ТАБЛИЦЫ';      // из адреса таблицы (между /d/ и /edit)

var SH_EXP = 'Таблица расходов', SH_INC = 'Таблица доходов', SH_DATA = 'Данные';
var MONTHS = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
var HDR_EXP = ['№','Год','Месяц','Имя','Группа','Категория','Сумма, руб.','Комментарий','Дата'];
var HDR_INC = ['№','Год','Месяц','Имя','Группа','Категория','Доход, руб.','Дата','Комментарий'];

/* ---------- HTTP ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) {}
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.token !== SECRET) return out_({ ok: false, error: 'bad token' });
    if (req.action === 'ping') return out_({ ok: true, rev: getRev_(), sheet: ss_().getName() });
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

/* ---------- таблица / ревизия ---------- */
function ss_() {
  if (SHEET_ID && SHEET_ID.indexOf('ВСТАВЬ') !== 0) return SpreadsheetApp.openById(SHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function getRev_() { var p = PropertiesService.getScriptProperties().getProperty('rev'); return p ? Number(p) : 0; }
function bumpRev_() { var r = getRev_() + 1; PropertiesService.getScriptProperties().setProperty('rev', String(r)); return r; }

/* ---------- даты ---------- */
function z_(n) { n = Number(n); return (n < 10 ? '0' : '') + n; }
function fmt_(d) { return d.getFullYear() + '-' + z_(d.getMonth() + 1) + '-' + z_(d.getDate()); }
function toYMD_(v, year, monthName) {
  if (v instanceof Date && !isNaN(v)) return fmt_(v);
  if (typeof v === 'number' && v > 20000 && v < 80000) { // дата пришла серийным числом Sheets
    var sd = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return sd.getUTCFullYear() + '-' + z_(sd.getUTCMonth() + 1) + '-' + z_(sd.getUTCDate());
  }
  var s = String(v || '').trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return m[1] + '-' + m[2] + '-' + m[3];
  if ((m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/))) return m[3] + '-' + z_(m[2]) + '-' + z_(m[1]);
  var y = parseInt(year, 10), mi = MONTHS.indexOf(String(monthName || '').trim());
  if (y && mi >= 0) return y + '-' + z_(mi + 1) + '-01';
  return '';
}
function fromYMD_(s) { var p = String(s || '').split('-'); if (p.length < 3) return ''; return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])); }
function yr_(ymd) { return Number(String(ymd || '').slice(0, 4)) || ''; }
function mon_(ymd) { var m = Number(String(ymd || '').slice(5, 7)); return (m >= 1 && m <= 12) ? MONTHS[m - 1] : ''; }

/* ---------- чтение ---------- */
function readAll_() {
  var ss = ss_(), tx = [];
  readTable_(ss.getSheetByName(SH_EXP), 'expense', { amount: 6, comment: 7, date: 8 }, tx);
  readTable_(ss.getSheetByName(SH_INC), 'income',  { amount: 6, comment: 8, date: 7 }, tx);
  var d = ss.getSheetByName(SH_DATA);
  return { v: 1, tx: tx, expenseCats: readPairs_(d, 10), incomeCats: readPairs_(d, 16), names: readCol_(d, 3) };
}
function readTable_(s, type, col, out) {
  if (!s) return;
  var last = s.getLastRow(); if (last < 2) return;
  s.getRange(2, 1, last - 1, 9).getValues().forEach(function (r, i) {
    var amount = Number(r[col.amount]) || 0, cat = String(r[5] || '').trim();
    if (!amount && !cat) return;
    out.push({
      id: (type === 'income' ? 'i' : 'e') + (i + 2), type: type,
      date: toYMD_(r[col.date], r[1], r[2]),
      name: String(r[3] || '').trim(), group: String(r[4] || '').trim(), category: cat,
      amount: amount, comment: String(r[col.comment] || '').trim()
    });
  });
}
function readPairs_(s, c1) {
  var out = []; if (!s) return out;
  var last = s.getLastRow(); if (last < 2) return out;
  s.getRange(2, c1, last - 1, 2).getValues().forEach(function (r) {
    var n = String(r[0] || '').trim(); if (n) out.push({ name: n, group: String(r[1] || '').trim() });
  });
  return out;
}
function readCol_(s, c) {
  var out = []; if (!s) return out;
  var last = s.getLastRow(); if (last < 2) return out;
  s.getRange(2, c, last - 1, 1).getValues().forEach(function (r) { var n = String(r[0] || '').trim(); if (n) out.push(n); });
  return out;
}

/* ---------- запись (только свои колонки, формулы/сводки не трогаем) ---------- */
function writeAll_(db) {
  var ss = ss_();
  var tx = (db.tx || []).slice();
  var byDate = function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; };
  var exp = tx.filter(function (t) { return t.type !== 'income'; }).sort(byDate);
  var inc = tx.filter(function (t) { return t.type === 'income'; }).sort(byDate);

  writeTable_(sheet_(ss, SH_EXP), HDR_EXP, exp, 9, function (t, i) {
    return [i + 1, yr_(t.date), mon_(t.date), t.name || '', t.group || '', t.category || '', Number(t.amount) || 0, t.comment || '', fromYMD_(t.date)];
  });
  writeTable_(sheet_(ss, SH_INC), HDR_INC, inc, 8, function (t, i) {
    return [i + 1, yr_(t.date), mon_(t.date), t.name || '', t.group || '', t.category || '', Number(t.amount) || 0, fromYMD_(t.date), t.comment || ''];
  });

  var d = sheet_(ss, SH_DATA);
  var ec = db.expenseCats || [], ic = db.incomeCats || [], names = db.names || [];
  var groups = []; ec.forEach(function (c) { if (c.group && groups.indexOf(c.group) < 0) groups.push(c.group); });
  if (!d.getRange(1, 1).getValue()) d.getRange(1, 1, 1, 3).setValues([['Группа', 'Категория', 'Имя']]);
  if (!d.getRange(1, 10).getValue()) d.getRange(1, 10, 1, 2).setValues([['Категория', 'Группа']]);
  if (!d.getRange(1, 16).getValue()) d.getRange(1, 16, 1, 2).setValues([['Категория', 'Группа']]);
  writeCol_(d, 1, groups);                                           // A: группы расходов
  writeCol_(d, 2, ec.map(function (c) { return c.name; }));            // B: категории расходов
  writeCol_(d, 3, names);                                            // C: имена
  writeCols_(d, 10, ec.map(function (c) { return [c.name, c.group]; })); // J:K
  writeCols_(d, 16, ic.map(function (c) { return [c.name, c.group]; })); // P:Q
}
function writeTable_(s, header, items, dateCol, rowFn) {
  if (!s.getRange(1, 1).getValue()) s.getRange(1, 1, 1, 9).setValues([header]).setFontWeight('bold');
  var last = s.getLastRow();
  if (last > 1) s.getRange(2, 1, last - 1, 9).clearContent();
  if (!items.length) return;
  var vals = items.map(rowFn);
  s.getRange(2, 1, vals.length, 9).setValues(vals);
  s.getRange(2, dateCol, vals.length, 1).setNumberFormat('dd.MM.yyyy');
}
function writeCol_(s, col, arr) {
  var last = s.getLastRow();
  if (last > 1) s.getRange(2, col, last - 1, 1).clearContent();
  if (arr.length) s.getRange(2, col, arr.length, 1).setValues(arr.map(function (v) { return [v]; }));
}
function writeCols_(s, col, rows) {
  var last = s.getLastRow();
  if (last > 1) s.getRange(2, col, last - 1, 2).clearContent();
  if (rows.length) s.getRange(2, col, rows.length, 2).setValues(rows);
}
