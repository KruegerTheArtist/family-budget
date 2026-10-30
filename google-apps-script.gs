/**
 * Семейный бюджет — бэкенд синхронизации PWA с Google-таблицей «Семейный бюджет».
 *
 * Листы:
 *  - «Таблица расходов»: приложение заполняет только Категория, Расходы, руб., Имя, Комментарии, Дата.
 *    №, Год, Месяц, Группа — формулы таблицы, скрипт лишь протягивает их на новые строки.
 *  - «Таблица доходов»: Имя, Группа, Категория, Доходы, руб., Дата, Комментарии (+ нумерация №).
 *  - «Данные»: справочники (A — группы расходов, B — категории расходов,
 *    J:K — категория→группа расходов, P:Q — категория→группа доходов).
 *  - Имена берутся из выпадающего списка колонки «Имя» таблицы расходов (если его нет — из «Данные»!C)
 *    и из приложения в таблицу не пишутся.
 * Колонки ищутся по заголовкам в первой строке, поэтому их порядок можно менять.
 * Сводки, месячные листы и всё правее таблиц скрипт не трогает.
 *
 * Ревизия (rev) — хэш содержимого таблиц: правка руками прямо в Google-таблице тоже меняет rev,
 * и приложение не перезапишет её молча, а покажет конфликт.
 *
 * Установка:
 *  1. Открой свой Apps Script проект (script.google.com → Мои проекты).
 *  2. Скопируй себе строку `var SECRET = '…'` из текущего кода, замени весь код этим,
 *     верни свой SECRET и впиши SHEET_ID.
 *  3. Deploy → Manage deployments → ✎ → Version: New version → Deploy (URL останется прежним).
 *     Если попросит заново разрешить доступ — разреши.
 *  4. В приложении: Данные → Синхронизация → «Взять из таблицы».
 */

var SECRET   = 'ЗАМЕНИ_НА_СВОЙ_ТОКЕН';                            // тот же токен, что введён в приложении
var SHEET_ID = '12TjAVOilg_F8OIiuAf6aw6svtuRX5yQjCeKIOSo2v9w';     // таблица «Семейный бюджет»

var SH_EXP = 'Таблица расходов', SH_INC = 'Таблица доходов', SH_DATA = 'Данные';
var MONTHS = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];

/* Поля операций → возможные заголовки колонок. Поля, которых нет в спецификации листа, берутся из формул. */
var SPEC_EXP = {
  sheet: SH_EXP, type: 'expense', prefix: 'e',
  fields: {
    category: ['Категория'], amount: ['Расходы, руб.', 'Сумма, руб.'], name: ['Имя'],
    comment: ['Комментарии', 'Комментарий'], date: ['Дата']
  }
};
var SPEC_INC = {
  sheet: SH_INC, type: 'income', prefix: 'i',
  fields: {
    name: ['Имя'], group: ['Группа'], category: ['Категория'], amount: ['Доходы, руб.', 'Доход, руб.'],
    date: ['Дата'], comment: ['Комментарии', 'Комментарий']
  }
};

/* ---------- HTTP ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) {}
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.token !== SECRET) return out_({ ok: false, error: 'bad token' });
    if (req.action === 'ping') return out_({ ok: true, sheet: ss_().getName() });
    if (req.action === 'pull') { var data = readAll_(); return out_({ ok: true, rev: revOf_(data), data: data }); }
    if (req.action === 'push') {
      var cur = revOf_(readAll_());
      if (req.baseRev !== cur && !req.force) return out_({ ok: false, error: 'conflict', rev: cur });
      writeAll_(req.data || {});
      SpreadsheetApp.flush();
      return out_({ ok: true, rev: revOf_(readAll_()) });
    }
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
function doGet() { return out_({ ok: true, info: 'budget sync alive' }); }
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- таблица / ревизия ---------- */
var TZ_ = null; // часовой пояс таблицы: даты пишем и читаем в нём, а не в поясе проекта Apps Script
function ss_() {
  var ss = (SHEET_ID && SHEET_ID.indexOf('ВСТАВЬ') !== 0) ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  TZ_ = ss.getSpreadsheetTimeZone();
  return ss;
}
function sheet_(ss, name) {
  var s = ss.getSheetByName(name);
  if (!s) throw new Error('нет листа «' + name + '»');
  return s;
}
/** Положительное 31-битное число из MD5 содержимого — приложение хранит rev как число. */
function revOf_(data) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(data), Utilities.Charset.UTF_8);
  return (((b[0] & 127) << 24) | ((b[1] & 255) << 16) | ((b[2] & 255) << 8) | (b[3] & 255)) || 1;
}

/* ---------- колонки по заголовкам ---------- */
function norm_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().toLowerCase(); }
/** Ширина таблицы — до первой пустой ячейки шапки (сводки правее, например Z:AD, не задеваются). */
function layout_(s, spec) {
  var lastCol = s.getLastColumn();
  var hdr = s.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
  var width = 0;
  while (width < hdr.length && String(hdr[width]).trim() !== '') width++;
  var cols = {};
  Object.keys(spec.fields).forEach(function (f) {
    var names = spec.fields[f].map(norm_);
    for (var i = 0; i < width; i++) if (names.indexOf(norm_(hdr[i])) >= 0) { cols[f] = i; return; }
    throw new Error('на листе «' + spec.sheet + '» нет колонки «' + spec.fields[f][0] + '»');
  });
  var noCol = -1;
  for (var j = 0; j < width; j++) if (norm_(hdr[j]) === '№') noCol = j;
  return { width: width, cols: cols, groupCol: hdr.slice(0, width).map(norm_).indexOf('группа'), noCol: noCol };
}

/* ---------- даты ---------- */
function z_(n) { n = Number(n); return (n < 10 ? '0' : '') + n; }
function toYMD_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, TZ_, 'yyyy-MM-dd');
  if (typeof v === 'number' && v > 20000 && v < 80000) { // дата пришла серийным числом Sheets
    var sd = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return sd.getUTCFullYear() + '-' + z_(sd.getUTCMonth() + 1) + '-' + z_(sd.getUTCDate());
  }
  var s = String(v || '').trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return m[1] + '-' + m[2] + '-' + m[3];
  if ((m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/))) return m[3] + '-' + z_(m[2]) + '-' + z_(m[1]);
  return '';
}
function fromYMD_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) ? Utilities.parseDate(s, TZ_, 'yyyy-MM-dd') : ''; }

/* ---------- чтение ---------- */
function readAll_() {
  var ss = ss_(), d = sheet_(ss, SH_DATA);
  var expenseCats = readPairs_(d, 10), incomeCats = readPairs_(d, 16);
  var tx = [];
  readTable_(sheet_(ss, SH_EXP), SPEC_EXP, groupMap_(expenseCats), tx);
  readTable_(sheet_(ss, SH_INC), SPEC_INC, groupMap_(incomeCats), tx);
  var names = allowedNames_(sheet_(ss, SH_EXP), SPEC_EXP);
  return { v: 1, tx: tx, expenseCats: expenseCats, incomeCats: incomeCats, names: names.length ? names : readCol_(d, 3) };
}
/** Имена — из выпадающего списка колонки «Имя» в таблице расходов (как разрешено в самой таблице). */
function allowedNames_(s, spec) {
  var rule = s.getRange(2, layout_(s, spec).cols.name + 1).getDataValidation();
  if (!rule) return [];
  var type = rule.getCriteriaType(), arg = rule.getCriteriaValues()[0], vals = [];
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) vals = arg || [];
  else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && arg) vals = [].concat.apply([], arg.getValues());
  var out = [];
  vals.forEach(function (v) { v = String(v || '').trim(); if (v && out.indexOf(v) < 0) out.push(v); });
  return out;
}
function groupMap_(cats) { var m = {}; cats.forEach(function (c) { m[c.name] = c.group; }); return m; }
function readTable_(s, spec, groups, out) {
  var L = layout_(s, spec), last = s.getLastRow();
  if (last < 2) return;
  s.getRange(2, 1, last - 1, L.width).getValues().forEach(function (r, i) {
    var c = L.cols, cat = String(r[c.category] || '').trim(), amount = Number(r[c.amount]) || 0;
    if (!amount && !cat) return; // пустые строки с протянутыми формулами
    var g = c.group != null ? r[c.group] : groups[cat];
    if (g == null && L.groupCol >= 0) g = r[L.groupCol];
    g = String(g || '').trim();
    out.push({
      id: spec.prefix + (i + 2), type: spec.type, date: toYMD_(r[c.date]),
      name: String(r[c.name] || '').trim(), group: g.charAt(0) === '#' ? '' : g, category: cat,
      amount: amount, comment: String(r[c.comment] || '').trim()
    });
  });
}
function readPairs_(s, c1) {
  var out = [], last = s.getLastRow(); if (last < 2) return out;
  s.getRange(2, c1, last - 1, 2).getValues().forEach(function (r) {
    var n = String(r[0] || '').trim(); if (n) out.push({ name: n, group: String(r[1] || '').trim() });
  });
  return out;
}
function readCol_(s, c) {
  var out = [], last = s.getLastRow(); if (last < 2) return out;
  s.getRange(2, c, last - 1, 1).getValues().forEach(function (r) { var n = String(r[0] || '').trim(); if (n) out.push(n); });
  return out;
}

/* ---------- запись (только колонки ввода; формулы протягиваются, сводки не трогаем) ---------- */
function writeAll_(db) {
  var ss = ss_(), tx = db.tx || [];
  // порядок строк — как в приложении: после «Взять из таблицы» он совпадает с таблицей, новые операции в конце
  writeTable_(sheet_(ss, SH_EXP), SPEC_EXP, tx.filter(function (t) { return t.type !== 'income'; }));
  writeTable_(sheet_(ss, SH_INC), SPEC_INC, tx.filter(function (t) { return t.type === 'income'; }));

  var d = sheet_(ss, SH_DATA);
  var ec = db.expenseCats || [], ic = db.incomeCats || [];
  var groups = []; ec.forEach(function (c) { if (c.group && groups.indexOf(c.group) < 0) groups.push(c.group); });
  writeCol_(d, 1, groups);                                              // A: группы расходов
  writeCol_(d, 2, ec.map(function (c) { return c.name; }));               // B: категории расходов
  // C (имена) не пишем: список имён ведётся в самой таблице (выпадающий список колонки «Имя»)
  writeCols_(d, 10, ec.map(function (c) { return [c.name, c.group]; }));   // J:K
  writeCols_(d, 16, ic.map(function (c) { return [c.name, c.group]; }));   // P:Q
}
function writeTable_(s, spec, items) {
  var L = layout_(s, spec), n = items.length, last = s.getLastRow();
  var fields = Object.keys(L.cols);
  if (last > 1) fields.forEach(function (f) { s.getRange(2, L.cols[f] + 1, last - 1, 1).clearContent(); });
  if (!n) return;

  fields.forEach(function (f) {
    var vals = items.map(function (t) {
      if (f === 'amount') return [Number(t.amount) || 0];
      if (f === 'date') return [fromYMD_(t.date)];
      return [t[f] || ''];
    });
    // формат не задаём: в «Таблице» Google колонка «Дата» типизирована, и setNumberFormat там запрещён
    s.getRange(2, L.cols[f] + 1, n, 1).setValues(vals);
  });

  // остальные колонки шапки: формулы из строки 2 протягиваем на все строки, «№» без формулы — нумеруем
  var r1c1 = s.getRange(2, 1, 1, L.width).getFormulasR1C1()[0];
  for (var c = 0; c < L.width; c++) {
    if (fields.some(function (f) { return L.cols[f] === c; })) continue;
    var rg2 = s.getRange(2, c + 1, n, 1);
    if (r1c1[c]) rg2.setFormulasR1C1(items.map(function () { return [r1c1[c]]; }));
    else if (c === L.noCol) rg2.setValues(items.map(function (t, i) { return [i + 1]; }));
  }
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
