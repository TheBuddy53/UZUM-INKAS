/**
 * ИНКАССАЦИЯ ПВЗ — бэкенд сайта (Google Apps Script, веб-приложение). Версия 1.
 * Сайт лежит на GitHub Pages, все запросы с него приходят сюда. Сам скрипт пишет в таблицу и в Google Диск.
 *
 * ЧТО ДЕЛАЕТ
 *  • Принимает отчёт: ПВЗ + фото + подтверждённый вход Telegram. Без входа отправить нельзя — подпись Telegram
 *    проверяется здесь токеном бота (свойство скрипта BOT_TOKEN), подделать Telegram ID нельзя.
 *  • Фото в полном размере сохраняет в Google Диск (папка «Инкассация — фото» / «ММ.ГГГГ»),
 *    в таблицу кладёт превью прямо в ячейку и ссылку на оригинал.
 *  • Каждый месяц — новый лист «ММ.ГГГГ» (создаётся сам при первом отчёте месяца).
 *  • Колонки: Дата | Время | ПВЗ | Telegram ID | Имя в Telegram | Username | Фото | Фото (оригинал) | День | Месяц | Год |
 *             Отметка времени | Номер отчёта
 *  • Справочник точек — лист «ПВЗ», колонка A (функция setup() заполнит его ТУ Светланы, дальше правишь сам).
 *
 * УСТАНОВКА (подробно — в инструкции)
 *  1) Таблица → Расширения → Apps Script → вставить этот код в Code.gs (заменить всё) → Сохранить.
 *  2) Настройки проекта (шестерёнка) → Часовой пояс: (GMT+05:00) Tashkent.
 *     Там же «Свойства скрипта» → Добавить: BOT_TOKEN = токен бота @UZUM_SLA_BOT (из @BotFather).
 *  3) Выбрать функцию setup → Выполнить → разрешить доступ (таблица + Диск). Посмотреть журнал выполнения.
 *  4) Развернуть → Новое развертывание → Веб-приложение → Запуск от: Я, Доступ: Все → Развернуть → скопировать ссылку.
 *  5) Ссылку вставить в файл config.js в репозитории сайта.
 * Обновление кода потом: Развернуть → Управление развертываниями → ✎ → Версия: Новая версия → Развернуть (ссылка та же).
 */

// ==== НАСТРОЙКИ ====
var SPREADSHEET_ID = '1wQF3DwHXCiGCDdIGlT1A7rny80cW4YyLcFa_3YuVvIo'; // таблица «Инкассация»
var TIMEZONE = 'Asia/Tashkent';
var PVZ_SHEET_NAME = 'ПВЗ';
var PHOTO_ROOT_NAME = 'Инкассация — фото';
var MAX_AUTH_AGE_DAYS = 30;            // вход Telegram старше — попросим войти заново
var MAX_PHOTO_BYTES = 4 * 1024 * 1024; // сайт сам сжимает фото до ~1,5 МБ; больше 4 МБ сервер не примет
var MAX_THUMB_BYTES = 250 * 1024;      // превью для ячейки таблицы
var HEADERS = ['Дата', 'Время', 'ПВЗ', 'Telegram ID', 'Имя в Telegram', 'Username', 'Фото', 'Фото (оригинал)',
  'День', 'Месяц', 'Год', 'Отметка времени', 'Номер отчёта'];
var COL = { date: 1, time: 2, pvz: 3, id: 4, name: 5, user: 6, photo: 7, link: 8, d: 9, m: 10, y: 11, ts: 12, no: 13 };
var MONTHS_RU = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];

// ТУ Светланы — стартовое содержимое листа «ПВЗ» (setup() запишет, если лист пуст)
var DEFAULT_PVZ = [
  'FrТАШ-9', 'FrТАШ-160', 'FrТАШ-166', 'FrТАШ-167', 'FrТАШ-169', 'FrТАШ-171', 'FrТАШ-175', 'FrТАШ-176',
  'FrТАШ-179', 'FrТАШ-181', 'FrТАШ-184', 'FrТАШ-185', 'FrТАШ-187', 'FrТАШ-188', 'FrТАШ-189', 'FrТАШ-192',
  'FrТАШ-195', 'FrТАШ-196', 'FrТАШ-198', 'FrТАШ-200', 'FrТАШ-203', 'FrТАШ-205', 'FrТАШ-206', 'FrТАШ-208',
  'FrТАШ-213', 'FrТАШ-215', 'FrТАШ-219', 'FrТАШ-220', 'FrТАШ-225', 'FrТАШ-227', 'FrТАШ-230', 'FrТАШ-232',
  'FrТАШ-237', 'FrТАШ-239', 'FrТАШ-243', 'FrТАШ-247', 'FrТАШ-250', 'FrТАШ-251', 'FrТАШ-258', 'FrТАШ-261',
  'FrТАШ-265', 'FrТАШ-266', 'FrТАШ-269', 'FrТАШ-271', 'FrТАШ-275', 'FrТАШ-277', 'FrТАШ-282', 'FrТАШ-283',
  'FrТАШ-284', 'FrТАШ-285', 'FrТАШ-286', 'FrТАШ-287', 'FrТАШ-292', 'FrТАШ-293', 'FrТАШ-296', 'FrТАШ-297',
  'FrТАШ-298', 'FrТАШ-304', 'FrТАШ-309', 'FrТАШ-310', 'FrТАШ-317', 'FrТАШ-318', 'FrТАШ-321', 'FrТАШ-322',
  'FrТАШ-323', 'FrТАШ-325', 'FrТАШ-330', 'FrТАШ-331', 'FrТАШ-332', 'FrТАШ-333', 'FrТАШ-335', 'FrТАШ-336',
  'FrТАШ-337', 'FrТАШ-338', 'FrТАШ-339', 'FrТАШ-341', 'FrТАШ-342', 'FrТАШ-347', 'FrТАШ-352', 'FrТАШ-353',
  'FrТАШ-355', 'FrТАШ-365', 'FrТАШ-369', 'FrТАШ-370', 'FrТАШ-373', 'FrТАШ-376', 'FrТАШ-378', 'FrТАШ-379',
  'FrТАШ-389', 'FrТАШ-393', 'FrТАШ-395', 'FrТАШ-397', 'FrТАШ-398', 'FrТАШ-401', 'FrТАШ-406', 'FrТАШ-407',
  'FrТАШ-409', 'FrТАШ-411', 'FrТАШ-412', 'FrТАШ-414', 'FrТАШ-415', 'FrТАШ-417', 'FrТАШ-418', 'FrТАШ-419',
  'FrТАШ-426', 'FrТАШ-427', 'FrТАШ-428', 'FrТАШ-429', 'FrТАШ-431', 'FrТАШ-432', 'FrТАШ-435', 'FrТАШ-436',
  'FrТАШ-438', 'FrТАШ-439', 'FrТАШ-442', 'FrТАШ-447', 'FrТАШ-449', 'FrТАШ-450', 'FrТАШ-452', 'FrТАШ-459',
  'FrАТП-1', 'FrБЕК-10', 'FrГСТ-1', 'FrКМН-1', 'FrМРТ-1', 'FrОКТ-1', 'FrСАИ-1', 'FrЧИН-9',
  'FrЯНГ-16'
];

// ==== ВХОДНЫЕ ТОЧКИ ====
function doGet(e) {
  try {
    var a = (e && e.parameter && e.parameter.action) || '';
    if (a === 'pvz_list') return cachedJson_('pvz', 120, function () { return { ok: true, pvzList: pvzList_() }; });
    if (a === 'ink_today') return cachedJson_('today:' + fmt_(new Date(), 'dd.MM.yyyy'), 30, getToday_);
    if (a === 'ping') return json_({ ok: true, v: 1, app: 'ink', token: !!botToken_() });
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.type === 'ink_submit') return handleSubmit_(body);
    if (body.type === 'ink_mine') return handleMine_(body);
    return json_({ ok: false, error: 'unknown type' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ==== ПРОВЕРКА ВХОДА TELEGRAM ====
// Официальный алгоритм Telegram Login Widget: secret = SHA256(токен бота); hash = HMAC-SHA256(строка полей, secret).
function botToken_() {
  var t = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  return t ? String(t).trim() : '';
}

function verifyTelegram_(tg, token, maxDays) {
  if (!tg || typeof tg !== 'object' || !tg.hash || !tg.id || !tg.auth_date) return { ok: false, why: 'no-data' };
  var keys = Object.keys(tg).filter(function (k) {
    return k !== 'hash' && tg[k] !== undefined && tg[k] !== null && typeof tg[k] !== 'object';
  }).sort();
  var dcs = keys.map(function (k) { return k + '=' + tg[k]; }).join('\n');
  var secret = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token, Utilities.Charset.UTF_8);
  var sig = Utilities.computeHmacSha256Signature(Utilities.newBlob('').setDataFromString(dcs, 'UTF-8').getBytes(), secret);
  var hex = sig.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  if (hex !== String(tg.hash).toLowerCase()) return { ok: false, why: 'bad-hash' };
  var age = Date.now() / 1000 - Number(tg.auth_date);
  if (!(age < (maxDays || MAX_AUTH_AGE_DAYS) * 86400) || age < -600) return { ok: false, why: 'old' };
  return { ok: true, id: String(tg.id) };
}

function authOrError_(tg) {
  var tok = botToken_();
  if (!tok) return { err: json_({ ok: false, code: 'setup', error: 'Сервер не настроен: добавь BOT_TOKEN в свойства скрипта' }) };
  var v = verifyTelegram_(tg, tok, MAX_AUTH_AGE_DAYS);
  if (!v.ok) return { err: json_({ ok: false, code: 'auth', error: v.why === 'old' ? 'Вход устарел — войди через Telegram заново' : 'Вход Telegram не подтверждён — войди заново' }) };
  return { id: v.id };
}

// ==== ПРИЁМ ОТЧЁТА ====
function handleSubmit_(d) {
  var au = authOrError_(d.tg);
  if (au.err) return au.err;
  var id = au.id;

  var pvz = normalizePvz_(d.pvz);
  if (!pvz || pvzList_().indexOf(pvz) === -1) return json_({ ok: false, code: 'pvz', error: 'ПВЗ нет в справочнике: ' + pvz });

  var photo = parseImage_(d.photo, MAX_PHOTO_BYTES);
  if (!photo) return json_({ ok: false, code: 'photo', error: 'Фото не прочиталось или слишком большое (до 4 МБ)' });
  var thumb = parseImage_(d.thumb, MAX_THUMB_BYTES);

  var rid = String(d.rid || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(rid)) rid = '';
  if (rid) {
    var dup = cacheGet_('rid:' + rid);
    if (dup) return textOut_(dup); // повтор после обрыва связи — тот же ответ, без второй записи
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var now = new Date();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = monthSheet_(ss, now);
    var dateStr = fmt_(now, 'dd.MM.yyyy'), timeStr = fmt_(now, 'HH:mm');
    var no = 'INK-' + fmt_(now, 'yyMMdd') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 5).toUpperCase();
    var tg = d.tg || {};
    var name = safeText_([tg.first_name, tg.last_name].filter(Boolean).join(' '), 80);
    var user = tg.username ? '@' + safeText_(tg.username, 60) : '';

    // фото в Диск: «Инкассация — фото / 11.2026 / 08.11.2026 14-32 · FrТАШ-160 · INK-….jpg»
    var folder = monthFolder_(now);
    var fname = dateStr + ' ' + timeStr.replace(':', '-') + ' · ' + pvz + ' · ' + no + '.' + photo.ext;
    var file = folder.createFile(Utilities.newBlob(photo.bytes, photo.mime, fname));
    file.setDescription('Инкассация · ' + pvz + ' · Telegram ID ' + id + ' · ' + dateStr + ' ' + timeStr);

    var row = sheet.getLastRow() + 1;
    if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), row - sheet.getMaxRows());
    sheet.getRangeList(['A' + row + ':B' + row, 'D' + row]).setNumberFormat('@'); // текстом: дата, время, ID
    sheet.getRange(row, 1, 1, HEADERS.length).setValues([[
      dateStr, timeStr, pvz, id, name, user, '', '',
      +fmt_(now, 'd'), MONTHS_RU[+fmt_(now, 'M') - 1], +fmt_(now, 'yyyy'), now, no
    ]]);
    sheet.getRange(row, COL.link).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText('Открыть фото').setLinkUrl(file.getUrl()).build());
    var shown = false;
    var img = thumb || (photo.bytes.length <= MAX_THUMB_BYTES ? photo : null);
    if (img) {
      try {
        var ci = SpreadsheetApp.newCellImage().setSourceUrl('data:' + img.mime + ';base64,' + Utilities.base64Encode(img.bytes))
          .setAltTextTitle(pvz + ' · ' + dateStr + ' ' + timeStr).build();
        sheet.getRange(row, COL.photo).setValue(ci);
        shown = true;
      } catch (imgErr) {
        Logger.log('Превью не вставилось: %s', imgErr);
      }
    }
    if (!shown) sheet.getRange(row, COL.photo).setValue('см. ссылку →');
    sheet.setRowHeight(row, 96);

    SpreadsheetApp.flush();
    bumpCache_();
    var res = JSON.stringify({ ok: true, v: 1, no: no, date: dateStr, time: timeStr, pvz: pvz });
    if (rid) cachePut_('rid:' + rid, res, 900);
    Logger.log('Отчёт: лист=%s строка=%s ПВЗ=%s id=%s №=%s', sheet.getName(), row, pvz, id, no);
    return textOut_(res);
  } finally {
    lock.releaseLock();
  }
}

// ==== МОИ ОТЧЁТЫ (вкладка «История») ====
function handleMine_(d) {
  var au = authOrError_(d.tg);
  if (au.err) return au.err;
  var id = au.id, ss = SpreadsheetApp.openById(SPREADSHEET_ID), now = new Date(), out = [];
  var months = [now, new Date(now.getFullYear(), now.getMonth() - 1, 15)];
  months.forEach(function (dt) {
    var sh = ss.getSheetByName(fmt_(dt, 'MM.yyyy'));
    if (!sh || sh.getLastRow() < 2) return;
    var n = sh.getLastRow() - 1;
    var v = sh.getRange(2, 1, n, 4).getValues(), nos = sh.getRange(2, COL.no, n, 1).getValues();
    for (var i = v.length - 1; i >= 0; i--) {
      if (String(v[i][3]).trim() !== id) continue;
      out.push({ date: dateOf_(v[i][0]), time: timeOf_(v[i][1]), pvz: normalizePvz_(v[i][2]), no: String(nos[i][0] || '') });
      if (out.length >= 120) break;
    }
  });
  return json_({ ok: true, v: 1, items: out });
}

// ==== СЕГОДНЯ ПО ТЕРРИТОРИИ (кто сдал / кто нет) ====
function getToday_() {
  var now = new Date(), today = fmt_(now, 'dd.MM.yyyy'), list = pvzList_();
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(fmt_(now, 'MM.yyyy'));
  var done = {}, order = [];
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < v.length; i++) {
      if (dateOf_(v[i][0]) !== today) continue;
      var p = normalizePvz_(v[i][2]);
      if (!done[p]) order.push(p);
      done[p] = timeOf_(v[i][1]); // последняя сдача за день
    }
  }
  return { ok: true, v: 1, date: today, total: list.length, done: order.map(function (p) { return { pvz: p, time: done[p] }; }) };
}

// ==== СПРАВОЧНИК ПВЗ ====
function pvzList_() {
  var hit = cacheGet_('pvzl:' + (cacheGet_('dv') || '0'));
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PVZ_SHEET_NAME), list = [], seen = {};
  if (sh && sh.getLastRow() > 0) {
    sh.getRange(1, 1, sh.getLastRow(), 1).getValues().forEach(function (r) {
      var v = normalizePvz_(r[0]);
      if (v && v !== 'ПВЗ' && !seen[v]) { seen[v] = 1; list.push(v); }
    });
  }
  if (!list.length) list = DEFAULT_PVZ.slice();
  cachePut_('pvzl:' + (cacheGet_('dv') || '0'), JSON.stringify(list), 120);
  return list;
}

// ==== ЛИСТ МЕСЯЦА И ПАПКИ ====
function monthSheet_(ss, dt) {
  var name = fmt_(dt, 'MM.yyyy'), sh = ss.getSheetByName(name);
  if (sh) return sh;
  try {
    sh = ss.insertSheet(name, 0);
  } catch (e) {
    sh = ss.getSheetByName(name);
    if (sh) return sh;
    throw e;
  }
  var head = sh.getRange(1, 1, 1, HEADERS.length);
  head.setValues([HEADERS]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5b17c9')
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);
  var widths = [92, 64, 112, 118, 150, 130, 150, 120, 56, 92, 60, 150, 132];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('A:F').setVerticalAlignment('middle');
  sh.getRange('H:M').setVerticalAlignment('middle');
  if (sh.getMaxColumns() > HEADERS.length) sh.deleteColumns(HEADERS.length + 1, sh.getMaxColumns() - HEADERS.length);
  return sh;
}

function rootFolder_() {
  var props = PropertiesService.getScriptProperties(), fid = props.getProperty('PHOTO_ROOT_ID');
  if (fid) { try { return DriveApp.getFolderById(fid); } catch (e) {} }
  var it = DriveApp.getFoldersByName(PHOTO_ROOT_NAME), f = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_ROOT_NAME);
  props.setProperty('PHOTO_ROOT_ID', f.getId());
  return f;
}

function monthFolder_(dt) {
  var root = rootFolder_(), name = fmt_(dt, 'MM.yyyy'), it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

// ==== ПЕРВЫЙ ЗАПУСК: Выполнить → setup ====
function setup() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('Таблица: %s', ss.getName());
  var sh = ss.getSheetByName(PVZ_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(PVZ_SHEET_NAME);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, DEFAULT_PVZ.length, 1).setValues(DEFAULT_PVZ.map(function (p) { return [p]; }));
    sh.setColumnWidth(1, 120);
    Logger.log('Лист «ПВЗ» заполнен: %s точек ТУ Светланы', DEFAULT_PVZ.length);
  } else {
    Logger.log('Лист «ПВЗ» уже есть: %s строк (не трогаю)', sh.getLastRow());
  }
  var m = monthSheet_(ss, new Date());
  Logger.log('Лист месяца: %s', m.getName());
  // пустой стартовый лист «Лист1»/«Sheet1» убираем, если в нём ничего нет
  ['Лист1', 'Sheet1'].forEach(function (n) {
    var x = ss.getSheetByName(n);
    if (x && x.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(x);
  });
  var f = rootFolder_();
  Logger.log('Папка для фото в Google Диске: «%s» — %s', f.getName(), f.getUrl());
  if (ss.getSpreadsheetTimeZone() !== TIMEZONE) {
    ss.setSpreadsheetTimeZone(TIMEZONE);
    Logger.log('Часовой пояс таблицы выставлен: %s', TIMEZONE);
  }
  Logger.log('Пояс проекта Apps Script: %s %s', Session.getScriptTimeZone(), Session.getScriptTimeZone() === TIMEZONE ? '✓' : '— поменяй в Настройках проекта на Asia/Tashkent');
  Logger.log(botToken_() ? 'BOT_TOKEN задан ✓' : 'ВНИМАНИЕ: BOT_TOKEN не задан — Настройки проекта → Свойства скрипта → BOT_TOKEN');
  bumpCache_();
  Logger.log('Готово. Теперь: Развернуть → Новое развертывание → Веб-приложение (Запуск от: Я, Доступ: Все).');
}

// Проверка токена: Выполнить → checkToken (покажет имя бота, если токен верный)
function checkToken() {
  var t = botToken_();
  if (!t) { Logger.log('BOT_TOKEN не задан'); return; }
  var r = UrlFetchApp.fetch('https://api.telegram.org/bot' + t + '/getMe', { muteHttpExceptions: true });
  Logger.log('Telegram ответил: %s', r.getContentText());
}

// ==== СЛУЖЕБНОЕ ====
function fmt_(d, p) { return Utilities.formatDate(d, TIMEZONE, p); }
function dateOf_(v) { return v instanceof Date ? fmt_(v, 'dd.MM.yyyy') : String(v == null ? '' : v).trim(); }
function timeOf_(v) { return v instanceof Date ? fmt_(v, 'HH:mm') : String(v == null ? '' : v).trim(); }

function normalizePvz_(v) {
  return String(v == null ? '' : v).replace(/[​‌‍﻿]/g, '').replace(/ /g, ' ').trim().replace(/\s+/g, ' ');
}

function safeText_(v, maxLen) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim();
  if (maxLen && s.length > maxLen) s = s.slice(0, maxLen);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // защита от формул в ячейках
}

// data:image/...;base64,... -> { bytes, mime, ext } (проверяем и заголовок, и реальные первые байты файла)
function parseImage_(dataUrl, maxBytes) {
  var m = /^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+\/=\s]+)$/.exec(String(dataUrl || ''));
  if (!m) return null;
  if (m[2].length > maxBytes * 1.4 + 100) return null;
  var bytes;
  try { bytes = Utilities.base64Decode(m[2].replace(/\s/g, '')); } catch (e) { return null; }
  if (!bytes || !bytes.length || bytes.length > maxBytes) return null;
  var b = function (i) { return bytes[i] & 0xff; };
  var jpg = b(0) === 0xFF && b(1) === 0xD8 && b(2) === 0xFF;
  var png = b(0) === 0x89 && b(1) === 0x50 && b(2) === 0x4E && b(3) === 0x47;
  var webp = b(0) === 0x52 && b(1) === 0x49 && b(2) === 0x46 && b(3) === 0x46 && b(8) === 0x57 && b(9) === 0x45;
  if (!(jpg || png || webp)) return null;
  var ext = jpg ? 'jpg' : png ? 'png' : 'webp';
  return { bytes: bytes, mime: 'image/' + (jpg ? 'jpeg' : ext), ext: ext };
}

function cacheGet_(k) { try { return CacheService.getScriptCache().get(k); } catch (e) { return null; } }
function cachePut_(k, v, ttl) { try { CacheService.getScriptCache().put(k, v, ttl); } catch (e) {} }
function bumpCache_() { cachePut_('dv', Date.now() + '' + Math.floor(Math.random() * 1e6), 21600); }
function cachedJson_(key, ttl, compute) {
  var k = 'c:' + key + ':' + (cacheGet_('dv') || '0'), hit = cacheGet_(k);
  if (hit) return textOut_(hit);
  var obj = compute(), text = JSON.stringify(obj);
  if (obj && obj.ok && text.length < 90000) cachePut_(k, text, ttl);
  return textOut_(text);
}
// Ручной сброс кэша (если правил лист «ПВЗ» руками и хочешь увидеть сразу): Выполнить → clearCache
function clearCache() { bumpCache_(); Logger.log('Кэш сброшен'); }

function textOut_(t) { return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.JSON); }
function json_(o) { return textOut_(JSON.stringify(o)); }
