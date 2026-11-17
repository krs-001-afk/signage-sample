/**
 * サイネージ配置管理（サンプル）— 共有保存用 Apps Script
 *
 * このスクリプトは、サイト（GitHub Pages）とスプレッドシートの間でデータを中継します。
 * - 「_data」シート：サイトのデータ本体（触らないでください）
 * - 「台帳」シート：保存のたびに自動で書き出される一覧（見る用。ここを編集してもサイトには反映されません）
 *
 * 設定するのは下の EDIT_KEY（合言葉）だけです。
 */
const EDIT_KEY = 'ここに合言葉を入れてください';

const SHEET_DATA = '_data';
const SHEET_VIEW = '台帳';
const CHUNK = 40000; // 1セルに入れる文字数（セルの上限は5万文字）

function doGet(e) {
  return respond_(function () {
    checkKey_(e.parameter.key);
    return readData_();
  });
}

function doPost(e) {
  return respond_(function () {
    var body = JSON.parse(e.postData.contents);
    checkKey_(body.key);
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var cur = readData_();
      if (Number(body.baseRev || 0) !== cur.rev) {
        return { ok: false, error: 'conflict', rev: cur.rev, updatedAt: cur.updatedAt, updatedBy: cur.updatedBy };
      }
      var rev = cur.rev + 1;
      var updatedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      writeData_(body.data, rev, updatedAt, String(body.user || ''));
      if (Array.isArray(body.table)) writeView_(body.table, updatedAt, String(body.user || ''));
      return { ok: true, rev: rev, updatedAt: updatedAt };
    } finally {
      lock.releaseLock();
    }
  });
}

function respond_(fn) {
  var out;
  try { out = fn(); } catch (err) { out = { ok: false, error: String((err && err.message) || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function checkKey_(k) {
  if (!EDIT_KEY || EDIT_KEY === 'ここに合言葉を入れてください') throw new Error('key_not_set');
  if (String(k || '') !== EDIT_KEY) throw new Error('bad_key');
}

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function readData_() {
  var s = sheet_(SHEET_DATA);
  var meta = s.getRange(1, 1, 1, 3).getDisplayValues()[0];
  var rev = Number(meta[0]) || 0;
  var last = s.getLastRow();
  var json = '';
  if (last >= 2) json = s.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }).join('');
  return { ok: true, rev: rev, updatedAt: meta[1] || '', updatedBy: meta[2] || '', data: json ? JSON.parse(json) : null };
}

function writeData_(data, rev, updatedAt, user) {
  var s = sheet_(SHEET_DATA);
  var json = JSON.stringify(data);
  var rows = [];
  // 先頭に ' を付けて、数式や数値として解釈されないようにする
  for (var i = 0; i < json.length; i += CHUNK) rows.push(["'" + json.slice(i, i + CHUNK)]);
  s.clearContents();
  s.getRange(1, 1, 1, 3).setValues([[rev, "'" + updatedAt, "'" + user]]);
  if (rows.length) s.getRange(2, 1, rows.length, 1).setValues(rows);
}

function writeView_(table, updatedAt, user) {
  var s = sheet_(SHEET_VIEW);
  var width = 0;
  table.forEach(function (r) { width = Math.max(width, r.length); });
  if (!width) return;
  var values = table.map(function (r) {
    var row = [];
    for (var i = 0; i < width; i++) {
      var v = r[i] == null ? '' : String(r[i]);
      row.push(/^[=+\-@']/.test(v) ? "'" + v : v);
    }
    return row;
  });
  s.clear();
  s.getRange(1, 1).setValue('最終保存：' + updatedAt + (user ? '（' + user + '）' : '') + '　※この一覧は自動で書き出されます。編集はサイトで行ってください。');
  s.getRange(3, 1, values.length, width).setNumberFormat('@').setValues(values);
  s.getRange(3, 1, 1, width).setFontWeight('bold').setBackground('#f1f3f6');
  s.setFrozenRows(3);
}
