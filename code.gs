/*************************************************************
 * ボウリングコンペ集計アプリ  ―  GAS バックエンド (code.gs)
 *
 * 使い方:
 *  1. 専用の Google スプレッドシートを新規作成
 *  2. 拡張機能 > Apps Script を開き、このコードを全文貼り付け
 *  3. デプロイ > 新しいデプロイ > 種類「ウェブアプリ」
 *       - 次のユーザーとして実行: 自分
 *       - アクセスできるユーザー: 全員
 *  4. 発行された /exec で終わる URL を index.html の GAS_URL に貼る
 *
 *  シートは初回アクセス時に自動生成されます。
 *  ※ CORS 回避のため、フロントは全リクエストを GET で送信します。
 *************************************************************/

var SHEET_MEMBERS  = 'members';
var SHEET_SCORES   = 'scores';
var SHEET_LEDGER   = 'ledger';
var SHEET_SETTINGS = 'settings';

function doGet(e)  { return handle_(e); }

function doPost(e) {
  // 念のため POST も受ける（本アプリは主に GET を使用）
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      e.parameter = e.parameter || {};
      if (body.action) e.parameter.action = body.action;
      if (body.data)   e.parameter.data   = JSON.stringify(body.data);
    } catch (err) {}
  }
  return handle_(e);
}

function handle_(e) {
  var action = (e && e.parameter && e.parameter.action) || 'load';
  var data = {};
  try {
    if (e && e.parameter && e.parameter.data) data = JSON.parse(e.parameter.data);
  } catch (err) {
    return json_({ ok: false, error: 'データの解析に失敗しました' });
  }
  try {
    ensureSheets_();
    switch (action) {
      case 'load':              break;
      case 'addMember':         addMember_(data.name);            break;
      case 'removeMember':      removeMember_(data.name);         break;
      case 'saveScore':         saveScore_(data);                 break;
      case 'deleteScore':       deleteScore_(data);               break;
      case 'addLedger':         addLedger_(data);                 break;
      case 'deleteLedger':      deleteLedger_(data.id);           break;
      case 'setInitialBalance': setInitialBalance_(data.amount);  break;
      default: return json_({ ok: false, error: '不明なアクション: ' + action });
    }
    return json_({ ok: true, data: loadAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function ensureSheets_() {
  var s = ss_();
  if (!s.getSheetByName(SHEET_MEMBERS)) {
    s.insertSheet(SHEET_MEMBERS).getRange(1, 1, 1, 1).setValues([['名前']]);
  }
  if (!s.getSheetByName(SHEET_SCORES)) {
    s.insertSheet(SHEET_SCORES).getRange(1, 1, 1, 7)
      .setValues([['年', '月', '名前', 'スコア1', 'スコア2', 'スコア3', '更新日時']]);
  }
  if (!s.getSheetByName(SHEET_LEDGER)) {
    s.insertSheet(SHEET_LEDGER).getRange(1, 1, 1, 9)
      .setValues([['ID', '年', '月', '区分', '費目', '金額', 'メモ', '記録者', '日時']]);
  }
  if (!s.getSheetByName(SHEET_SETTINGS)) {
    var set = s.insertSheet(SHEET_SETTINGS);
    set.getRange(1, 1, 1, 2).setValues([['キー', '値']]);
    set.getRange(2, 1, 1, 2).setValues([['初期残高', 0]]);
  }
}

function values_(name) {
  return ss_().getSheetByName(name).getDataRange().getValues();
}

function loadAll_() {
  // members
  var mVals = values_(SHEET_MEMBERS);
  var members = [];
  for (var i = 1; i < mVals.length; i++) {
    var nm = String(mVals[i][0]).trim();
    if (nm) members.push(nm);
  }
  // scores
  var sVals = values_(SHEET_SCORES);
  var scores = [];
  for (var i = 1; i < sVals.length; i++) {
    var r = sVals[i];
    if (String(r[2]).trim() === '') continue;
    scores.push({
      year:  Number(r[0]),
      month: Number(r[1]),
      name:  String(r[2]).trim(),
      s1: (r[3] === '' || r[3] == null) ? null : Number(r[3]),
      s2: (r[4] === '' || r[4] == null) ? null : Number(r[4]),
      s3: (r[5] === '' || r[5] == null) ? null : Number(r[5])
    });
  }
  // ledger
  var lVals = values_(SHEET_LEDGER);
  var ledger = [];
  for (var i = 1; i < lVals.length; i++) {
    var r = lVals[i];
    if (String(r[0]).trim() === '') continue;
    ledger.push({
      id:       String(r[0]),
      year:     Number(r[1]),
      month:    Number(r[2]),
      type:     String(r[3]),
      category: String(r[4]),
      amount:   Number(r[5]) || 0,
      memo:     String(r[6]),
      recorder: String(r[7])
    });
  }
  // settings
  var setVals = values_(SHEET_SETTINGS);
  var initialBalance = 0;
  for (var i = 1; i < setVals.length; i++) {
    if (String(setVals[i][0]).trim() === '初期残高') initialBalance = Number(setVals[i][1]) || 0;
  }
  return {
    members: members,
    scores: scores,
    ledger: ledger,
    settings: { initialBalance: initialBalance }
  };
}

function addMember_(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('名前が空です');
  var sh = ss_().getSheetByName(SHEET_MEMBERS);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === name) return; // 既に存在
  }
  sh.appendRow([name]);
}

function removeMember_(name) {
  name = String(name || '').trim();
  var sh = ss_().getSheetByName(SHEET_MEMBERS);
  var vals = sh.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]).trim() === name) sh.deleteRow(i + 1);
  }
}

function saveScore_(d) {
  var year = Number(d.year), month = Number(d.month), name = String(d.name || '').trim();
  if (!name)         throw new Error('名前を選択してください');
  if (!year || !month) throw new Error('年月が不正です');
  var s1 = (d.s1 === '' || d.s1 == null) ? '' : Number(d.s1);
  var s2 = (d.s2 === '' || d.s2 == null) ? '' : Number(d.s2);
  var s3 = (d.s3 === '' || d.s3 == null) ? '' : Number(d.s3);
  var sh = ss_().getSheetByName(SHEET_SCORES);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (Number(vals[i][0]) === year && Number(vals[i][1]) === month &&
        String(vals[i][2]).trim() === name) {
      sh.getRange(i + 1, 4, 1, 4).setValues([[s1, s2, s3, new Date()]]);
      return;
    }
  }
  sh.appendRow([year, month, name, s1, s2, s3, new Date()]);
}

function deleteScore_(d) {
  var year = Number(d.year), month = Number(d.month), name = String(d.name || '').trim();
  var sh = ss_().getSheetByName(SHEET_SCORES);
  var vals = sh.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (Number(vals[i][0]) === year && Number(vals[i][1]) === month &&
        String(vals[i][2]).trim() === name) sh.deleteRow(i + 1);
  }
}

function addLedger_(d) {
  var sh = ss_().getSheetByName(SHEET_LEDGER);
  sh.appendRow([
    Utilities.getUuid(),
    Number(d.year),
    Number(d.month),
    String(d.type),               // 収入 / 支出
    String(d.category || ''),
    Number(d.amount) || 0,
    String(d.memo || ''),
    String(d.recorder || ''),
    new Date()
  ]);
}

function deleteLedger_(id) {
  id = String(id || '');
  var sh = ss_().getSheetByName(SHEET_LEDGER);
  var vals = sh.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === id) sh.deleteRow(i + 1);
  }
}

function setInitialBalance_(amount) {
  amount = Number(amount) || 0;
  var sh = ss_().getSheetByName(SHEET_SETTINGS);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === '初期残高') {
      sh.getRange(i + 1, 2).setValue(amount);
      return;
    }
  }
  sh.appendRow(['初期残高', amount]);
}
