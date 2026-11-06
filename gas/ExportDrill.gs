/**
 * 技術士 問題Ⅰドリル — スプレッドシート → JSON 書き出し（GAS）
 * ---------------------------------------------------------------------------
 * RCCM択一（ExportToJson.gs）の形式を流用し、問題Ⅰドリル用に作り直したもの。
 * RCCMのファイルとは別物。RCCMのスプレッドシートには入れないこと。
 *
 * ・「建設_解決策出し_1」のような「部門_種別_番号」のシートを全部読み、
 *   シートごとに JSON を1つ書き出す（例：KA_1.json）。
 * ・M列（公開）が ○ の行だけを書き出す。
 * ・問題が0件のシート（_2 など）も、空のJSONとして書き出す（アプリのメニューに (0) と出る）。
 * ・manifest.json も毎回作り直す。
 * ・書き出し先は、このスプレッドシートと同じフォルダの中の「問題Ⅰドリル_data」フォルダ。
 *   （RCCMの manifest.json を上書きしないよう、ドライブ全体は探さない）
 *
 * 使い方：メニュー「問題Ⅰドリル書き出し」→「JSONを書き出す（検証あり）」
 *        書き出された data フォルダの中身を、GitHub の data/ に上書きでアップロードする。
 */

// ===== 設定（通常は変えなくてよい） =====
var EXAM_NAME = '技術士 問題Ⅰドリル';
var SCHEMA_VERSION = '1.0';
var OUTPUT_FOLDER_NAME = '問題Ⅰドリル_data';
var OUTPUT_FOLDER_ID = '';   // 書き出し先を固定したいときだけ、フォルダIDを入れる

// 部門（並び順＝この順）。IDの1文字目
var DEPTS = [
  { name: '建設', code: 'K' },
  { name: '農業', code: 'N' }
];
// 種別（並び順＝この順）。IDの2文字目。「知識」を足すときは、ここに1行足してシートを作るだけ
var TYPES = [
  { name: '解決策出し',   code: 'A' },
  { name: '観点と課題',   code: 'B' },
  { name: '観点の構築',   code: 'C' },
  { name: '説明文の論理', code: 'D' },
  { name: '知識',         code: 'E' }
];
// 列の見出し（1行目）。列の位置ではなく見出しの文字で探す
var HEAD = {
  id: 'ID', dept: '部門', type: '種別', category: 'カテゴリー', source: '出典', question: '問い',
  c1: '選択肢1', c2: '選択肢2', c3: '選択肢3', c4: '選択肢4',
  answer: '正答', explanation: '解説', publish: '公開'
};
// =========================================

var SHEET_RE = /^(.+?)_(.+)_(\d+)$/;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('問題Ⅰドリル書き出し')
    .addItem('JSONを書き出す（検証あり）', 'exportDrill')
    .addItem('検証だけする', 'validateDrill')
    .addToUi();
}

function findDept_(name) { for (var i = 0; i < DEPTS.length; i++) if (DEPTS[i].name === name) return { d: DEPTS[i], order: i + 1 }; return null; }
function findType_(name) { for (var i = 0; i < TYPES.length; i++) if (TYPES[i].name === name) return { t: TYPES[i], order: i + 1 }; return null; }
function isPublished_(v) { var s = String(v).trim(); return s === '○' || s === '〇' || s === '◯'; }
function str_(v) { return v === null || v === undefined ? '' : String(v); }

/** 全シートを読み、パートと検証エラーを返す */
function buildAll_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var parts = [], errors = [], seenId = {};

  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    var m = name.match(SHEET_RE);
    if (!m) return;                                  // 「使い方」などは対象外
    var dept = findDept_(m[1]), type = findType_(m[2]), no = parseInt(m[3], 10);
    if (!dept) { errors.push(name + '：部門「' + m[1] + '」が設定（DEPTS）にありません'); return; }
    if (!type) { errors.push(name + '：種別「' + m[2] + '」が設定（TYPES）にありません'); return; }

    var prefix = dept.d.code + type.t.code;
    var values = sh.getDataRange().getValues();
    var header = (values[0] || []).map(function (h) { return String(h).trim(); });
    var idx = {};
    Object.keys(HEAD).forEach(function (k) { idx[k] = header.indexOf(HEAD[k]); });
    var missing = Object.keys(HEAD).filter(function (k) { return idx[k] < 0; }).map(function (k) { return HEAD[k]; });
    if (missing.length) { errors.push(name + '：見出しが足りません（' + missing.join('、') + '）'); return; }

    var qs = [];
    for (var r = 1; r < values.length; r++) {
      var row = values[r], rowNo = r + 1;
      var id = str_(row[idx.id]).trim();
      if (!id) continue;
      var at = name + ' ' + rowNo + '行（' + id + '）：';

      if (seenId[id]) errors.push(at + 'IDが重複しています（' + seenId[id] + ' と同じ）');
      seenId[id] = name + ' ' + rowNo + '行';
      if (id.indexOf(prefix) !== 0 || !/^[A-Z]{2}\d{3}$/.test(id))
        errors.push(at + 'IDは「' + prefix + '」＋3桁にしてください');

      if (!isPublished_(row[idx.publish])) continue;  // 公開○だけ書き出す（検証もここから）

      var ans = parseInt(row[idx.answer], 10);
      if (isNaN(ans) || ans < 1 || ans > 4) errors.push(at + '正答は1〜4（今：' + row[idx.answer] + '）');
      if (!str_(row[idx.category]).trim()) errors.push(at + 'カテゴリーが空です');
      if (!str_(row[idx.question]).trim()) errors.push(at + '問いが空です');
      ['c1', 'c2', 'c3', 'c4'].forEach(function (c, i) {
        if (!str_(row[idx[c]]).trim()) errors.push(at + '選択肢' + (i + 1) + 'が空です');
      });
      if (str_(row[idx.dept]).trim() && str_(row[idx.dept]).trim() !== dept.d.name)
        errors.push(at + '部門の列がシート名と違います');
      if (str_(row[idx.type]).trim() && str_(row[idx.type]).trim() !== type.t.name)
        errors.push(at + '種別の列がシート名と違います');

      qs.push({
        id: id,
        category: str_(row[idx.category]).trim(),
        source: str_(row[idx.source]).trim(),
        question: str_(row[idx.question]),
        choice_1: str_(row[idx.c1]), choice_2: str_(row[idx.c2]),
        choice_3: str_(row[idx.c3]), choice_4: str_(row[idx.c4]),
        answer: ans,
        explanation: str_(row[idx.explanation])
      });
    }
    // ID順（連問の順番を守る）
    qs.sort(function (a, b) { return a.id < b.id ? -1 : a.id > b.id ? 1 : 0; });

    parts.push({
      key: prefix + '_' + no,
      label: dept.d.name + ' ' + type.t.name + ' ' + no,
      file: prefix + '_' + no + '.json',
      sheet: name,
      dept: dept.d.name, dept_order: dept.order,
      type: type.t.name, type_order: type.order,
      sheet_no: no,
      count: qs.length,
      questions: qs
    });
  });

  parts.sort(function (a, b) {
    return (a.dept_order - b.dept_order) || (a.type_order - b.type_order) || (a.sheet_no - b.sheet_no);
  });
  return { parts: parts, errors: errors };
}

function validateDrill() {
  var res = buildAll_(), ui = SpreadsheetApp.getUi();
  if (res.errors.length) { ui.alert('検証エラー（' + res.errors.length + '件）\n\n' + res.errors.slice(0, 40).join('\n')); return; }
  ui.alert('検証OK\n\n' + summary_(res.parts));
}

function exportDrill() {
  var res = buildAll_(), ui = SpreadsheetApp.getUi();
  if (res.errors.length) { ui.alert('検証エラーのため中止（' + res.errors.length + '件）\n\n' + res.errors.slice(0, 40).join('\n')); return; }
  if (!res.parts.length) { ui.alert('「部門_種別_番号」の名前のシートが見つかりません。'); return; }

  var folder = outputFolder_();
  var now = isoJst_(new Date());
  var manifest = { schema_version: SCHEMA_VERSION, exam: EXAM_NAME, generated_at: now, parts: [] };

  res.parts.forEach(function (p) {
    var payload = {
      schema_version: SCHEMA_VERSION, exam: EXAM_NAME,
      part_key: p.key, part_label: p.label, sheet: p.sheet,
      dept: p.dept, type: p.type, sheet_no: p.sheet_no,
      generated_at: now, questions: p.questions
    };
    writeFile_(folder, p.file, JSON.stringify(payload, null, 2));
    manifest.parts.push({
      key: p.key, label: p.label, file: p.file, sheet: p.sheet,
      dept: p.dept, dept_order: p.dept_order, type: p.type, type_order: p.type_order,
      sheet_no: p.sheet_no, count: p.count, updated_at: now
    });
  });
  writeFile_(folder, 'manifest.json', JSON.stringify(manifest, null, 2));

  ui.alert('書き出し完了\n\n' + summary_(res.parts) +
    '\n\n書き出し先：' + folder.getName() + '\n' + folder.getUrl() +
    '\n\nこのフォルダの中身（manifest.json と各JSON）を、GitHub の data/ に上書きでアップロードしてください。');
}

function summary_(parts) {
  var total = 0;
  var lines = parts.map(function (p) { total += p.count; return p.label + '：' + p.count + '問'; });
  return lines.join('\n') + '\n合計：' + total + '問（公開○のみ）';
}

function outputFolder_() {
  if (OUTPUT_FOLDER_ID) return DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parents = file.getParents();
  var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var it = parent.getFoldersByName(OUTPUT_FOLDER_NAME);
  return it.hasNext() ? it.next() : parent.createFolder(OUTPUT_FOLDER_NAME);
}

/** 書き出し先フォルダの中だけを探して上書き（ドライブ全体は探さない） */
function writeFile_(folder, name, content) {
  var it = folder.getFilesByName(name);
  if (it.hasNext()) { var f = it.next(); f.setContent(content); return f; }
  return folder.createFile(name, content, MimeType.PLAIN_TEXT);
}

function isoJst_(d) { return Utilities.formatDate(d, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss") + '+09:00'; }
