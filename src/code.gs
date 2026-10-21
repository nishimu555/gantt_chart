/**
 * ガントチャート Web表示 ＆ Googleスライド出力
 * バージョン: 1.0.0
 *
 * 変更履歴
 * 1.0.0 (2026-10-21)
 *   - 初版
 *
 * シート構成
 *  ・「タスク」    ：分類 / 作業ID / 作業名 / 区分 / 開始日 / 終了日 / 作業明細 / 後続の作業ID / 担当 / 進捗率
 *  ・「スプリント」：名称 / 表示名 / 開始日 / 終了日 / メモ（任意。期間が1件もなければ期間行は表示しない）
 *                    ※スプリント以外の期間（凍結期間・休業期間など）も入力可。旧形式（スプリント名／ゴール）も読める
 *  ・「設定」      ：A列に項目名、B列に値
 *                    （プロジェクト名 / 基準日 / 会社名 / 機密表示 / 期間行の見出し / 進捗を表示 / 1ページ行数 / 出力フォルダID）
 *
 * 描画は index.html（ブラウザ側）で行い、スライドには1枚1画像として貼り付けます。
 */

const APP_VERSION = '1.0.0';
const SHEET_TASKS = 'タスク';
const SHEET_SPRINTS = 'スプリント';
const SHEET_SETTINGS = '設定';
const HEADERS = ['分類', '作業ID', '作業名', '区分', '開始日', '終了日', '作業明細', '後続の作業ID', '担当', '進捗率'];
const REQUIRED_HEADERS = ['分類', '作業ID', '作業名', '区分', '開始日'];
const SPRINT_HEADERS = ['名称', '表示名', '開始日', '終了日', 'メモ'];
const DEFAULT_CONFIDENTIAL = 'CONFIDENTIAL｜本資料は関係者限りでお取り扱いください';
const SLIDE_FONT = 'Noto Sans JP';

/* ───────── 画面 ───────── */

/** Web app の入口 */
function doGet() {
  const t = HtmlService.createTemplateFromFile('index');
  t.context = 'web';
  return t.evaluate()
    .setTitle('ガントチャート')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** スプレッドシートを開いたときにメニューを追加 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ガントチャート')
    .addItem('表示・スライド出力', 'showDialog')
    .addSeparator()
    .addItem('入力シートを作成（サンプル付き）', 'setupSheetsWithSample')
    .addItem('入力シートを作成（空）', 'setupSheetsEmpty')
    .addToUi();
}

/** スプレッドシート上でダイアログとして開く（できるだけ大きく） */
function showDialog() {
  const t = HtmlService.createTemplateFromFile('index');
  t.context = 'dialog';
  const html = t.evaluate()
    .setWidth(1600)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'ガントチャート');
}

/* ───────── データ読込 ───────── */

/**
 * 対象スプレッドシート。
 * スクリプトプロパティ SPREADSHEET_ID があればそれを、なければこのスクリプトが紐づくシートを使う。
 */
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

/** 画面から呼ぶ：設定・タスク・スプリント・警告を返す（日付は yyyy-MM-dd の文字列） */
function getGanttData() {
  const ss = getSpreadsheet_();
  const tz = ss.getSpreadsheetTimeZone();
  const settings = readSettings_(ss, tz);
  const warnings = [];
  const sprints = readPeriods_(ss, tz, warnings);

  const sh = ss.getSheetByName(SHEET_TASKS);
  if (!sh) {
    throw new Error('「' + SHEET_TASKS + '」シートがありません。メニュー「ガントチャート ＞ 入力シートを作成」を実行してください。');
  }
  const range = sh.getDataRange();
  const values = range.getValues();
  const display = range.getDisplayValues();
  if (values.length < 2) {
    warnings.push('タスクが入力されていません。');
    return { version: APP_VERSION, settings: settings, tasks: [], sprints: sprints, warnings: warnings };
  }

  const head = values[0].map(function (v) { return String(v).trim(); });
  const col = {};
  HEADERS.forEach(function (h) { col[h] = head.indexOf(h); });
  REQUIRED_HEADERS.forEach(function (h) {
    if (col[h] < 0) throw new Error('「' + SHEET_TASKS + '」シートの1行目に見出し「' + h + '」が見つかりません。');
  });

  const tasks = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const disp = display[r];
    const get = function (h) { return col[h] >= 0 ? row[col[h]] : ''; };
    const getDisp = function (h) { return col[h] >= 0 ? String(disp[col[h]]).trim() : ''; };

    const id = String(get('作業ID')).trim();
    const name = String(get('作業名')).trim();
    if (!id && !name) continue; // 空行
    const rowNo = r + 1;
    if (!id) { warnings.push(rowNo + '行目：作業IDがないためスキップしました（' + name + '）'); continue; }

    const isMs = /マイル|milestone|^ms$/i.test(String(get('区分')).trim());
    let s = toDateStr_(get('開始日'), tz);
    let e = toDateStr_(get('終了日'), tz);
    if (isMs) { s = e || s; e = s; } // マイルストーンは終了日（なければ開始日）の1点
    if (!s) { warnings.push(rowNo + '行目：' + id + ' の日付がないためスキップしました'); continue; }
    if (!e) e = s;
    if (e < s) { warnings.push(rowNo + '行目：' + id + ' の終了日が開始日より前のため、開始日に合わせました'); e = s; }

    tasks.push({
      group: String(get('分類')).trim() || '（未分類）',
      id: id,
      name: name,
      ms: isMs,
      s: s,
      e: e,
      detail: String(get('作業明細') || '').trim(),
      next: String(get('後続の作業ID') || '')
        .split(/[,、，\s]+/)
        .map(function (x) { return x.trim(); })
        .filter(Boolean),
      who: String(get('担当') || '').trim(),
      prog: toProgress_(get('進捗率'), getDisp('進捗率'))
    });
  }

  // ID の重複・後続IDの存在チェック
  const ids = {};
  tasks.forEach(function (t) {
    if (ids[t.id]) warnings.push('作業IDが重複しています：' + t.id);
    ids[t.id] = true;
  });
  tasks.forEach(function (t) {
    t.next = t.next.filter(function (n) {
      if (!ids[n]) { warnings.push(t.id + ' の後続の作業ID「' + n + '」が見つかりません'); return false; }
      if (n === t.id) { warnings.push(t.id + ' の後続に自分自身が指定されています'); return false; }
      return true;
    });
  });

  return { version: APP_VERSION, settings: settings, tasks: tasks, sprints: sprints, warnings: warnings };
}

/** 設定シートを読む（A列：項目名、B列：値） */
function readSettings_(ss, tz) {
  const out = {
    project: ss.getName(),
    today: '',   // 基準日が空（または設定シートがない）なら本日線を引かない
    company: '',
    confidential: DEFAULT_CONFIDENTIAL,
    periodTitle: 'スプリント',
    rowsPerPage: 14,
    showProgress: true,
    folderId: ''
  };
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) return out;
  sh.getDataRange().getValues().forEach(function (r) {
    const key = String(r[0]).trim();
    const v = r[1];
    if (key === '機密表示') { out.confidential = String(v === null ? '' : v).trim(); return; } // 空欄なら表示しない
    if (v === '' || v === null) return;
    if (key === 'プロジェクト名') out.project = String(v).trim();
    else if (key === '基準日') { const d = toDateStr_(v, tz); if (d) out.today = d; }
    else if (key === '会社名') out.company = String(v).trim();
    else if (key === '進捗を表示') out.showProgress = !/^(しない|非表示|なし|no|off|false|0)$/i.test(String(v).trim());
    else if (key === '1ページ行数') { const n = parseInt(v, 10); if (n >= 5) out.rowsPerPage = Math.min(n, 18); }
    else if (key === '出力フォルダID') out.folderId = String(v).trim();
    else if (key === '期間行の見出し') out.periodTitle = String(v).trim();
  });
  return out;
}

/**
 * 期間シート（スプリントなど）を読む。シートがない・1件もない場合は空配列（＝期間行は表示しない）
 * 見出しは新形式（名称／表示名／開始日／終了日／メモ）と旧形式（スプリント名／ゴール）のどちらでも可
 */
function readPeriods_(ss, tz, warnings) {
  const sh = ss.getSheetByName(SHEET_SPRINTS);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0].map(function (v) { return String(v).trim(); });
  const find = function (names) {
    for (let i = 0; i < names.length; i++) { const k = head.indexOf(names[i]); if (k >= 0) return k; }
    return -1;
  };
  const cName = find(['名称', 'スプリント名']), cLabel = find(['表示名']);
  const cS = find(['開始日']), cE = find(['終了日']), cMemo = find(['メモ', 'ゴール']);
  if (cS < 0 && cE < 0) { warnings.push(SHEET_SPRINTS + 'シートに「開始日」「終了日」の見出しがありません'); return []; }
  const cell = function (row, c) { return c >= 0 ? row[c] : ''; };

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = String(cell(row, cName) || '').trim();
    const label = String(cell(row, cLabel) || '').trim();
    let s = toDateStr_(cell(row, cS), tz);
    let e = toDateStr_(cell(row, cE), tz);
    if (!s && !e) {
      if (name || label) warnings.push(SHEET_SPRINTS + 'シート ' + (r + 1) + '行目：日付がないためスキップしました（' + (name || label) + '）');
      continue;
    }
    if (!s) s = e;   // 片方だけなら1日の期間
    if (!e) e = s;
    if (e < s) { warnings.push(SHEET_SPRINTS + 'シート ' + (r + 1) + '行目：終了日が開始日より前のためスキップしました'); continue; }
    const fullName = name || label || '期間' + (out.length + 1);
    out.push({
      name: fullName,
      label: label || fullName.replace(/^sprint\s*/i, 'S'),
      s: s,
      e: e,
      goal: String(cell(row, cMemo) || '').trim()
    });
  }
  out.sort(function (x, y) { return x.s < y.s ? -1 : x.s > y.s ? 1 : 0; });
  return out;
}

/** Date／文字列 → 'yyyy-MM-dd'（読めなければ空文字） */
function toDateStr_(v, tz) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  const m = String(v).trim().match(/^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

/** 進捗率 → 0〜100 の数値（未入力は null）。「60%」表示のセルも、60 と入力したセルも可 */
function toProgress_(value, display) {
  if (value === '' || value === null || value === undefined) return null;
  let n = Number(value);
  if (isNaN(n)) n = parseFloat(String(value).replace('%', ''));
  if (isNaN(n)) return null;
  if (/%$/.test(display) && n <= 1) n = n * 100;
  return Math.max(0, Math.min(100, Math.round(n)));
}

/* ───────── スライド出力（画面から1枚ずつ呼ばれる） ───────── */

/** 新しいプレゼンテーションを作り、IDを返す */
function startExport(title) {
  return SlidesApp.create(title).getId();
}

/**
 * 1枚分を追加する
 *  ┌ タイトル（左）                         基準日（右寄せ）┐
 *  │ ガントチャート画像（できるだけ大きく）               │
 *  └ 機密表示（左）                             会社名（右）┘
 * @param {string} presId プレゼンテーションID
 * @param {string} b64 PNG画像（Base64）
 * @param {Object} page {title, baseDate, confidential, company, notes, imgW, imgH}
 */
function addSlidePage(presId, b64, page) {
  const pres = SlidesApp.openById(presId);
  const W = pres.getPageWidth();
  const H = pres.getPageHeight();
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  const m = W * 0.022;      // 左右の余白
  const top = H * 0.025;
  const headH = 26;
  const footH = 13;

  addText_(slide, page.title, m, top, W * 0.66, headH, 18, true, '#0F172A');
  addText_(slide, page.baseDate, W - m - W * 0.3, top, W * 0.3, headH, 11, false, '#334155',
    SlidesApp.ParagraphAlignment.END);

  // 画像：タイトルと注記の間いっぱいに、縦横比を保って配置
  const areaTop = top + headH + 2;
  const areaBottom = H - footH - 4;
  const areaW = W - m * 2;
  const areaH = areaBottom - areaTop;
  let imgW = areaW, imgH = imgW * page.imgH / page.imgW;
  if (imgH > areaH) { imgH = areaH; imgW = imgH * page.imgW / page.imgH; }
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', 'gantt.png');
  slide.insertImage(blob, (W - imgW) / 2, areaTop, imgW, imgH);

  const footTop = H - footH - 3;
  addText_(slide, page.confidential, m, footTop, W * 0.6, footH, 8, true, '#B91C1C');
  addText_(slide, page.company, W - m - W * 0.35, footTop, W * 0.35, footH, 8, false, '#64748B',
    SlidesApp.ParagraphAlignment.END);

  if (page.notes) {
    slide.getNotesPage().getSpeakerNotesShape().getText().setText(page.notes);
  }
}

/** 仕上げ：最初の空スライドを削除し、フォルダへ移動してURLを返す */
function finishExport(presId) {
  const pres = SlidesApp.openById(presId);
  const slides = pres.getSlides();
  if (slides.length > 1 && slides[0].getImages().length === 0) slides[0].remove();
  pres.saveAndClose();

  // 出力先：設定の「出力フォルダID」→ なければスプレッドシートと同じフォルダ
  try {
    const ss = getSpreadsheet_();
    const settings = readSettings_(ss, ss.getSpreadsheetTimeZone());
    const folder = settings.folderId
      ? DriveApp.getFolderById(settings.folderId)
      : DriveApp.getFileById(ss.getId()).getParents().next();
    DriveApp.getFileById(presId).moveTo(folder);
  } catch (err) {
    console.warn('フォルダ移動をスキップしました: ' + err);
  }
  return 'https://docs.google.com/presentation/d/' + presId + '/edit';
}

function addText_(slide, text, left, top, width, height, size, bold, color, align) {
  if (!text) return null;
  const shape = slide.insertTextBox(text, left, top, width, height);
  const tr = shape.getText();
  tr.getTextStyle()
    .setFontFamily(SLIDE_FONT)
    .setFontSize(size)
    .setBold(!!bold)
    .setForegroundColor(color);
  if (align) tr.getParagraphStyle().setParagraphAlignment(align);
  shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  return shape;
}

/* ───────── 入力シートの作成 ───────── */

function setupSheetsWithSample() { setupSheets_(true); }
function setupSheetsEmpty() { setupSheets_(false); }

/**
 * 「タスク」「スプリント」「設定」シートを作る（すでにあるシートは触らない）
 * @param {boolean} withSample サンプルデータを入れるか
 */
function setupSheets_(withSample) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const d = function (y, m, day) { return new Date(y, m - 1, day); };
  const made = [];

  // タスク
  if (!ss.getSheetByName(SHEET_TASKS)) {
    const sh = ss.insertSheet(SHEET_TASKS);
    styleHeader_(sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]));
    if (withSample) {
      const rows = [
        ['要件定義', 'A-01', '現状業務ヒアリング', 'ガント', d(2026, 9, 28), d(2026, 10, 9), '各部門へのヒアリング（計6回）', 'A-02', '共同', 100],
        ['要件定義', 'A-02', '要件定義書作成', 'ガント', d(2026, 10, 12), d(2026, 10, 23), '業務要件・機能要件の整理', 'A-03', '当社', 60],
        ['要件定義', 'A-03', '要件レビュー', 'ガント', d(2026, 10, 26), d(2026, 10, 30), '要件定義書のご確認', 'M-01', 'お客様', 0],
        ['要件定義', 'M-01', '要件確定', 'マイルストーン', '', d(2026, 10, 30), '', 'B-01', 'お客様', 0],
        ['設計', 'B-01', '基本設計', 'ガント', d(2026, 11, 2), d(2026, 11, 13), 'システム構成・機能一覧', 'B-02, B-03', '当社', 0],
        ['設計', 'B-02', '画面設計', 'ガント', d(2026, 11, 16), d(2026, 11, 27), '画面レイアウト・画面遷移', 'B-04', '当社', 0],
        ['設計', 'B-03', '外部連携設計', 'ガント', d(2026, 11, 16), d(2026, 11, 27), '外部システムとのインターフェース定義', 'B-04', '当社', 0],
        ['設計', 'B-04', '詳細設計', 'ガント', d(2026, 11, 30), d(2026, 12, 11), '', 'M-02', '当社', 0],
        ['設計', 'M-02', '設計承認', 'マイルストーン', '', d(2026, 12, 11), '', 'C-01', 'お客様', 0],
        ['開発', 'C-01', '開発環境構築', 'ガント', d(2026, 12, 14), d(2026, 12, 18), '', 'C-02, C-03', '当社', 0],
        ['開発', 'C-02', '画面開発', 'ガント', d(2026, 12, 21), d(2027, 1, 22), '', 'C-04', '当社', 0],
        ['開発', 'C-03', '外部連携開発', 'ガント', d(2026, 12, 21), d(2027, 1, 15), '', 'C-04', '当社', 0],
        ['開発', 'C-04', '単体テスト', 'ガント', d(2027, 1, 25), d(2027, 2, 5), '', 'D-01', '当社', 0],
        ['テスト', 'D-01', '結合テスト', 'ガント', d(2027, 2, 8), d(2027, 2, 19), '', 'D-02', '当社', 0],
        ['テスト', 'D-02', '受入テスト', 'ガント', d(2027, 2, 22), d(2027, 3, 5), 'お客様による業務シナリオ確認', 'M-03', 'お客様', 0],
        ['テスト', 'M-03', '受入完了', 'マイルストーン', '', d(2027, 3, 5), '', 'E-01', 'お客様', 0],
        ['移行・リリース', 'E-01', '本番移行リハーサル', 'ガント', d(2027, 3, 8), d(2027, 3, 12), '', 'E-02', '共同', 0],
        ['移行・リリース', 'E-02', '本番移行', 'ガント', d(2027, 3, 15), d(2027, 3, 17), '', 'M-04', '共同', 0],
        ['移行・リリース', 'M-04', '本番稼働', 'マイルストーン', '', d(2027, 3, 19), '', '', '共同', 0]
      ];
      sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
    }
    sh.getRange(2, 5, 500, 2).setNumberFormat('yyyy/mm/dd');
    sh.getRange(2, 4, 500, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['ガント', 'マイルストーン'], true).build());
    sh.getRange(1, 8).setNote('複数ある場合はカンマ区切り（例：B-02, B-03）');
    sh.getRange(1, 10).setNote('0〜100 で入力（例：60）。「60%」表示のセルでも可');
    sh.setFrozenRows(1);
    [110, 70, 200, 100, 95, 95, 260, 120, 70, 60].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    made.push(SHEET_TASKS);
  }

  // スプリント
  if (!ss.getSheetByName(SHEET_SPRINTS)) {
    const sp = ss.insertSheet(SHEET_SPRINTS);
    styleHeader_(sp.getRange(1, 1, 1, SPRINT_HEADERS.length).setValues([SPRINT_HEADERS]));
    if (withSample) {
      const goals = { 1: '現状業務の把握', 2: '要件定義書ドラフト', 3: '要件確定', 5: '基本設計完了' };
      const rows = [];
      for (let i = 0; i < 13; i++) { // 2週間スプリント × 13
        rows.push(['Sprint ' + (i + 1), 'S' + (i + 1), d(2026, 9, 28 + i * 14), d(2026, 9, 28 + i * 14 + 13), goals[i + 1] || '']);
      }
      rows.push(['年末年始休業', '休業', d(2026, 12, 29), d(2027, 1, 3), '期間行にはスプリント以外の期間も入れられます']);
      sp.getRange(2, 1, rows.length, SPRINT_HEADERS.length).setValues(rows);
    }
    sp.getRange(2, 3, 200, 2).setNumberFormat('yyyy/mm/dd');
    sp.getRange(1, 2).setNote('ガントチャートの期間行に表示する短い名前（自由入力）。空欄なら名称から自動で作ります');
    sp.getRange(1, 3).setNote('開始日・終了日の片方だけでも可（1日の期間として表示）');
    sp.getRange(1, 5).setNote('任意。スライドのスピーカーノートに出力されます');
    sp.setFrozenRows(1);
    [140, 80, 95, 95, 320].forEach(function (w, i) { sp.setColumnWidth(i + 1, w); });
    made.push(SHEET_SPRINTS);
  }

  // 設定
  if (!ss.getSheetByName(SHEET_SETTINGS)) {
    const st = ss.insertSheet(SHEET_SETTINGS);
    st.getRange(1, 1, 9, 3).setValues([
      ['項目', '値', '説明'],
      ['プロジェクト名', withSample ? '[プロジェクト名]' : '', 'スライドのタイトルに表示（空欄ならファイル名）'],
      ['基準日', new Date(), '本日線を引く日付（空欄なら本日線を引かない）'],
      ['会社名', withSample ? '[会社名]' : '', 'スライド右下に表示'],
      ['機密表示', DEFAULT_CONFIDENTIAL, 'スライド左下に表示（空欄なら表示しない）'],
      ['期間行の見出し', 'スプリント', '期間行の左に表示する見出し（例：スプリント、フェーズ）'],
      ['進捗を表示', 'する', '「しない」にすると進捗列を消し、バーをすべて塗りつぶして表示'],
      ['1ページ行数', 14, 'スライド1枚の最大行数（分類見出し含む、5〜18）'],
      ['出力フォルダID', '', '空欄ならこのスプレッドシートと同じフォルダ']
    ]);
    styleHeader_(st.getRange(1, 1, 1, 3));
    st.getRange(3, 2).setNumberFormat('yyyy/mm/dd');
    st.getRange(7, 2).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['する', 'しない'], true).build());
    st.setColumnWidth(1, 130); st.setColumnWidth(2, 320); st.setColumnWidth(3, 340);
    made.push(SHEET_SETTINGS);
  }

  ui.alert(made.length
    ? '「' + made.join('」「') + '」シートを作成しました。'
    : '入力シートはすべて作成済みです。作り直す場合は、既存のシートを削除またはリネームしてから実行してください。');
}

function styleHeader_(range) {
  range.setFontWeight('bold').setBackground('#E2E8F0');
}