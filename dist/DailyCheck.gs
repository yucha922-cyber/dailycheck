/**
 * ============================================================
 *  サロン日次業務チェック  ―  一括貼り付け用ファイル
 *  version 2.0.1
 * ------------------------------------------------------------
 *  このファイル1つを Apps Script エディタの「コード.gs」に
 *  そのまま貼り付ければ導入できます。
 *  （apps-script/ の各ファイルを結合した自動生成ファイルです。
 *    修正は apps-script/ 側で行い、node tools/build.js で再生成してください）
 * ============================================================
 */
// ============================================================
// Config.gs
// ============================================================

/**
 * ============================================================
 *  Config.gs  ―  シート名の解決ルールと既定値
 * ------------------------------------------------------------
 *  ★重要★
 *  店舗名・営業時間などの「運用の設定」はコードではなく
 *  スプレッドシートの「設定」タブで編集します（Settings.gs）。
 *  このファイルでは
 *    ・どのタブを使うか（別名の吸収）
 *    ・列見出しの別名（既存シートの列をそのまま使うため）
 *    ・初期データ
 *  だけを定義します。
 * ============================================================
 */

const APP_VERSION = '2.0.1';

/**
 * 使用するタブ名。
 *  canonical … 無かったときに新規作成する名前
 *  aliases   … 既にこの名前のタブがあれば「それを使う」（新規作成しない）
 * ------------------------------------------------------------
 * これにより、既存の「各業務マスター」「今日のチェック」「履歴」「設定」
 * をそのまま生かして動きます（別タブを増やしません）。
 */
const SHEET_DEFS = {
  MASTER: {
    canonical: '各業務マスター',
    aliases: ['各業務マスター', '業務マスター', '業務マスタ', '各業務マスタ', 'マスター', 'マスタ'],
  },
  CHECK: {
    canonical: '今日のチェック',
    aliases: ['今日のチェック', '本日のチェック', 'チェック', 'デイリーチェック', '今日のタスク'],
  },
  LOG: {
    canonical: '履歴',
    aliases: ['履歴', '日次ログ', 'ログ', '記録', '実績'],
  },
  SETTINGS: {
    canonical: '設定',
    aliases: ['設定', '設定シート', 'セッティング', 'Settings', 'config'],
  },
  DASH: {
    canonical: 'ダッシュボード',
    aliases: ['ダッシュボード', 'Dashboard', '集計'],
  },
};

/**
 * 「各業務マスター」の列。見出し文字で列を探します。
 *  ・既存の見出しがあれば その列をそのまま使う
 *  ・見つからない列だけ、右端に追加する（既存データは壊しません）
 */
const MASTER_FIELDS = [
  { key: 'id',       header: 'ID',       aliases: ['id', '業務id', 'タスクid', 'no', 'no.', '番号', '#'] },
  { key: 'phase',    header: '区分',     aliases: ['区分', 'カテゴリ', 'カテゴリー', 'タイミング', '分類', 'フェーズ', '時間帯'] },
  { key: 'name',     header: '業務名',   aliases: ['業務名', 'タスク名', '業務', '内容', '項目', 'チェック項目', '作業内容', 'タスク'] },
  { key: 'min',      header: '目安(分)', aliases: ['目安(分)', '目安（分）', '目安時間', '目安分', '目安', '所要時間', '所要(分)', '所要', '時間(分)', '分', '想定時間', '作業時間'] },
  { key: 'freq',     header: '頻度',     aliases: ['頻度', '実施頻度', 'サイクル', '周期'] },
  { key: 'required', header: '必須',     aliases: ['必須', '必須?', '重要', '必須フラグ'] },
  { key: 'memo',     header: '備考',     aliases: ['備考', 'メモ', '補足', '注意点', '説明'] },
  { key: 'enabled',  header: '有効',     aliases: ['有効', '有効?', '使用', '利用', 'on', '稼働'] },
];

/**
 * 「履歴」の列。こちらも見出しで解決し、足りない列だけ追加します。
 */
const LOG_FIELDS = [
  { key: 'date',    header: '日付',     aliases: ['日付', '年月日', 'date'] },
  { key: 'storeId', header: '店舗ID',   aliases: ['店舗id', 'storeid', '店id'] },
  { key: 'store',   header: '店舗名',   aliases: ['店舗名', '店舗', 'store'] },
  { key: 'taskId',  header: '業務ID',   aliases: ['業務id', 'タスクid', 'taskid'] },
  { key: 'name',    header: '業務名',   aliases: ['業務名', 'タスク名', '内容', '項目'] },
  { key: 'phase',   header: '区分',     aliases: ['区分', 'カテゴリ', 'カテゴリー', '分類'] },
  { key: 'done',    header: '完了',     aliases: ['完了', '完了?', 'done', '実施'] },
  { key: 'time',    header: '完了時刻', aliases: ['完了時刻', '時刻', '完了時間'] },
  { key: 'staff',   header: '担当',     aliases: ['担当', '担当者', 'スタッフ'] },
  { key: 'min',     header: '目安(分)', aliases: ['目安(分)', '目安（分）', '目安分', '目安'] },
  { key: 'memo',    header: 'メモ',     aliases: ['メモ', '備考'] },
  { key: 'updated', header: '更新時刻', aliases: ['更新時刻', '更新日時', '更新'] },
];

/**
 * 「設定」タブの項目。
 *  label … 設定タブに表示される項目名（この文字で値を探します）
 *  def   … 既定値
 */
const SETTING_DEFS = [
  { key: 'STORE_ID',      label: '店舗ID',                   def: 'STORE_001', hint: '本部で一意になるID（例 STORE_001）' },
  { key: 'STORE_NAME',    label: '店舗名',                   def: '本店',      hint: '表示用の店舗名' },
  { key: 'OPEN_TIME',     label: '開店時刻',                 def: '11:00',     hint: '例 11:00' },
  { key: 'CLOSE_TIME',    label: '閉店時刻',                 def: '20:00',     hint: '例 20:00' },
  { key: 'BUFFER_MIN',    label: '目標作業時間(分)',          def: 10,          hint: '開店前・閉店前それぞれの目標（この分数以内で終える）' },
  { key: 'ALERT_EMAIL',   label: '未完了アラート送信先',       def: '',          hint: '空欄なら送信しません（例 you@example.com）' },
  { key: 'CENTRAL_SS_ID', label: '本部集約スプレッドシートID', def: '',          hint: '空欄なら送信しません。URLの /d/◯◯/edit の◯◯' },
  { key: 'MORNING_HOUR',  label: '朝の自動リセット時刻(時)',   def: 6,           hint: '0〜23。前日を確定して当日分を用意する時刻' },
  { key: 'NIGHT_HOUR',    label: '夜の自動確定時刻(時)',       def: 20,          hint: '0〜23。その日を締めて履歴に保存する時刻' },
  { key: 'USE_DASHBOARD', label: 'ダッシュボードを使う',       def: false,       hint: 'TRUE にすると「ダッシュボード」タブを作成・更新します' },
];

/** 区分の並び順・色 */
const PHASE_ORDER = ['開店前', '営業中', '閉店前'];
const PHASE_COLOR = {
  '開店前': '#e8f0fe',
  '営業中': '#e6f4ea',
  '閉店前': '#fce8e6',
};
const DEFAULT_PHASE = '営業中';

/** 頻度の入力候補（各業務マスターのプルダウン用。自由入力も可） */
const FREQ_OPTIONS = [
  '毎日', '毎回', '平日', '土日',
  '月', '火', '水', '木', '金', '土', '日',
  '月末',
];

/**
 * 「各業務マスター」が空だったときだけ入れる初期データ。
 *  [区分, 業務名, 目安分, 頻度, 必須, 備考]
 *  ※既にデータがある場合は一切上書きしません。
 */
const DEFAULT_TASKS = [
  ['開店前', '清掃',              10, '毎日', true,  ''],
  ['開店前', 'カルテ準備',          5, '毎日', true,  ''],
  ['営業中', 'カルテ記入',          0, '毎回', true,  '施術ごとに随時'],
  ['営業中', '39メール',            3, '毎回', true,  '新規来院者へ'],
  ['営業中', '口コミ返信',          5, '毎日', false, ''],
  ['営業中', 'SATTOU入力',          3, '毎日', true,  '予約者の来院有無を入力'],
  ['営業中', '前日リマインド',       5, '毎日', true,  '翌日予約者へ'],
  ['営業中', 'ブログ作成',          10, '毎日', false, 'サロンボード'],
  ['閉店前', '売上確認',            5, '毎日', true,  'Square × スプレッドシート照合'],
  ['閉店前', '現金確認',            3, '毎日', true,  ''],
  ['閉店前', '事業計画更新',         5, '毎日', true,  '売上・会員管理'],
  ['閉店前', 'カルテスクリプト',      5, '毎日', false, ''],
  ['閉店前', '週次棚卸し',          10, '月',   false, '毎週月曜だけ実施の例'],
];

// ============================================================
// Common.gs
// ============================================================

/**
 * ============================================================
 *  Common.gs  ―  共通ユーティリティ
 * ------------------------------------------------------------
 *  ここでの一番の役割は「既存のタブ・既存の列を壊さずに使う」こと。
 *   ・タブ …… 別名リストで探し、見つかればそれを使う（新規作成しない）
 *   ・列 …… 見出し文字で探し、無い列だけ右端に追加する
 * ============================================================
 */

const WEEKDAY_JP = ['日', '月', '火', '水', '木', '金', '土'];
const TZ = 'Asia/Tokyo';

/** スプレッドシート本体 */
function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/* ============================================================
 *  タブの解決
 * ========================================================== */

/** 比較用に正規化（全角/半角・空白・大文字小文字・記号ゆれを吸収） */
function norm_(s) {
  return String(s == null ? '' : s)
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[（）]/g, function (c) { return c === '（' ? '(' : ')'; })
    .replace(/[\s　_\-・]/g, '')
    .toLowerCase()
    .trim();
}

/**
 * 定義済みのシートを取得。
 * @param {string} defKey  SHEET_DEFS のキー（MASTER / CHECK / LOG / SETTINGS / DASH）
 * @param {boolean} create 見つからないとき作成するか
 * @return {Sheet|null}
 */
function sheetOf_(defKey, create) {
  const def = SHEET_DEFS[defKey];
  const ss = ss_();
  const sheets = ss.getSheets();

  // 1) 別名の順に、既存タブを探す（＝既にあるタブを最優先で使う）
  for (let a = 0; a < def.aliases.length; a++) {
    const target = norm_(def.aliases[a]);
    for (let i = 0; i < sheets.length; i++) {
      if (norm_(sheets[i].getName()) === target) return sheets[i];
    }
  }
  // 2) 部分一致（「今日のチェック（新）」のような名前も拾う）
  for (let a = 0; a < def.aliases.length; a++) {
    const target = norm_(def.aliases[a]);
    if (target.length < 3) continue;
    for (let i = 0; i < sheets.length; i++) {
      if (norm_(sheets[i].getName()).indexOf(target) >= 0) return sheets[i];
    }
  }
  if (!create) return null;
  return ss.insertSheet(def.canonical);
}

/** よく使うシートの取得ショートカット */
function masterSheet_(create)   { return sheetOf_('MASTER',   create !== false); }
function checkSheet_(create)    { return sheetOf_('CHECK',    create !== false); }
function logSheet_(create)      { return sheetOf_('LOG',      create !== false); }
function settingsSheet_(create) { return sheetOf_('SETTINGS', create !== false); }

/* ============================================================
 *  列の解決（見出しベース）
 * ========================================================== */

/**
 * 見出し行から、フィールド定義に対応する列番号を求める。
 * 見つからないフィールドは（createMissing=true なら）右端に列を追加。
 *
 * @param {Sheet} sh
 * @param {Array} fields   MASTER_FIELDS / LOG_FIELDS
 * @param {number} headerRow 見出しの行番号（既定1）
 * @param {boolean} createMissing 足りない列を追加するか
 * @return {Object} { key: 列番号(1始まり), ... }  ＋ _headerRow, _width
 */
function resolveColumns_(sh, fields, headerRow, createMissing) {
  headerRow = headerRow || 1;
  createMissing = createMissing !== false;

  let width = Math.max(sh.getLastColumn(), 1);
  let headers = sh.getRange(headerRow, 1, 1, width).getValues()[0];

  const map = { _headerRow: headerRow };
  const used = {};

  fields.forEach(function (f) {
    const cands = [f.header].concat(f.aliases || []).map(norm_);
    for (let c = 0; c < headers.length; c++) {
      const h = norm_(headers[c]);
      if (!h || used[c + 1]) continue;
      if (cands.indexOf(h) >= 0) {
        map[f.key] = c + 1;
        used[c + 1] = true;
        return;
      }
    }
  });

  // 完全一致で見つからなかったものは部分一致で再挑戦
  fields.forEach(function (f) {
    if (map[f.key]) return;
    const cands = [f.header].concat(f.aliases || []).map(norm_).filter(function (x) { return x.length >= 2; });
    for (let c = 0; c < headers.length; c++) {
      const h = norm_(headers[c]);
      if (!h || used[c + 1]) continue;
      for (let k = 0; k < cands.length; k++) {
        if (h === cands[k] || h.indexOf(cands[k]) >= 0 || cands[k].indexOf(h) >= 0) {
          map[f.key] = c + 1;
          used[c + 1] = true;
          return;
        }
      }
    }
  });

  // それでも無い列は右端に追加
  if (createMissing) {
    let next = width;
    // 末尾の空見出しを詰める
    while (next > 0 && !String(headers[next - 1] || '').trim()) next--;
    fields.forEach(function (f) {
      if (map[f.key]) return;
      next++;
      if (next > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), next - sh.getMaxColumns());
      sh.getRange(headerRow, next).setValue(f.header);
      map[f.key] = next;
      used[next] = true;
    });
    width = Math.max(width, next);
  }

  map._width = width;
  return map;
}

/* ============================================================
 *  日付・時刻
 * ========================================================== */

/** 日付セルかどうか（Sheetsが返す Date を判定） */
function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

function today_()          { return new Date(); }
function todayStr_()       { return dateStr_(new Date()); }
function dateStr_(d)       { return Utilities.formatDate(d, TZ, 'yyyy/MM/dd'); }
function nowTimeStr_()     { return Utilities.formatDate(new Date(), TZ, 'HH:mm'); }
function weekdayChar_(d)   { return WEEKDAY_JP[d.getDay()]; }
function dateLabel_(d)     { return dateStr_(d) + '(' + weekdayChar_(d) + ')'; }
function todayLabel_()     { return dateLabel_(new Date()); }

/** 'yyyy/MM/dd' 文字列 → Date（不正なら今日） */
function parseDate_(s) {
  if (isDate_(s)) return s;
  const m = String(s || '').match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return new Date();
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** セルの値を 'HH:mm' 表記に（Date でも文字列でもOK） */
function timeStr_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, TZ, 'HH:mm');
  return String(v == null ? '' : v).trim();
}

/**
 * 目安時間を「分」の数値にする。
 *  5        → 5
 *  '5分'     → 5
 *  '10 min' → 10
 *  '1時間30分' → 90
 *  '約5分程度' → 5
 *  時刻セル（0:05）→ 5
 *  空欄・読めない値 → 0
 */
function toMinutes_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  if (isDate_(v)) return v.getHours() * 60 + v.getMinutes();

  let s = String(v)
    .replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[．，,]/g, function (c) { return c === '．' ? '.' : ''; })
    .replace(/\s/g, '');

  let total = 0, hit = false;
  const h = s.match(/(\d+(?:\.\d+)?)(?:時間|hours?|hrs?|h)/i);
  if (h) { total += parseFloat(h[1]) * 60; hit = true; s = s.replace(h[0], ''); }
  const m = s.match(/(\d+(?:\.\d+)?)(?:分|minutes?|mins?|m)/i);
  if (m) { total += parseFloat(m[1]); hit = true; }
  if (hit) return Math.round(total);

  const n = parseFloat(s.replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

/** その日が月末か */
function isMonthEnd_(d) {
  const n = new Date(d.getFullYear(), d.getMonth() + 1, 0);
  return d.getDate() === n.getDate();
}

/** true/false のゆるい判定（チェックボックス・文字列どちらでも） */
function toBool_(v) {
  if (v === true) return true;
  if (v === false) return false;
  const s = String(v == null ? '' : v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === '○' || s === '◯' || s === '✓' ||
         s === 'yes' || s === 'y' || s === 'はい' || s === '必須' || s === '有効';
}

/* ============================================================
 *  頻度の判定
 * ========================================================== */

/**
 * その業務を「その日」に出すか判定する。
 *
 *  毎日 / 毎回 / 日次      → 毎日
 *  （空欄）                → 毎日（追加直後に消えないように）
 *  平日                    → 月〜金
 *  土日 / 週末             → 土・日
 *  月 火 水 木 金 土 日     → その曜日（「月,木」「月水金」「毎週火」も可）
 *  月末 / 末日             → 月末日
 *  毎月15日 / 15日         → 毎月その日（複数可「1日,15日」）
 *
 * @param {Object} task getMasterTasks_() の1件
 * @param {Date} date 判定する日
 */
function isScheduledOn_(task, date) {
  if (task.enabled === false) return false;
  const raw = String(task.freq == null ? '' : task.freq).trim();
  if (!raw) return true;                       // 未入力は「毎日」扱い
  if (/毎日|毎回|日次|随時/.test(raw)) return true;

  const dow = date.getDay();
  const dom = date.getDate();
  let s = raw;

  // 月末（「月」が曜日と誤判定されないよう先に処理）
  if (/月末|末日/.test(s)) {
    if (isMonthEnd_(date)) return true;
    s = s.replace(/月末|末日/g, '');
  }

  // 「15日」「毎月1日,15日」（曜日の「日」と区別するため数字付きのみ）
  const days = s.match(/\d{1,2}\s*日/g);
  if (days) {
    for (let i = 0; i < days.length; i++) {
      if (parseInt(days[i], 10) === dom) return true;
    }
    s = s.replace(/\d{1,2}\s*日/g, '');
  }

  // 「毎月」だけ書かれていたら 1日 とみなす
  if (/毎月/.test(s) && !days) {
    if (dom === 1) return true;
  }
  s = s.replace(/毎月|毎週|隔週/g, '');

  if (/平日/.test(s)) {
    if (dow >= 1 && dow <= 5) return true;
    s = s.replace(/平日/g, '');
  }
  if (/土日|週末/.test(s)) {
    if (dow === 0 || dow === 6) return true;
    s = s.replace(/土日|週末/g, '');
  }

  // 曜日
  return s.indexOf(WEEKDAY_JP[dow]) >= 0;
}

/** その日に実施すべき業務の一覧（区分順に並べる） */
function getScheduledTasks_(date) {
  date = date || new Date();
  return getMasterTasks_()
    .filter(function (t) { return t.name && isScheduledOn_(t, date); })
    .sort(function (a, b) {
      const pa = PHASE_ORDER.indexOf(a.phase), pb = PHASE_ORDER.indexOf(b.phase);
      const na = pa < 0 ? 99 : pa, nb = pb < 0 ? 99 : pb;
      if (na !== nb) return na - nb;
      return a.row - b.row;   // 同じ区分内はマスターの行順
    });
}

// ============================================================
// Settings.gs
// ============================================================

/**
 * ============================================================
 *  Settings.gs  ―  「設定」タブ
 * ------------------------------------------------------------
 *  店舗名・営業時間・通知先などをシート上で編集できるようにします。
 *  コードを触らずに運用を変えられるので、100店舗展開でも
 *  「同じコード＋各店の設定タブ」だけで済みます。
 * ============================================================
 */

let _cfgCache = null;

/** 設定を読み込む（1実行内はキャッシュ） */
function cfg_() {
  if (_cfgCache) return _cfgCache;
  const out = {};
  SETTING_DEFS.forEach(function (d) { out[d.key] = d.def; });

  const sh = settingsSheet_(false);
  if (sh && sh.getLastRow() >= 1) {
    const last = sh.getLastRow();
    const width = Math.max(sh.getLastColumn(), 2);
    const vals = sh.getRange(1, 1, last, width).getValues();
    const index = {};
    SETTING_DEFS.forEach(function (d) { index[norm_(d.label)] = d; });

    for (let r = 0; r < vals.length; r++) {
      const label = norm_(vals[r][0]);
      if (!label) continue;
      const d = index[label];
      if (!d) continue;
      let v = vals[r][1];
      if (v === '' || v === null) continue;
      out[d.key] = castSetting_(d, v);
    }
  }
  _cfgCache = out;
  return out;
}

/** 設定値を型に合わせて変換 */
function castSetting_(d, v) {
  if (typeof d.def === 'number') {
    const n = Number(String(v).replace(/[^\d.\-]/g, ''));
    return isNaN(n) ? d.def : n;
  }
  if (typeof d.def === 'boolean') return toBool_(v);
  if (d.key === 'OPEN_TIME' || d.key === 'CLOSE_TIME') return timeStr_(v);
  return String(v).trim();
}

/** キャッシュを捨てる（設定タブを編集したとき） */
function clearCfgCache_() { _cfgCache = null; }

/**
 * 「設定」タブを用意する。
 *  ・既にある項目は値を保持（絶対に上書きしない）
 *  ・足りない項目だけ追記する
 */
function setupSettingsSheet_() {
  const sh = settingsSheet_(true);
  const last = sh.getLastRow();

  // 既存の項目名を集める
  const existing = {};
  if (last >= 1) {
    const vals = sh.getRange(1, 1, last, 1).getValues();
    for (let r = 0; r < vals.length; r++) {
      const n = norm_(vals[r][0]);
      if (n) existing[n] = r + 1;
    }
  }

  // 見出しが無ければ付ける
  let startRow = last + 1;
  if (!existing[norm_('項目')] && last === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['項目', '値', '説明']])
      .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    startRow = 2;
  }

  // 足りない項目を追記
  const add = [];
  SETTING_DEFS.forEach(function (d) {
    if (existing[norm_(d.label)]) return;
    add.push([d.label, d.def, d.hint]);
  });
  if (add.length) {
    sh.getRange(startRow, 1, add.length, 3).setValues(add);
    // 真偽値の行はチェックボックスに
    for (let i = 0; i < add.length; i++) {
      if (typeof add[i][1] === 'boolean') sh.getRange(startRow + i, 2).insertCheckboxes();
    }
  }

  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 420);
  sh.getRange(1, 3, Math.max(sh.getLastRow(), 1), 1).setFontColor('#78909c').setFontSize(9);
  clearCfgCache_();
  return sh;
}

// ============================================================
// Master.gs
// ============================================================

/**
 * ============================================================
 *  Master.gs  ―  「各業務マスター」
 * ------------------------------------------------------------
 *  すべての起点。ここを編集すると「今日のチェック」に反映されます。
 *
 *  既存シートを尊重する設計：
 *   ・列は見出しで探し、足りない列だけ右端に追加
 *   ・既存の行データは書き換えません（空欄の既定値だけ補完）
 * ============================================================
 */

/** 業務マスターを整える（列の用意・既定値の補完・入力補助） */
function setupMasterSheet_() {
  const sh = masterSheet_(true);

  // 空シートなら見出し＋初期データを入れる
  if (sh.getLastRow() === 0) {
    const headers = MASTER_FIELDS.map(function (f) { return f.header; });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    const rows = DEFAULT_TASKS.map(function (t, i) {
      // [ID, 区分, 業務名, 目安, 頻度, 必須, 備考, 有効]
      return ['T' + ('0' + (i + 1)).slice(-2), t[0], t[1], t[2], t[3], t[4], t[5], true];
    });
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  normalizeMaster_();

  // 見た目
  const col = resolveColumns_(sh, MASTER_FIELDS);
  sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(col.id, 55);
  sh.setColumnWidth(col.phase, 75);
  sh.setColumnWidth(col.name, 170);
  sh.setColumnWidth(col.freq, 95);
  sh.setColumnWidth(col.memo, 220);
  applyPhaseColors_(sh, 2, col.phase, col.phase);
  return sh;
}

/**
 * マスターの整形。
 *  ・ID が空なら自動採番（重複しない T番号）
 *  ・区分／頻度／必須／有効／目安 の空欄に既定値を入れる
 *    → 行を足しただけで「今日のチェック」に出るようにするため
 *  ・チェックボックス／頻度プルダウンを付与
 * 変更が無ければ書き込みません（編集トリガーからの呼び出しに耐える軽さ）。
 */
function normalizeMaster_() {
  const sh = masterSheet_(true);
  const last = sh.getLastRow();
  if (last < 2) return;

  const col = resolveColumns_(sh, MASTER_FIELDS);
  const n = last - 1;
  const width = Math.max(sh.getLastColumn(), col._width);
  const vals = sh.getRange(2, 1, n, width).getValues();

  // 既存IDを集めて採番の続きを決める
  const usedIds = {};
  let maxNum = 0;
  vals.forEach(function (r) {
    const id = String(r[col.id - 1] || '').trim();
    if (!id) return;
    usedIds[id] = true;
    const m = id.match(/^T(\d+)$/i);
    if (m) maxNum = Math.max(maxNum, Number(m[1]));
  });

  // 既存の「必須」「有効」列が ○/× などの文字で運用されている場合は
  // チェックボックス化も true/false の書き込みもしない（既存の書き方を尊重）
  const boolish = function (c) {
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i][c - 1];
      if (v === '' || v === null) continue;
      if (v !== true && v !== false) return false;
    }
    return true;
  };
  const requiredIsBool = boolish(col.required);
  const enabledIsBool = boolish(col.enabled);

  const changed = {};   // key -> true
  vals.forEach(function (r) {
    if (!String(r[col.name - 1] || '').trim()) return;   // 業務名が無い行は対象外

    if (!String(r[col.id - 1] || '').trim()) {
      let id;
      do { maxNum++; id = 'T' + ('0' + maxNum).slice(-2); } while (usedIds[id]);
      usedIds[id] = true;
      r[col.id - 1] = id;
      changed.id = true;
    }
    if (!String(r[col.phase - 1] || '').trim()) { r[col.phase - 1] = DEFAULT_PHASE; changed.phase = true; }
    if (String(r[col.freq - 1] || '').trim() === '') { r[col.freq - 1] = '毎日'; changed.freq = true; }
    // 目安時間は「5分」のような書き方も使われるため、マスター側の値は書き換えない
    // （読み取るときに toMinutes_() で分に変換する）
    if (requiredIsBool && (r[col.required - 1] === '' || r[col.required - 1] === null)) {
      r[col.required - 1] = false; changed.required = true;
    }
    if (enabledIsBool && (r[col.enabled - 1] === '' || r[col.enabled - 1] === null)) {
      r[col.enabled - 1] = true; changed.enabled = true;
    }
  });

  // 変わった列だけ書き戻す
  ['id', 'phase', 'freq', 'required', 'enabled'].forEach(function (k) {
    if (!changed[k]) return;
    const c = col[k];
    const out = vals.map(function (r) { return [r[c - 1]]; });
    sh.getRange(2, c, n, 1).setValues(out);
  });

  // 入力補助
  if (requiredIsBool) sh.getRange(2, col.required, n, 1).insertCheckboxes();
  if (enabledIsBool) sh.getRange(2, col.enabled, n, 1).insertCheckboxes();
  sh.getRange(2, col.freq, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(FREQ_OPTIONS, true)
      .setAllowInvalid(true)   // 「月,木」「毎月15日」なども入力できるように
      .setHelpText('毎日 / 毎回 / 平日 / 土日 / 曜日（月〜日・複数可「月,木」）/ 月末 / 毎月15日')
      .build());
}

/**
 * 業務マスターを配列で取得。
 * @return {Array} [{row, id, phase, name, min, freq, required, memo, enabled}]
 */
function getMasterTasks_() {
  const sh = masterSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  const col = resolveColumns_(sh, MASTER_FIELDS, 1, false);
  if (!col.name) return [];    // 業務名の列が特定できない＝マスターとして読めない

  const n = sh.getLastRow() - 1;
  const width = Math.max(sh.getLastColumn(), 1);
  const vals = sh.getRange(2, 1, n, width).getValues();

  const get = function (r, key) { return col[key] ? r[col[key] - 1] : ''; };

  return vals.map(function (r, i) {
    return {
      row: i + 2,
      id: String(get(r, 'id') || '').trim() || ('R' + (i + 2)),
      phase: String(get(r, 'phase') || DEFAULT_PHASE).trim(),
      name: String(get(r, 'name') || '').trim(),
      min: toMinutes_(get(r, 'min')),   // 「5分」のような文字列でも分に変換
      freq: get(r, 'freq'),
      required: toBool_(get(r, 'required')),
      memo: get(r, 'memo') || '',
      // 有効列が無い／空欄の行は「有効」扱い（行を足しただけで表示される）
      enabled: enabledOf_(col.enabled ? get(r, 'enabled') : ''),
    };
  }).filter(function (t) { return t.name; });
}

/** 「有効」の判定。空欄は有効（＝表示する）扱い */
function enabledOf_(v) {
  if (v === '' || v === null || v === undefined) return true;
  return toBool_(v);
}

/**
 * 区分に応じた背景色を塗る。
 */
function applyPhaseColors_(sh, startRow, phaseCol, colorCol) {
  phaseCol = phaseCol || 2;
  colorCol = colorCol || phaseCol;
  const last = sh.getLastRow();
  if (last < startRow) return;
  const n = last - startRow + 1;
  const phases = sh.getRange(startRow, phaseCol, n, 1).getValues();
  const colors = phases.map(function (p) { return [PHASE_COLOR[String(p[0]).trim()] || null]; });
  sh.getRange(startRow, colorCol, n, 1).setBackgrounds(colors);
}

// ============================================================
// History.gs
// ============================================================

/**
 * ============================================================
 *  History.gs  ―  「履歴」タブへの読み書き
 * ------------------------------------------------------------
 *  集計・本部集約の元データです。
 *  キーは（日付 + 店舗ID + 業務ID）。同じキーがあれば上書き（upsert）。
 *  既存の「履歴」タブがある場合は、その列見出しをそのまま使います。
 * ============================================================
 */

/** 履歴シートの見出しを用意（既存見出しは残す） */
function setupLogSheet_() {
  const sh = logSheet_(true);
  if (sh.getLastRow() === 0) {
    const headers = LOG_FIELDS.map(function (f) { return f.header; });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    resolveColumns_(sh, LOG_FIELDS);   // 足りない列だけ追加
  }
  sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

/** 履歴の日付セルを 'yyyy/MM/dd' 文字列に正規化 */
function logDateStr_(v) {
  if (isDate_(v)) return dateStr_(v);
  return String(v == null ? '' : v).trim();
}

/**
 * 複数件をまとめて upsert（1回の読み書きで済ませる）。
 * @param {string} date 'yyyy/MM/dd'
 * @param {Array} records [{taskId,name,phase,done,time,staff,min,memo}]
 */
function writeHistoryRows_(date, records) {
  if (!records || !records.length) return;
  const sh = setupLogSheet_();
  const col = resolveColumns_(sh, LOG_FIELDS);
  const width = Math.max(sh.getLastColumn(), col._width);
  const c = cfg_();
  const now = nowTimeStr_();

  const last = sh.getLastRow();
  let vals = [];
  const index = {};
  if (last >= 2) {
    vals = sh.getRange(2, 1, last - 1, width).getValues();
    for (let i = 0; i < vals.length; i++) {
      const tid = String(vals[i][col.taskId - 1] || '').trim();
      if (!tid) continue;
      const k = logDateStr_(vals[i][col.date - 1]) + '|' +
                String(vals[i][col.storeId - 1] || '').trim() + '|' + tid;
      index[k] = i;
    }
  }

  const fill = function (row, rec) {
    row[col.date - 1]    = date;
    row[col.storeId - 1] = c.STORE_ID;
    row[col.store - 1]   = c.STORE_NAME;
    row[col.taskId - 1]  = rec.taskId;
    row[col.name - 1]    = rec.name;
    row[col.phase - 1]   = rec.phase;
    row[col.done - 1]    = rec.done ? 'TRUE' : 'FALSE';
    row[col.time - 1]    = rec.time || '';
    row[col.staff - 1]   = rec.staff || '';
    row[col.min - 1]     = rec.min || 0;
    row[col.memo - 1]    = rec.memo || '';
    row[col.updated - 1] = now;
    return row;
  };

  const appends = [];
  let dirty = false;
  records.forEach(function (rec) {
    const k = date + '|' + c.STORE_ID + '|' + rec.taskId;
    if (index[k] !== undefined) {
      fill(vals[index[k]], rec);       // 他の列（独自に追加した列）は保持される
      dirty = true;
    } else {
      const row = [];
      for (let i = 0; i < width; i++) row.push('');
      appends.push(fill(row, rec));
      index[k] = -1;                   // 同一実行内の重複追記を防ぐ
    }
  });

  if (dirty && vals.length) sh.getRange(2, 1, vals.length, width).setValues(vals);
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, width).setValues(appends);
}

/** 1件だけ upsert */
function writeHistoryRow_(date, rec) { writeHistoryRows_(date, [rec]); }

/** 指定日の履歴（この店舗分） */
function getLogByDate_(date) {
  const sh = logSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  const col = resolveColumns_(sh, LOG_FIELDS, 1, false);
  if (!col.date || !col.taskId) return [];
  const c = cfg_();
  const width = Math.max(sh.getLastColumn(), 1);
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  const g = function (r, k) { return col[k] ? r[col[k] - 1] : ''; };

  return vals
    .filter(function (r) {
      if (logDateStr_(g(r, 'date')) !== String(date)) return false;
      const sid = String(g(r, 'storeId') || '').trim();
      return !sid || sid === String(c.STORE_ID);
    })
    .map(function (r) {
      return {
        date: date,
        taskId: String(g(r, 'taskId') || '').trim(),
        name: g(r, 'name'),
        phase: g(r, 'phase'),
        done: toBool_(g(r, 'done')),
        time: timeStr_(g(r, 'time')),
        staff: g(r, 'staff') || '',
        memo: g(r, 'memo') || '',
        min: toMinutes_(g(r, 'min')),
      };
    });
}

/** 指定日の完了統計 */
function getDailyStats_(date) {
  const logs = getLogByDate_(date);
  const doneMap = {};
  logs.forEach(function (l) { doneMap[l.taskId] = l.done; });

  // 当日 …… 母数は「今日やるべき業務」。まだ触っていない業務も未完了として数える
  // 過去日 …… 確定済みの履歴をそのまま使う
  let base;
  if (String(date) === todayStr_()) {
    base = getScheduledTasks_(parseDate_(date)).map(function (t) {
      return { id: t.id, name: t.name, phase: t.phase, min: t.min, done: doneMap[t.id] === true };
    });
  } else {
    base = logs.map(function (l) {
      return { id: l.taskId, name: l.name, phase: l.phase, min: l.min, done: l.done };
    });
  }

  const byPhase = {};
  PHASE_ORDER.forEach(function (p) { byPhase[p] = { total: 0, done: 0 }; });
  const notDone = [];
  let done = 0;

  base.forEach(function (t) {
    if (!byPhase[t.phase]) byPhase[t.phase] = { total: 0, done: 0 };
    byPhase[t.phase].total++;
    if (t.done) { done++; byPhase[t.phase].done++; }
    else notDone.push(t.phase + '：' + t.name);
  });

  const total = base.length;
  return {
    date: date,
    total: total,
    done: done,
    pct: total ? Math.round((done / total) * 100) : 0,
    byPhase: byPhase,
    notDone: notDone,
  };
}

// ============================================================
// Checklist.gs
// ============================================================

/**
 * ============================================================
 *  Checklist.gs  ―  「今日のチェック」画面
 * ------------------------------------------------------------
 *  ・「各業務マスター」から、その日やるべき業務だけを並べます。
 *  ・マスターを編集すると即座に作り直されます（付けた✓は保持）。
 *  ・✓ を付けると完了時刻が入り、「履歴」に自動記録されます。
 * ============================================================
 */

// 列レイアウト
const CHK = { DONE: 1, ID: 2, PHASE: 3, NAME: 4, MIN: 5, TIME: 6, STAFF: 7, MEMO: 8 };
const CHK_COLS = 8;
const CHK_DATA_START = 5;   // データ開始行
const CHK_HEADER = ['完了', 'ID', '区分', '業務名', '目安(分)', '完了時刻', '担当', 'メモ'];

/* ============================================================
 *  生成
 * ========================================================== */

/**
 * 「今日のチェック」を（再）生成する。
 *  ・表示中の日付が対象日と違う場合は、先に前の日を履歴へ確定してから作り直す
 *  ・同じ日なら、画面上の ✓／担当／メモを引き継ぐ
 * @param {string=} targetDate 'yyyy/MM/dd'（省略時は今日）
 */
function buildChecklist_(targetDate) {
  const date = targetDate ? parseDate_(targetDate) : new Date();
  const ds = dateStr_(date);
  const sh = checkSheet_(true);
  const c = cfg_();

  // --- 直前の状態を退避 ---
  const prevDs = getCheckDate_(sh);
  let carry = {};
  if (prevDs && prevDs !== ds && sh.getLastRow() >= CHK_DATA_START) {
    finalizeDay_(prevDs, true);          // 日付が変わっていたら前日を確定
  } else if (sh.getLastRow() >= CHK_DATA_START) {
    carry = readCheckState_(sh);
  }
  // 履歴からも復元（別端末で付けた✓や、朝の自動処理後の復帰に対応）
  getLogByDate_(ds).forEach(function (l) {
    if (carry[l.taskId]) return;
    carry[l.taskId] = { done: l.done, time: l.time, staff: l.staff, memo: l.memo };
  });

  const tasks = getScheduledTasks_(date);

  // --- クリア ---
  if (sh.getMaxColumns() < CHK_COLS) sh.insertColumnsAfter(sh.getMaxColumns(), CHK_COLS - sh.getMaxColumns());
  try { sh.getRange(1, 1, Math.min(sh.getMaxRows(), 4), CHK_COLS).breakApart(); } catch (err) {}
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), CHK_COLS).clearDataValidations().clearNote();

  // --- ヘッダー部 ---
  sh.getRange(1, 1, 1, CHK_COLS).merge()
    .setValue('📅 ' + dateLabel_(date) + '　｜　店舗：' + c.STORE_NAME + '（' + c.STORE_ID + '）')
    .setFontSize(14).setFontWeight('bold')
    .setBackground('#263238').setFontColor('#ffffff')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);

  sh.getRange(2, 1, 1, CHK_COLS).merge()
    .setFontColor('#37474f').setVerticalAlignment('middle');

  sh.getRange(3, 1, 1, CHK_COLS).merge()
    .setValue('※ 業務は「' + SHEET_DEFS.MASTER.canonical + '」から自動で並びます。' +
              'マスターを編集すると、このシートもすぐ更新されます（付けた✓は消えません）。')
    .setFontColor('#90a4ae').setFontSize(9).setVerticalAlignment('middle');

  sh.getRange(4, 1, 1, CHK_COLS).setValues([CHK_HEADER])
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');

  // --- データ ---
  if (tasks.length) {
    const rows = tasks.map(function (t) {
      const st = carry[t.id] || {};
      const done = st.done === true;
      return [
        done,
        t.id,
        t.phase,
        (t.required ? '★ ' : '') + t.name,
        t.min,
        done ? (st.time || nowTimeStr_()) : '',
        st.staff || '',
        st.memo || '',
      ];
    });
    sh.getRange(CHK_DATA_START, 1, rows.length, CHK_COLS).setValues(rows);
    sh.getRange(CHK_DATA_START, CHK.DONE, rows.length, 1).insertCheckboxes();

    // マスターの備考は業務名セルのメモ（ホバー表示）に
    const notes = tasks.map(function (t) { return [t.memo ? String(t.memo) : '']; });
    sh.getRange(CHK_DATA_START, CHK.NAME, tasks.length, 1).setNotes(notes);
    sh.getRange(CHK_DATA_START, CHK.NAME, tasks.length, 1).setFontWeight('bold');
    applyPhaseColors_(sh, CHK_DATA_START, CHK.PHASE, CHK.PHASE);
    sh.getRange(CHK_DATA_START, CHK.TIME, tasks.length, 1).setHorizontalAlignment('center');
  } else {
    sh.getRange(CHK_DATA_START, 1, 1, CHK_COLS).merge()
      .setValue('今日やる業務がありません。「' + SHEET_DEFS.MASTER.canonical +
                '」の「頻度」と「有効」を確認してください。')
      .setFontColor('#d93025');
  }

  // --- 見た目 ---
  sh.setFrozenRows(4);
  sh.setColumnWidth(CHK.DONE, 50);
  sh.setColumnWidth(CHK.ID, 45);
  sh.setColumnWidth(CHK.PHASE, 70);
  sh.setColumnWidth(CHK.NAME, 200);
  sh.setColumnWidth(CHK.MIN, 65);
  sh.setColumnWidth(CHK.TIME, 80);
  sh.setColumnWidth(CHK.STAFF, 80);
  sh.setColumnWidth(CHK.MEMO, 220);

  updateChecklistStatus_(sh);
  return sh;
}

/** メニュー・トリガーから呼ぶ「今日の分に更新」 */
function refreshChecklist() {
  clearCfgCache_();
  normalizeMaster_();
  buildChecklist_();
  if (cfg_().USE_DASHBOARD) buildDashboard_();
  ss_().toast('「' + SHEET_DEFS.MASTER.canonical + '」の内容を反映しました。', '更新完了', 4);
}

/** 画面に表示中の対象日を取得（タイトル行から読む） */
function getCheckDate_(sh) {
  sh = sh || checkSheet_(false);
  if (!sh) return '';
  const title = String(sh.getRange(1, 1).getValue() || '');
  const m = title.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (!m) return '';
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
}

/** 画面上の入力状態を { 業務ID: {done,time,staff,memo} } で取得 */
function readCheckState_(sh) {
  const out = {};
  const last = sh.getLastRow();
  if (last < CHK_DATA_START) return out;
  const vals = sh.getRange(CHK_DATA_START, 1, last - CHK_DATA_START + 1, CHK_COLS).getValues();
  vals.forEach(function (r) {
    const id = String(r[CHK.ID - 1] || '').trim();
    if (!id) return;
    out[id] = {
      done: r[CHK.DONE - 1] === true,
      time: timeStr_(r[CHK.TIME - 1]),
      staff: r[CHK.STAFF - 1] || '',
      memo: r[CHK.MEMO - 1] || '',
    };
  });
  return out;
}

/** 画面の行を履歴レコードの形にする */
function checkRowsToRecords_(sh) {
  const last = sh.getLastRow();
  if (last < CHK_DATA_START) return [];
  const vals = sh.getRange(CHK_DATA_START, 1, last - CHK_DATA_START + 1, CHK_COLS).getValues();
  const out = [];
  vals.forEach(function (r) {
    const id = String(r[CHK.ID - 1] || '').trim();
    if (!id) return;
    out.push({
      taskId: id,
      name: String(r[CHK.NAME - 1] || '').replace(/^★\s*/, ''),
      phase: r[CHK.PHASE - 1],
      done: r[CHK.DONE - 1] === true,
      time: timeStr_(r[CHK.TIME - 1]),
      staff: r[CHK.STAFF - 1] || '',
      memo: r[CHK.MEMO - 1] || '',
      min: toMinutes_(r[CHK.MIN - 1]),
    });
  });
  return out;
}

/** 上部の進捗表示を更新 */
function updateChecklistStatus_(sh) {
  sh = sh || checkSheet_(false);
  if (!sh) return;
  const c = cfg_();
  const last = sh.getLastRow();
  const n = Math.max(last - CHK_DATA_START + 1, 0);
  let done = 0, total = 0;
  if (n > 0) {
    const vals = sh.getRange(CHK_DATA_START, 1, n, CHK.ID).getValues();
    vals.forEach(function (r) {
      if (!String(r[CHK.ID - 1] || '').trim()) return;
      total++;
      if (r[CHK.DONE - 1] === true) done++;
    });
  }
  const pct = total ? Math.round((done / total) * 100) : 0;
  const cell = sh.getRange(2, 1);
  cell.setValue('営業時間 ' + c.OPEN_TIME + '〜' + c.CLOSE_TIME +
                '　／　完了 ' + done + ' / ' + total + '（' + pct + '%）　' + bar_(pct));
  cell.setFontColor(pct >= 100 ? '#1e8e3e' : (pct >= 70 ? '#f9ab00' : '#546e7a'))
      .setFontWeight(pct >= 100 ? 'bold' : 'normal');
}

/* ============================================================
 *  編集の受け取り
 * ========================================================== */

/**
 * onEdit 本体。
 *  ・各業務マスターを編集 → 今日のチェックを作り直す（★これが「即反映」）
 *  ・設定を編集        → 設定を読み直して表示を更新
 *  ・今日のチェックの✓ → 完了時刻を入れて履歴に記録
 */
function handleEdit_(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const name = norm_(sh.getName());

  const master = masterSheet_(false);
  const settings = settingsSheet_(false);
  const check = checkSheet_(false);

  if (master && norm_(master.getName()) === name) {
    normalizeMaster_();
    buildChecklist_();
    ss_().toast('「今日のチェック」に反映しました。', 'マスター更新', 3);
    return;
  }

  if (settings && norm_(settings.getName()) === name) {
    clearCfgCache_();
    buildChecklist_();
    ss_().toast('設定を反映しました。', '設定更新', 3);
    return;
  }

  if (!check || norm_(check.getName()) !== name) return;

  // --- 今日のチェック上の編集 ---
  const range = e.range;
  const startRow = Math.max(range.getRow(), CHK_DATA_START);
  const endRow = range.getRow() + range.getNumRows() - 1;
  if (endRow < CHK_DATA_START) return;

  const c1 = range.getColumn();
  const c2 = c1 + range.getNumColumns() - 1;
  const touchesDone  = c1 <= CHK.DONE  && c2 >= CHK.DONE;
  const touchesInput = (c1 <= CHK.MEMO && c2 >= CHK.STAFF) || touchesDone ||
                       (c1 <= CHK.TIME && c2 >= CHK.TIME);
  if (!touchesInput) return;

  const date = getCheckDate_(sh) || todayStr_();
  const n = endRow - startRow + 1;
  const vals = sh.getRange(startRow, 1, n, CHK_COLS).getValues();

  // ✓ が付いた行に完了時刻を入れる／外れたら消す
  if (touchesDone) {
    const times = [];
    let timeChanged = false;
    for (let i = 0; i < n; i++) {
      const cur = timeStr_(vals[i][CHK.TIME - 1]);
      let next = cur;
      if (String(vals[i][CHK.ID - 1] || '').trim()) {
        if (vals[i][CHK.DONE - 1] === true && !cur) next = nowTimeStr_();
        else if (vals[i][CHK.DONE - 1] !== true) next = '';
      }
      if (next !== cur) timeChanged = true;
      times.push([next]);
      vals[i][CHK.TIME - 1] = next;
    }
    if (timeChanged) sh.getRange(startRow, CHK.TIME, n, 1).setValues(times);
  }

  // 履歴へ反映
  const records = [];
  for (let i = 0; i < n; i++) {
    const id = String(vals[i][CHK.ID - 1] || '').trim();
    if (!id) continue;
    records.push({
      taskId: id,
      name: String(vals[i][CHK.NAME - 1] || '').replace(/^★\s*/, ''),
      phase: vals[i][CHK.PHASE - 1],
      done: vals[i][CHK.DONE - 1] === true,
      time: timeStr_(vals[i][CHK.TIME - 1]),
      staff: vals[i][CHK.STAFF - 1] || '',
      memo: vals[i][CHK.MEMO - 1] || '',
      min: toMinutes_(vals[i][CHK.MIN - 1]),
    });
  }
  if (records.length) writeHistoryRows_(date, records);
  updateChecklistStatus_(sh);
}

/* ============================================================
 *  確定・日次処理
 * ========================================================== */

/**
 * その日を確定して履歴に保存（未完了も FALSE として残す）。
 * @param {string=} date 省略時は画面の対象日
 * @param {boolean=} silent トーストを出さない
 */
function finalizeDay_(date, silent) {
  const sh = checkSheet_(false);
  if (!sh) return;
  const ds = date || getCheckDate_(sh) || todayStr_();
  const records = checkRowsToRecords_(sh);
  if (records.length) writeHistoryRows_(ds, records);

  const c = cfg_();
  if (c.USE_DASHBOARD) buildDashboard_();
  if (c.CENTRAL_SS_ID) pushDailySummaryToCentral_(ds);
  if (!silent) ss_().toast(ds + ' の分を履歴に保存しました。', '確定完了', 5);
}

/** メニュー用 */
function finalizeToday() { clearCfgCache_(); finalizeDay_(); }

/** 毎朝：前日を確定 → 当日分を用意 */
function dailyMorning_() {
  clearCfgCache_();
  buildChecklist_();          // 中で前日の確定も行う
  if (cfg_().USE_DASHBOARD) buildDashboard_();
}

/** 毎晩：その日を確定（＋通知・本部送信） */
function nightlyClose_() {
  clearCfgCache_();
  finalizeDay_(null, true);
  const c = cfg_();
  if (c.ALERT_EMAIL) sendIncompleteAlert_();
}

/** テキストの進捗バー */
function bar_(pct) {
  const n = Math.max(0, Math.min(10, Math.round((pct / 100) * 10)));
  return '█'.repeat(n) + '░'.repeat(10 - n);
}

// ============================================================
// Dashboard.gs
// ============================================================

/**
 * ============================================================
 *  Dashboard.gs  ―  ダッシュボード（任意）
 * ------------------------------------------------------------
 *  「設定」タブの『ダッシュボードを使う』が TRUE のときだけ
 *  作成・更新されます（既定は OFF ＝ タブを増やしません）。
 * ============================================================
 */

function buildDashboard_() {
  const c = cfg_();
  const sh = sheetOf_('DASH', true);
  sh.clear();
  const today = todayStr_();
  const s = getDailyStats_(today);

  sh.setColumnWidth(1, 20);
  sh.setColumnWidths(2, 6, 110);

  sh.getRange(2, 2, 1, 6).merge()
    .setValue('📊 ' + c.STORE_NAME + '　業務ダッシュボード')
    .setFontSize(16).setFontWeight('bold').setFontColor('#263238');
  sh.getRange(3, 2, 1, 6).merge()
    .setValue('更新：' + todayLabel_() + ' ' + nowTimeStr_() +
              '　／　営業時間 ' + c.OPEN_TIME + '〜' + c.CLOSE_TIME)
    .setFontColor('#78909c');

  // --- 本日の完了率 ---
  sh.getRange(5, 2, 1, 2).merge().setValue('本日の完了率').setFontWeight('bold').setFontColor('#546e7a');
  sh.getRange(6, 2, 2, 2).merge()
    .setValue(s.pct + '%')
    .setFontSize(40).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontColor(s.pct >= 100 ? '#1e8e3e' : (s.pct >= 70 ? '#f9ab00' : '#d93025'));
  sh.getRange(8, 2, 1, 2).merge()
    .setValue('完了 ' + s.done + ' / ' + s.total + ' 業務')
    .setHorizontalAlignment('center').setFontColor('#546e7a');

  // --- 区分別 ---
  sh.getRange(5, 5, 1, 3).merge().setValue('区分別の進捗').setFontWeight('bold').setFontColor('#546e7a');
  let r = 6;
  PHASE_ORDER.forEach(function (p) {
    const bp = s.byPhase[p] || { total: 0, done: 0 };
    const pct = bp.total ? Math.round((bp.done / bp.total) * 100) : 0;
    sh.getRange(r, 5).setValue(p).setBackground(PHASE_COLOR[p]).setFontWeight('bold');
    sh.getRange(r, 6).setValue(bp.done + '/' + bp.total);
    sh.getRange(r, 7).setValue(bar_(pct) + ' ' + pct + '%').setFontFamily('monospace');
    r++;
  });

  // --- 想定作業時間 vs 目標 ---
  const est = estimateMinutes_();
  sh.getRange(10, 2, 1, 6).merge().setValue('⏱ 想定作業時間（施術以外）')
    .setFontWeight('bold').setFontColor('#546e7a');
  sh.getRange(11, 2, 1, 4).setValues([['区分', '想定(分)', '目標(分)', '判定']])
    .setFontWeight('bold').setBackground('#eceff1');
  [['開店前', est.before], ['閉店前', est.after]].forEach(function (row, i) {
    const ok = row[1] <= c.BUFFER_MIN;
    sh.getRange(12 + i, 2, 1, 4).setValues([[row[0], row[1], c.BUFFER_MIN, ok ? '✅ OK' : '⚠️ 超過']]);
    sh.getRange(12 + i, 5).setFontColor(ok ? '#1e8e3e' : '#d93025');
  });
  sh.getRange(14, 2, 1, 6).merge()
    .setValue('※「営業中」は施術と並行のため対象外。開店前・閉店前を目標時間内に収めるのが狙いです。')
    .setFontColor('#90a4ae').setFontSize(9);

  // --- 未完了リスト ---
  sh.getRange(16, 2, 1, 3).merge()
    .setValue('❗ 本日の未完了（' + s.notDone.length + '件）')
    .setFontWeight('bold').setFontColor('#d93025');
  if (s.notDone.length === 0) {
    sh.getRange(17, 2, 1, 3).merge().setValue('🎉 すべて完了しています！').setFontColor('#1e8e3e');
  } else {
    sh.getRange(17, 2, s.notDone.length, 1)
      .setValues(s.notDone.map(function (x) { return [x]; })).setFontColor('#c5221f');
  }

  // --- 直近7日間の推移 ---
  sh.getRange(16, 5, 1, 3).merge().setValue('📈 直近7日間の完了率')
    .setFontWeight('bold').setFontColor('#546e7a');
  const trend = getTrend_(7);
  if (trend.length) {
    sh.getRange(17, 5, trend.length, 2)
      .setValues(trend.map(function (t) { return [t.label, t.has ? t.pct + '%' : '―']; }));
    sh.getRange(17, 7, trend.length, 1)
      .setValues(trend.map(function (t) { return [t.has ? bar_(t.pct) : '']; }))
      .setFontFamily('monospace');
  }

  sh.getRange(1, 1).activate();
  return sh;
}

/** 開店前・閉店前の想定作業時間合計 */
function estimateMinutes_() {
  const tasks = getScheduledTasks_(new Date());
  let before = 0, after = 0;
  tasks.forEach(function (t) {
    if (t.phase === '開店前') before += t.min;
    if (t.phase === '閉店前') after += t.min;
  });
  return { before: before, after: after };
}

/** 直近 n 日の完了率 */
function getTrend_(n) {
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    const ds = dateStr_(d);
    const logs = getLogByDate_(ds);
    const st = getDailyStats_(ds);
    out.push({
      label: Utilities.formatDate(d, TZ, 'MM/dd') + '(' + WEEKDAY_JP[d.getDay()] + ')',
      pct: st.pct,
      has: logs.length > 0,
    });
  }
  return out;
}

// ============================================================
// CentralSync.gs
// ============================================================

/**
 * ============================================================
 *  CentralSync.gs  ―  本部（多店舗）集約
 * ------------------------------------------------------------
 *  各店舗のスプレッドシートから、本部の集約スプレッドシートへ
 *  「その日1行のサマリ」を送ります。
 *
 *  前提：
 *   ・「設定」タブの『本部集約スプレッドシートID』を入力
 *   ・そのスプレッドシートに、この Googleアカウントが編集権限を持つこと
 *   ・集約側に「店舗別サマリ」シートが無ければ自動作成します
 * ============================================================
 */

const CENTRAL_SHEET = '店舗別サマリ';
const CENTRAL_HEADER = ['日付', '店舗ID', '店舗名', '完了率', '完了数', '総数', '未完了数', '未完了業務', '更新時刻'];

/** 指定日のサマリを本部へ upsert 送信 */
function pushDailySummaryToCentral_(date) {
  const c = cfg_();
  if (!c.CENTRAL_SS_ID) return;

  let central;
  try {
    central = SpreadsheetApp.openById(c.CENTRAL_SS_ID);
  } catch (err) {
    try { ss_().toast('本部スプレッドシートを開けません。IDと共有権限を確認してください。', 'エラー', 8); } catch (e2) {}
    return;
  }

  let sh = central.getSheetByName(CENTRAL_SHEET);
  if (!sh) {
    sh = central.insertSheet(CENTRAL_SHEET);
    sh.getRange(1, 1, 1, CENTRAL_HEADER.length).setValues([CENTRAL_HEADER])
      .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }

  const s = getDailyStats_(date);
  const row = [
    date, c.STORE_ID, c.STORE_NAME, s.pct + '%',
    s.done, s.total, s.notDone.length, s.notDone.join(' / '), nowTimeStr_(),
  ];

  const last = sh.getLastRow();
  if (last >= 2) {
    const keys = sh.getRange(2, 1, last - 1, 2).getValues();
    for (let i = 0; i < keys.length; i++) {
      const d = isDate_(keys[i][0]) ? dateStr_(keys[i][0]) : String(keys[i][0]).trim();
      if (d === String(date) && String(keys[i][1]).trim() === String(c.STORE_ID)) {
        sh.getRange(i + 2, 1, 1, CENTRAL_HEADER.length).setValues([row]);
        return;
      }
    }
  }
  sh.getRange(last + 1, 1, 1, CENTRAL_HEADER.length).setValues([row]);
}

// ============================================================
// Diagnostics.gs
// ============================================================

/**
 * ============================================================
 *  Diagnostics.gs  ―  動作チェック（診断）
 * ------------------------------------------------------------
 *  「マスターを直したのに今日のチェックが変わらない」を
 *  自分で確認できるようにするための点検ツールです。
 *  メニュー「🩺 動作チェック（診断）」から実行します。
 * ============================================================
 */

function runDiagnostics() {
  clearCfgCache_();
  const lines = [];
  const ng = [];

  lines.push('■ バージョン：' + APP_VERSION);
  lines.push('■ 今日：' + todayLabel_());
  lines.push('');

  // --- タブの認識状況 ---
  lines.push('■ 使用中のタブ');
  const bind = [
    ['業務マスター', masterSheet_(false)],
    ['今日のチェック', checkSheet_(false)],
    ['履歴', logSheet_(false)],
    ['設定', settingsSheet_(false)],
  ];
  bind.forEach(function (b) {
    if (b[1]) lines.push('　✅ ' + b[0] + ' → 「' + b[1].getName() + '」');
    else { lines.push('　❌ ' + b[0] + ' → 見つかりません'); ng.push(b[0] + 'タブがありません（🏗 初期セットアップを実行）'); }
  });
  lines.push('');

  // --- マスターの読み取り ---
  const master = masterSheet_(false);
  if (master) {
    const col = resolveColumns_(master, MASTER_FIELDS, 1, false);
    const found = [];
    const missing = [];
    MASTER_FIELDS.forEach(function (f) {
      if (col[f.key]) found.push(f.header + '=' + colLetter_(col[f.key]) + '列');
      else missing.push(f.header);
    });
    lines.push('■ マスターの列の認識');
    lines.push('　' + found.join(' / '));
    if (missing.length) lines.push('　⚠ 未検出：' + missing.join('、') + '（初期セットアップで自動追加されます）');

    const tasks = getMasterTasks_();
    const todays = getScheduledTasks_(new Date());
    lines.push('　登録業務：' + tasks.length + '件　／　有効：' +
               tasks.filter(function (t) { return t.enabled; }).length + '件');
    lines.push('　今日の対象：' + todays.length + '件');
    if (!tasks.length) ng.push('マスターから業務を読み取れません（「業務名」の列見出しを確認）');
    if (tasks.length && !todays.length) ng.push('今日の対象が0件です（「頻度」「有効」を確認）');
    lines.push('');
  }

  // --- 今日のチェックの状態 ---
  const check = checkSheet_(false);
  if (check) {
    const ds = getCheckDate_(check);
    const rows = Math.max(check.getLastRow() - CHK_DATA_START + 1, 0);
    lines.push('■ 今日のチェック');
    lines.push('　表示中の日付：' + (ds || '（未生成）'));
    lines.push('　行数：' + rows + '件');
    if (ds && ds !== todayStr_()) ng.push('表示中の日付が今日ではありません（🔄 マスターの内容を反映 を実行）');
    lines.push('');
  }

  // --- 突き合わせ：マスターの今日の対象 vs 画面 ---
  if (master && check) {
    const want = getScheduledTasks_(new Date()).map(function (t) { return t.id; });
    const have = Object.keys(readCheckState_(check));
    const missingOnCheck = want.filter(function (id) { return have.indexOf(id) < 0; });
    const extraOnCheck = have.filter(function (id) { return want.indexOf(id) < 0; });
    lines.push('■ マスターと画面の一致');
    if (!missingOnCheck.length && !extraOnCheck.length) {
      lines.push('　✅ 一致しています');
    } else {
      if (missingOnCheck.length) lines.push('　⚠ 画面に出ていない業務：' + missingOnCheck.join(', '));
      if (extraOnCheck.length) lines.push('　⚠ マスターに無い業務が画面に：' + extraOnCheck.join(', '));
      ng.push('マスターと画面がずれています（🔄 マスターの内容を反映 を実行）');
    }
    lines.push('');
  }

  // --- トリガー ---
  lines.push('■ 自動化トリガー');
  const handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  const need = [
    ['onEditInstallable', 'マスター編集の即反映／✓の自動記録'],
    ['dailyMorningJob', '毎朝の当日分の用意'],
    ['nightlyCloseJob', '毎晩の確定'],
  ];
  need.forEach(function (t) {
    if (handlers.indexOf(t[0]) >= 0) lines.push('　✅ ' + t[1]);
    else { lines.push('　❌ ' + t[1] + ' 未設定'); ng.push('「⏰ 自動化トリガーを設定」が未実行です'); }
  });
  lines.push('');

  // --- 設定 ---
  const c = cfg_();
  lines.push('■ 設定');
  lines.push('　店舗：' + c.STORE_NAME + '（' + c.STORE_ID + '）　営業 ' + c.OPEN_TIME + '〜' + c.CLOSE_TIME);
  lines.push('　未完了メール：' + (c.ALERT_EMAIL || '（なし）'));
  lines.push('　本部集約ID：' + (c.CENTRAL_SS_ID ? '設定あり' : '（なし）'));
  lines.push('　ダッシュボード：' + (c.USE_DASHBOARD ? '使う' : '使わない'));
  lines.push('');

  // --- 履歴 ---
  const todayLogs = getLogByDate_(todayStr_());
  lines.push('■ 履歴');
  lines.push('　今日の記録：' + todayLogs.length + '件（完了 ' +
             todayLogs.filter(function (l) { return l.done; }).length + '件）');
  lines.push('');

  lines.push(ng.length ? '▼ 対応が必要です\n・' + uniq_(ng).join('\n・')
                       : '🎉 すべて正常です。マスターを編集すると今日のチェックに即反映されます。');

  const text = lines.join('\n');
  Logger.log(text);
  alert_('動作チェック結果', text);
  return text;
}

function uniq_(arr) {
  const seen = {}, out = [];
  arr.forEach(function (x) { if (!seen[x]) { seen[x] = true; out.push(x); } });
  return out;
}

/** 列番号 → A1 の列文字 */
function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}

// ============================================================
// Menu.gs
// ============================================================

/**
 * ============================================================
 *  Menu.gs  ―  カスタムメニュー
 * ============================================================
 */

/** シートを開いたとき */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌸 サロン業務')
    .addItem('▶ 今日のチェックを開く', 'openChecklist')
    .addItem('🔄 マスターの内容を反映（今日の分を作り直す）', 'refreshChecklist')
    .addItem('✅ 今日の分を確定（履歴に保存）', 'finalizeToday')
    .addSeparator()
    .addItem('🩺 動作チェック（診断）', 'runDiagnostics')
    .addItem('🏗 初期セットアップ（初回のみ）', 'setupAll')
    .addItem('⏰ 自動化トリガーを設定', 'installTriggers')
    .addSeparator()
    .addItem('📊 ダッシュボードを更新', 'refreshDashboard')
    .addItem('☁ 本部へ今日のサマリを送信', 'pushTodayToCentral')
    .addToUi();

  // 日付が変わっていたら開いた時点で当日分に作り直す
  try {
    const sh = checkSheet_(false);
    if (sh && getCheckDate_(sh) !== todayStr_()) buildChecklist_();
  } catch (err) {
    Logger.log('onOpen refresh skipped: ' + err);
  }
}

/** チェック画面へ移動（無ければ作る） */
function openChecklist() {
  clearCfgCache_();
  let sh = checkSheet_(false);
  if (!sh || getCheckDate_(sh) !== todayStr_()) sh = buildChecklist_();
  ss_().setActiveSheet(sh);
}

/** ダッシュボード更新（設定でOFFなら案内） */
function refreshDashboard() {
  clearCfgCache_();
  if (!cfg_().USE_DASHBOARD) {
    alert_('ダッシュボードは無効です',
      '「' + SHEET_DEFS.SETTINGS.canonical + '」タブの「ダッシュボードを使う」を TRUE（✓）にすると使えます。');
    return;
  }
  buildDashboard_();
  ss_().setActiveSheet(sheetOf_('DASH', true));
}

/** 本部送信（メニュー用） */
function pushTodayToCentral() {
  clearCfgCache_();
  const c = cfg_();
  if (!c.CENTRAL_SS_ID) {
    alert_('未設定',
      '「' + SHEET_DEFS.SETTINGS.canonical + '」タブの「本部集約スプレッドシートID」を入力してください。');
    return;
  }
  finalizeDay_(todayStr_(), true);
  pushDailySummaryToCentral_(todayStr_());
  ss_().toast('本部へ今日のサマリを送信しました。', '送信完了', 5);
}

// ============================================================
// Setup.gs
// ============================================================

/**
 * ============================================================
 *  Setup.gs  ―  初期セットアップ
 * ------------------------------------------------------------
 *  メニューの「🏗 初期セットアップ」から実行します。
 *
 *  ★既存のタブを最大限そのまま使います★
 *   「各業務マスター」「今日のチェック」「履歴」「設定」が既にあれば
 *   それを使い、無いものだけ作成します（データは消しません）。
 * ============================================================
 */

function setupAll() {
  clearCfgCache_();

  const settings = setupSettingsSheet_();
  const master   = setupMasterSheet_();
  const log      = setupLogSheet_();
  const check    = buildChecklist_();
  if (cfg_().USE_DASHBOARD) buildDashboard_();

  ss_().setActiveSheet(check);

  const msg =
    'セットアップが完了しました。\n\n' +
    '使用するタブ：\n' +
    '　・業務マスター：' + master.getName() + '\n' +
    '　・今日のチェック：' + check.getName() + '\n' +
    '　・履歴：' + log.getName() + '\n' +
    '　・設定：' + settings.getName() + '\n\n' +
    '次に「⏰ 自動化トリガーを設定」を実行してください。\n' +
    '（これを実行すると、マスターの編集が即反映されるようになります）';
  alert_('セットアップ完了', msg);
}

/** UI があればダイアログ、無ければログ＋トースト */
function alert_(title, msg) {
  try {
    SpreadsheetApp.getUi().alert(title + '\n\n' + msg);
  } catch (err) {
    Logger.log(title + ' / ' + msg);
    try { ss_().toast(msg, title, 10); } catch (e2) {}
  }
}

// ============================================================
// Triggers.gs
// ============================================================

/**
 * ============================================================
 *  Triggers.gs  ―  自動化トリガー
 * ------------------------------------------------------------
 *  1) 編集トリガー … マスター編集で「今日のチェック」を即再生成、
 *                    ✓ を履歴へ自動記録
 *  2) 毎朝        … 前日を確定 → 当日分を用意
 *  3) 毎晩        … その日を確定（＋未完了メール／本部送信）
 *
 *  ★ 1) を入れないと「マスターを直しても今日のチェックが変わらない」
 *    状態になります。必ず一度実行してください。
 * ============================================================
 */

const TRIGGER_FUNCTIONS = ['onEditInstallable', 'dailyMorningJob', 'nightlyCloseJob'];

function installTriggers() {
  clearCfgCache_();
  const c = cfg_();

  // このスクリプトが作った既存トリガーを削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (TRIGGER_FUNCTIONS.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });

  const ss = ss_();
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('dailyMorningJob').timeBased()
    .atHour(clampHour_(c.MORNING_HOUR, 6)).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('nightlyCloseJob').timeBased()
    .atHour(clampHour_(c.NIGHT_HOUR, 20)).nearMinute(30).everyDays(1).inTimezone(TZ).create();

  alert_('自動化を設定しました',
    '・「' + SHEET_DEFS.MASTER.canonical + '」の編集 → 「' + SHEET_DEFS.CHECK.canonical + '」へ即反映\n' +
    '・✓ を付けると「' + SHEET_DEFS.LOG.canonical + '」へ自動記録\n' +
    '・毎朝 ' + clampHour_(c.MORNING_HOUR, 6) + ':00 に当日分を用意\n' +
    '・毎晩 ' + clampHour_(c.NIGHT_HOUR, 20) + ':30 にその日を確定');
}

function clampHour_(v, def) {
  const n = Number(v);
  if (isNaN(n) || n < 0 || n > 23) return def;
  return Math.floor(n);
}

/** インストール型 onEdit のエントリポイント */
function onEditInstallable(e) {
  try {
    handleEdit_(e);
  } catch (err) {
    Logger.log('onEdit error: ' + err);
    try { ss_().toast('自動反映でエラー：' + err.message, 'エラー', 8); } catch (e2) {}
  }
}

/** 毎朝の処理（トリガー用） */
function dailyMorningJob() { dailyMorning_(); }

/** 毎晩の処理（トリガー用） */
function nightlyCloseJob() { nightlyClose_(); }

/** 未完了があればメール通知 */
function sendIncompleteAlert_() {
  const c = cfg_();
  if (!c.ALERT_EMAIL) return;
  const s = getDailyStats_(todayStr_());
  if (s.notDone.length === 0) return;
  const body =
    c.STORE_NAME + '（' + c.STORE_ID + '）\n' +
    todayLabel_() + ' の未完了業務：' + s.notDone.length + '件\n\n・' +
    s.notDone.join('\n・') + '\n\n完了率：' + s.pct + '%';
  MailApp.sendEmail(c.ALERT_EMAIL, '[サロン業務] 未完了あり ' + todayLabel_(), body);
}
