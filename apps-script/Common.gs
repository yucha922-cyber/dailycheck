/**
 * ============================================================
 *  Common.gs  ―  共通ユーティリティ
 * ============================================================
 */

const WEEKDAY_JP = ['日', '月', '火', '水', '木', '金', '土'];

/** スプレッドシート本体 */
function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** シートを取得（無ければ作成） */
function getOrCreateSheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/** 今日の日付文字列 yyyy/MM/dd */
function todayStr_() {
  return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd');
}

/** 指定日を yyyy/MM/dd に整形 */
function dateStr_(d) {
  return Utilities.formatDate(d, CONFIG.TZ, 'yyyy/MM/dd');
}

/** 今の時刻 HH:mm */
function nowTimeStr_() {
  return Utilities.formatDate(new Date(), CONFIG.TZ, 'HH:mm');
}

/** 曜日つき表示用 2026/07/23(木) */
function todayLabel_() {
  const d = new Date();
  return dateStr_(d) + '(' + WEEKDAY_JP[d.getDay()] + ')';
}

/** ドキュメントプロパティ（現在チェック対象の日付を保持） */
function getCheckDate_() {
  const p = PropertiesService.getDocumentProperties().getProperty('CHECK_DATE');
  return p || todayStr_();
}
function setCheckDate_(dateStr) {
  PropertiesService.getDocumentProperties().setProperty('CHECK_DATE', dateStr);
}

/**
 * 業務マスタを配列（オブジェクト）で取得。
 * 返り値: [{id, phase, name, min, freq, required, memo, enabled}, ...]
 */
function getMasterTasks_() {
  const sh = ss_().getSheetByName(CONFIG.SHEETS.MASTER);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  return values
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        id: String(r[0]),
        phase: r[1],
        name: r[2],
        min: Number(r[3]) || 0,
        freq: r[4],
        required: r[5] === true || r[5] === 'TRUE',
        memo: r[6] || '',
        enabled: r[7] === true || r[7] === 'TRUE',
      };
    });
}

/** 今日の曜日文字（日/月/火/水/木/金/土） */
function todayWeekdayChar_() {
  return WEEKDAY_JP[new Date().getDay()];
}

/**
 * その日にチェック対象とすべき業務か（頻度で判定）。
 * ★「本日の対象（自動）」シートの FILTER 数式と完全に同じルール。
 *   頻度に「毎日」「毎回」を含む → 常に対象
 *   頻度に 今日の曜日（例:「月」）を含む → その曜日だけ対象
 *   複数曜日は「月,木」「月水金」のように書けます。
 */
function isScheduledToday_(task /*, date */) {
  if (!task.enabled) return false;
  const f = String(task.freq || '');
  if (!f) return false;
  if (f.indexOf('毎日') >= 0 || f.indexOf('毎回') >= 0) return true;
  return f.indexOf(todayWeekdayChar_()) >= 0;
}

/** 本日スケジュールされる業務一覧 */
function getScheduledTasks_() {
  return getMasterTasks_()
    .filter(function (t) { return isScheduledToday_(t); })
    .sort(function (a, b) {
      return PHASE_ORDER.indexOf(a.phase) - PHASE_ORDER.indexOf(b.phase);
    });
}
