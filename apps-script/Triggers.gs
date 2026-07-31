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
