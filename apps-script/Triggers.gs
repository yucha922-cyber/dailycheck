/**
 * ============================================================
 *  Triggers.gs  ―  自動化トリガーの設定
 * ------------------------------------------------------------
 *  メニューの「⏰ 自動化トリガーを設定」から実行します。
 *   1) 編集トリガー : チェックON/OFFを日次ログへ自動記録
 *   2) 毎朝6時     : 前日を確定 → 当日分にリセット
 *   3) 毎晩20:30   : 本日分を確定（＋本部送信・通知）
 * ============================================================
 */

function installTriggers() {
  // 既存の（このスクリプトが作った）トリガーを一旦削除
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (['onEditInstallable', 'dailyFinalizeAndOpen_', 'nightlyClose_'].indexOf(fn) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });

  const ss = ss_();

  // 1) 編集トリガー
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(ss).onEdit().create();

  // 2) 毎朝6時：前日確定＆当日オープン
  ScriptApp.newTrigger('dailyFinalizeAndOpen_').timeBased().atHour(6).everyDays(1)
    .inTimezone(CONFIG.TZ).create();

  // 3) 毎晩20:30：本日確定（＋本部送信・通知）
  ScriptApp.newTrigger('nightlyClose_').timeBased().atHour(20).nearMinute(30).everyDays(1)
    .inTimezone(CONFIG.TZ).create();

  SpreadsheetApp.getUi().alert('自動化トリガーを設定しました。\n・チェック自動記録\n・毎朝6時リセット\n・毎晩20:30 確定');
}

/** インストール型 onEdit のエントリポイント */
function onEditInstallable(e) {
  handleEdit_(e);
}

/** 夜間クローズ処理 */
function nightlyClose_() {
  finalizeToday_();
  if (CONFIG.ALERT_EMAIL) sendIncompleteAlert_();
}

/** 未完了があればメール通知 */
function sendIncompleteAlert_() {
  const s = getDailyStats_(todayStr_());
  if (s.notDone.length === 0) return;
  const body =
    CONFIG.STORE_NAME + '（' + CONFIG.STORE_ID + '）\n' +
    todayLabel_() + ' の未完了業務：' + s.notDone.length + '件\n\n・' +
    s.notDone.join('\n・') + '\n\n完了率：' + s.pct + '%';
  MailApp.sendEmail(CONFIG.ALERT_EMAIL, '[サロン業務] 未完了あり ' + todayLabel_(), body);
}
