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
