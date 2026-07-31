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
