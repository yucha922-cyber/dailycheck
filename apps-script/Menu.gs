/**
 * ============================================================
 *  Menu.gs  ―  カスタムメニュー & トリガー設定
 * ============================================================
 */

/** スプレッドシートを開いたときにメニューを表示（簡易トリガー） */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌸 サロン業務')
    .addItem('▶ 本日のチェックを開く', 'openChecklist')
    .addItem('✅ 本日分を確定（履歴に保存）', 'finalizeToday_')
    .addItem('📊 ダッシュボードを更新', 'buildDashboard_')
    .addSeparator()
    .addItem('🏗 初期セットアップ（初回のみ）', 'setupAll')
    .addItem('⏰ 自動化トリガーを設定', 'installTriggers')
    .addSeparator()
    .addItem('☁ 本部へ本日サマリを送信', 'pushTodayToCentral')
    .addToUi();
}

/** チェック画面へ移動 */
function openChecklist() {
  const sh = ss_().getSheetByName(CONFIG.SHEETS.CHECK);
  if (sh) ss_().setActiveSheet(sh);
  else buildChecklist_();
}

/** 本部送信（メニュー用ラッパー） */
function pushTodayToCentral() {
  if (!CONFIG.CENTRAL_SS_ID) {
    ss_().toast('Config.gs の CENTRAL_SS_ID が未設定です。', '未設定', 6);
    return;
  }
  finalizeToday_();
  pushDailySummaryToCentral_(todayStr_());
  ss_().toast('本部へ本日サマリを送信しました。', '送信完了', 5);
}
