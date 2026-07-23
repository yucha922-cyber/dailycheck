/**
 * ============================================================
 *  Checklist.gs  ―  「本日のチェック」画面
 * ------------------------------------------------------------
 *  ・毎朝この画面のチェックボックスを ON にしていくだけで運用。
 *  ・ON/OFF は自動で「日次ログ」に記録されます（onEdit）。
 *  ・毎朝6時に自動で前日を確定し、当日分にリセットします。
 * ============================================================
 */

// 「本日のチェック」の列レイアウト
const CHK = { DONE: 1, ID: 2, PHASE: 3, NAME: 4, MIN: 5, TIME: 6, STAFF: 7, MEMO: 8 };
const CHK_DATA_START = 5;   // データ開始行
const CHK_COLS = 8;

/** 本日のチェック画面を（再）生成 */
function buildChecklist_() {
  const sh = getOrCreateSheet_(CONFIG.SHEETS.CHECK);
  sh.clear();
  setCheckDate_(todayStr_());

  const tasks = getScheduledTasks_();

  // --- タイトル・ステータス ---
  sh.getRange(1, 1, 1, CHK_COLS).merge()
    .setValue('📅 ' + todayLabel_() + '　｜　店舗：' + CONFIG.STORE_NAME + '（' + CONFIG.STORE_ID + '）')
    .setFontSize(14).setFontWeight('bold')
    .setBackground('#263238').setFontColor('#ffffff')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);

  sh.getRange(2, 1, 1, CHK_COLS).merge()
    .setValue('営業時間 ' + CONFIG.OPEN_TIME + '〜' + CONFIG.CLOSE_TIME +
              '　／　目標：開店前・閉店前それぞれ ' + CONFIG.BUFFER_MIN + '分以内で完了')
    .setFontColor('#546e7a').setVerticalAlignment('middle');

  // --- ヘッダー ---
  const header = ['完了', 'ID', '区分', '業務名', '目安(分)', '完了時刻', '担当', 'メモ'];
  sh.getRange(4, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');

  // --- データ ---
  const rows = tasks.map(function (t) {
    return [false, t.id, t.phase, (t.required ? '★' : '　') + t.name, t.min, '', '', t.memo];
  });
  if (rows.length) {
    sh.getRange(CHK_DATA_START, 1, rows.length, CHK_COLS).setValues(rows);
    sh.getRange(CHK_DATA_START, CHK.DONE, rows.length, 1).insertCheckboxes();
    applyPhaseColors_(sh, CHK_DATA_START, CHK.PHASE, CHK.PHASE);
  }

  // 見た目
  sh.setFrozenRows(4);
  sh.setColumnWidth(CHK.DONE, 50);
  sh.setColumnWidth(CHK.ID, 45);
  sh.setColumnWidth(CHK.PHASE, 70);
  sh.setColumnWidth(CHK.NAME, 190);
  sh.setColumnWidth(CHK.MIN, 65);
  sh.setColumnWidth(CHK.TIME, 80);
  sh.setColumnWidth(CHK.STAFF, 80);
  sh.setColumnWidth(CHK.MEMO, 220);
  sh.getRange(CHK_DATA_START, CHK.NAME, Math.max(rows.length, 1), 1).setFontWeight('bold');

  updateChecklistStatus_();
}

/** チェック画面上部の進捗表示を更新 */
function updateChecklistStatus_() {
  const sh = ss_().getSheetByName(CONFIG.SHEETS.CHECK);
  if (!sh) return;
  const last = sh.getLastRow();
  const n = last - CHK_DATA_START + 1;
  if (n <= 0) return;
  const dones = sh.getRange(CHK_DATA_START, CHK.DONE, n, 1).getValues();
  let done = 0;
  dones.forEach(function (r) { if (r[0] === true) done++; });
  const pct = Math.round((done / n) * 100);
  sh.getRange(2, 1, 1, CHK_COLS).getCell(1, 1)
    .setValue('営業時間 ' + CONFIG.OPEN_TIME + '〜' + CONFIG.CLOSE_TIME +
              '　／　本日の完了 ' + done + '/' + n + '（' + pct + '%）');
}

/**
 * onEdit（インストール型トリガー）本体。
 * 完了列がトグルされたら、時刻を記録し、日次ログへ反映。
 */
function handleEdit_(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== CONFIG.SHEETS.CHECK) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (col !== CHK.DONE || row < CHK_DATA_START) return;

  const checked = e.range.getValue() === true;
  const timeCell = sh.getRange(row, CHK.TIME);
  if (checked) timeCell.setValue(nowTimeStr_());
  else timeCell.clearContent();

  upsertLogRow_(sh, row);
  updateChecklistStatus_();
}

/** チェック画面の1行を日次ログへ upsert */
function upsertLogRow_(sh, row) {
  const vals = sh.getRange(row, 1, 1, CHK_COLS).getValues()[0];
  const done = vals[CHK.DONE - 1] === true;
  const id = vals[CHK.ID - 1];
  const name = String(vals[CHK.NAME - 1]).replace(/^[★　]/, '');
  const phase = vals[CHK.PHASE - 1];
  const min = vals[CHK.MIN - 1];
  const time = vals[CHK.TIME - 1];
  const date = getCheckDate_();
  writeLog_(date, id, name, phase, done, time, min);
}

/**
 * 「本日分を確定」— チェックの有無に関わらず全業務をログに書き込む。
 * （未完了も FALSE として履歴に残すことで、集計が正確になります）
 */
function finalizeToday_() {
  const sh = ss_().getSheetByName(CONFIG.SHEETS.CHECK);
  if (!sh) return;
  const last = sh.getLastRow();
  for (let row = CHK_DATA_START; row <= last; row++) {
    upsertLogRow_(sh, row);
  }
  buildDashboard_();
  if (CONFIG.CENTRAL_SS_ID) pushDailySummaryToCentral_(getCheckDate_());
  ss_().toast('本日分を確定し、履歴に保存しました。', '確定完了', 5);
}

/**
 * 毎朝の自動処理：前日を確定 → 当日分にリセット。
 * 時間主導トリガー（毎朝6時）から呼ばれます。
 */
function dailyFinalizeAndOpen_() {
  // 直前にチェック画面に表示されていた日を確定
  finalizeToday_();
  // 当日分を開く
  buildChecklist_();
  buildDashboard_();
}
