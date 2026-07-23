/**
 * ============================================================
 *  Setup.gs  ―  初期セットアップ
 * ------------------------------------------------------------
 *  メニューの「🏗 初期セットアップ」から実行します。
 *  必要なシート（業務マスタ / 本日のチェック / 日次ログ /
 *  ダッシュボード）をこのスプレッドシート内に自動生成します。
 * ============================================================
 */

/** すべてを初期化してセットアップ */
function setupAll() {
  setupMasterSheet_();
  setupLogSheet_();
  buildChecklist_();       // Checklist.gs
  buildDashboard_();       // Dashboard.gs
  reorderSheets_();
  SpreadsheetApp.getActive().toast('セットアップが完了しました。「本日のチェック」から使い始められます。', 'セットアップ完了', 8);
}

/** 業務マスタを作成（既にあれば中身は保持し、無ければ初期値を投入） */
function setupMasterSheet_() {
  const sh = getOrCreateSheet_(CONFIG.SHEETS.MASTER);
  const header = ['ID', '区分', '業務名', '目安(分)', '頻度', '必須', '備考', '有効'];

  if (sh.getLastRow() === 0) {
    sh.clear();
    sh.getRange(1, 1, 1, header.length).setValues([header]);

    const rows = DEFAULT_TASKS.map(function (t, i) {
      const id = 'T' + ('0' + (i + 1)).slice(-2);
      // [ID, 区分, 業務名, 目安, 頻度, 必須, 備考, 有効]
      return [id, t[0], t[1], t[2], t[3], t[4], t[5], true];
    });
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);

    // 必須・有効列はチェックボックス
    sh.getRange(2, 6, rows.length, 1).insertCheckboxes();
    sh.getRange(2, 8, rows.length, 1).insertCheckboxes();
  }

  // 見た目
  sh.getRange(1, 1, 1, header.length)
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 55);
  sh.setColumnWidth(2, 70);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(7, 220);
  applyPhaseColors_(sh, 2);
}

/** 日次ログ（履歴）シートを作成 */
function setupLogSheet_() {
  const sh = getOrCreateSheet_(CONFIG.SHEETS.LOG);
  if (sh.getLastRow() === 0) {
    const header = ['日付', '店舗ID', '店舗名', '業務ID', '業務名', '区分', '完了', '完了時刻', '目安(分)', '更新時刻'];
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.getRange(1, 1, 1, header.length)
      .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
}

/**
 * 区分に応じた背景色を付ける。
 * @param {Sheet} sh
 * @param {number} startRow データ開始行
 * @param {number} phaseCol 区分が入っている列（既定2）
 * @param {number} colorCol 色を塗る列（既定は phaseCol と同じ）
 */
function applyPhaseColors_(sh, startRow, phaseCol, colorCol) {
  phaseCol = phaseCol || 2;
  colorCol = colorCol || phaseCol;
  const last = sh.getLastRow();
  if (last < startRow) return;
  const phases = sh.getRange(startRow, phaseCol, last - startRow + 1, 1).getValues();
  for (let i = 0; i < phases.length; i++) {
    const c = PHASE_COLOR[phases[i][0]];
    if (c) sh.getRange(startRow + i, colorCol).setBackground(c);
  }
}

/** シートの並び順を整える */
function reorderSheets_() {
  const order = [CONFIG.SHEETS.CHECK, CONFIG.SHEETS.DASH, CONFIG.SHEETS.MASTER, CONFIG.SHEETS.LOG];
  const ss = ss_();
  order.forEach(function (name, idx) {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(idx + 1); }
  });
  const check = ss.getSheetByName(CONFIG.SHEETS.CHECK);
  if (check) ss.setActiveSheet(check);
}
