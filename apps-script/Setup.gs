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
  setupAutoSheet_();       // ★関数だけの自動抽出シート
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

  // 頻度列（5列目）に入力候補プルダウンを設定（複数曜日の自由入力も許可）
  const freqRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(FREQ_OPTIONS, true)
    .setAllowInvalid(true)  // 「月,木」等の複数指定も入力できるようにする
    .setHelpText('毎日 / 毎回 / 曜日（月火水木金土日・複数可 例「月,木」）')
    .build();
  sh.getRange(2, 5, 1000, 1).setDataValidation(freqRule);

  // 見た目
  sh.getRange(1, 1, 1, header.length)
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 55);
  sh.setColumnWidth(2, 70);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(5, 90);
  sh.setColumnWidth(7, 220);
  applyPhaseColors_(sh, 2);
}

/**
 * ★「本日の対象（自動）」シート ― 関数だけで業務マスタから自動抽出。
 *   曜日判定つき FILTER 数式で、今日やるべき業務が自動で並びます。
 *   （GASは一度書くだけ。以降はマスタを直すだけで全店に反映されます）
 */
function setupAutoSheet_() {
  const sh = getOrCreateSheet_(CONFIG.SHEETS.AUTO);
  sh.clear();
  const M = CONFIG.SHEETS.MASTER;

  // 見出し
  sh.getRange(1, 1, 1, 7).merge()
    .setValue('📌 本日の対象業務（業務マスタから自動抽出／関数）')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#263238').setFontColor('#ffffff').setVerticalAlignment('middle');
  sh.setRowHeight(1, 30);

  // 日付・曜日（すべて関数）
  sh.getRange(2, 1).setValue('日付').setFontWeight('bold');
  sh.getRange(2, 2).setFormula('=TODAY()').setNumberFormat('yyyy/mm/dd (ddd)');
  sh.getRange(2, 4).setValue('本日の曜日').setFontWeight('bold');
  // 今日の曜日文字（日〜土）。この1セルを FILTER が参照します。
  sh.getRange(2, 5).setFormula('=MID("日月火水木金土", WEEKDAY(TODAY()), 1) & "曜日"');

  // ヘッダー
  const header = ['ID', '区分', '業務名', '目安(分)', '頻度', '必須', '備考'];
  sh.getRange(4, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  sh.setFrozenRows(4);

  // ★中核の数式：有効=TRUE かつ 頻度が「毎日/毎回/本日の曜日」を含む行だけ抽出
  //   （並び順は業務マスタの行順＝開店前→営業中→閉店前 をそのまま引き継ぎます）
  const todayChar = 'MID("日月火水木金土", WEEKDAY(TODAY()), 1)';
  const formula =
    "=IFERROR(" +
      "FILTER('" + M + "'!A2:G, " +
        "ARRAYFORMULA(" +
          "('" + M + "'!H2:H=TRUE) * " +
          "REGEXMATCH('" + M + "'!E2:E & \"\", \"毎日|毎回|\" & " + todayChar + ")" +
        ")" +
      "), " +
      "\"対象業務がありません（業務マスタの頻度/有効を確認）\")";
  sh.getRange(5, 1).setFormula(formula);

  // 補足
  sh.getRange(3, 1, 1, 7).merge()
    .setValue('※ このシートは自動計算です（編集不可）。実際の✓は「' + CONFIG.SHEETS.CHECK + '」で行います。')
    .setFontColor('#78909c').setFontSize(9);

  sh.setColumnWidth(1, 55);
  sh.setColumnWidth(3, 170);
  sh.setColumnWidth(7, 220);
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
  const order = [CONFIG.SHEETS.CHECK, CONFIG.SHEETS.DASH, CONFIG.SHEETS.AUTO, CONFIG.SHEETS.MASTER, CONFIG.SHEETS.LOG];
  const ss = ss_();
  order.forEach(function (name, idx) {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(idx + 1); }
  });
  const check = ss.getSheetByName(CONFIG.SHEETS.CHECK);
  if (check) ss.setActiveSheet(check);
}
