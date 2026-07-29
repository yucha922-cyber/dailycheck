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
