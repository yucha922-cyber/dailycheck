/**
 * ============================================================
 *  Master.gs  ―  「各業務マスター」
 * ------------------------------------------------------------
 *  すべての起点。ここを編集すると「今日のチェック」に反映されます。
 *
 *  既存シートを尊重する設計：
 *   ・列は見出しで探し、足りない列だけ右端に追加
 *   ・既存の行データは書き換えません（空欄の既定値だけ補完）
 * ============================================================
 */

/** 業務マスターを整える（列の用意・既定値の補完・入力補助） */
function setupMasterSheet_() {
  const sh = masterSheet_(true);

  // 空シートなら見出し＋初期データを入れる
  if (sh.getLastRow() === 0) {
    const headers = MASTER_FIELDS.map(function (f) { return f.header; });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    const rows = DEFAULT_TASKS.map(function (t, i) {
      // [ID, 区分, 業務名, 目安, 頻度, 必須, 備考, 有効]
      return ['T' + ('0' + (i + 1)).slice(-2), t[0], t[1], t[2], t[3], t[4], t[5], true];
    });
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  normalizeMaster_();

  // 見た目
  const col = resolveColumns_(sh, MASTER_FIELDS);
  sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(col.id, 55);
  sh.setColumnWidth(col.phase, 75);
  sh.setColumnWidth(col.name, 170);
  sh.setColumnWidth(col.freq, 95);
  sh.setColumnWidth(col.memo, 220);
  applyPhaseColors_(sh, 2, col.phase, col.phase);
  return sh;
}

/**
 * マスターの整形。
 *  ・ID が空なら自動採番（重複しない T番号）
 *  ・区分／頻度／必須／有効／目安 の空欄に既定値を入れる
 *    → 行を足しただけで「今日のチェック」に出るようにするため
 *  ・チェックボックス／頻度プルダウンを付与
 * 変更が無ければ書き込みません（編集トリガーからの呼び出しに耐える軽さ）。
 */
function normalizeMaster_() {
  const sh = masterSheet_(true);
  const last = sh.getLastRow();
  if (last < 2) return;

  const col = resolveColumns_(sh, MASTER_FIELDS);
  const n = last - 1;
  const width = Math.max(sh.getLastColumn(), col._width);
  const vals = sh.getRange(2, 1, n, width).getValues();

  // 既存IDを集めて採番の続きを決める
  const usedIds = {};
  let maxNum = 0;
  vals.forEach(function (r) {
    const id = String(r[col.id - 1] || '').trim();
    if (!id) return;
    usedIds[id] = true;
    const m = id.match(/^T(\d+)$/i);
    if (m) maxNum = Math.max(maxNum, Number(m[1]));
  });

  // 既存の「必須」「有効」列が ○/× などの文字で運用されている場合は
  // チェックボックス化も true/false の書き込みもしない（既存の書き方を尊重）
  const boolish = function (c) {
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i][c - 1];
      if (v === '' || v === null) continue;
      if (v !== true && v !== false) return false;
    }
    return true;
  };
  const requiredIsBool = boolish(col.required);
  const enabledIsBool = boolish(col.enabled);

  const changed = {};   // key -> true
  vals.forEach(function (r) {
    if (!String(r[col.name - 1] || '').trim()) return;   // 業務名が無い行は対象外

    if (!String(r[col.id - 1] || '').trim()) {
      let id;
      do { maxNum++; id = 'T' + ('0' + maxNum).slice(-2); } while (usedIds[id]);
      usedIds[id] = true;
      r[col.id - 1] = id;
      changed.id = true;
    }
    if (!String(r[col.phase - 1] || '').trim()) { r[col.phase - 1] = DEFAULT_PHASE; changed.phase = true; }
    if (String(r[col.freq - 1] || '').trim() === '') { r[col.freq - 1] = '毎日'; changed.freq = true; }
    // 目安時間は「5分」のような書き方も使われるため、マスター側の値は書き換えない
    // （読み取るときに toMinutes_() で分に変換する）
    if (requiredIsBool && (r[col.required - 1] === '' || r[col.required - 1] === null)) {
      r[col.required - 1] = false; changed.required = true;
    }
    if (enabledIsBool && (r[col.enabled - 1] === '' || r[col.enabled - 1] === null)) {
      r[col.enabled - 1] = true; changed.enabled = true;
    }
  });

  // 変わった列だけ書き戻す
  ['id', 'phase', 'freq', 'required', 'enabled'].forEach(function (k) {
    if (!changed[k]) return;
    const c = col[k];
    const out = vals.map(function (r) { return [r[c - 1]]; });
    sh.getRange(2, c, n, 1).setValues(out);
  });

  // 入力補助
  if (requiredIsBool) sh.getRange(2, col.required, n, 1).insertCheckboxes();
  if (enabledIsBool) sh.getRange(2, col.enabled, n, 1).insertCheckboxes();
  sh.getRange(2, col.freq, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(FREQ_OPTIONS, true)
      .setAllowInvalid(true)   // 「月,木」「毎月15日」なども入力できるように
      .setHelpText('毎日 / 毎回 / 平日 / 土日 / 曜日（月〜日・複数可「月,木」）/ 月末 / 毎月15日')
      .build());
}

/**
 * 業務マスターを配列で取得。
 * @return {Array} [{row, id, phase, name, min, freq, required, memo, enabled}]
 */
function getMasterTasks_() {
  const sh = masterSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  const col = resolveColumns_(sh, MASTER_FIELDS, 1, false);
  if (!col.name) return [];    // 業務名の列が特定できない＝マスターとして読めない

  const n = sh.getLastRow() - 1;
  const width = Math.max(sh.getLastColumn(), 1);
  const vals = sh.getRange(2, 1, n, width).getValues();

  const get = function (r, key) { return col[key] ? r[col[key] - 1] : ''; };

  return vals.map(function (r, i) {
    return {
      row: i + 2,
      id: String(get(r, 'id') || '').trim() || ('R' + (i + 2)),
      phase: String(get(r, 'phase') || DEFAULT_PHASE).trim(),
      name: String(get(r, 'name') || '').trim(),
      min: toMinutes_(get(r, 'min')),   // 「5分」のような文字列でも分に変換
      freq: get(r, 'freq'),
      required: toBool_(get(r, 'required')),
      memo: get(r, 'memo') || '',
      // 有効列が無い／空欄の行は「有効」扱い（行を足しただけで表示される）
      enabled: enabledOf_(col.enabled ? get(r, 'enabled') : ''),
    };
  }).filter(function (t) { return t.name; });
}

/** 「有効」の判定。空欄は有効（＝表示する）扱い */
function enabledOf_(v) {
  if (v === '' || v === null || v === undefined) return true;
  return toBool_(v);
}

/**
 * 区分に応じた背景色を塗る。
 */
function applyPhaseColors_(sh, startRow, phaseCol, colorCol) {
  phaseCol = phaseCol || 2;
  colorCol = colorCol || phaseCol;
  const last = sh.getLastRow();
  if (last < startRow) return;
  const n = last - startRow + 1;
  const phases = sh.getRange(startRow, phaseCol, n, 1).getValues();
  const colors = phases.map(function (p) { return [PHASE_COLOR[String(p[0]).trim()] || null]; });
  sh.getRange(startRow, colorCol, n, 1).setBackgrounds(colors);
}
