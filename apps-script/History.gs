/**
 * ============================================================
 *  History.gs  ―  「履歴」タブへの読み書き
 * ------------------------------------------------------------
 *  集計・本部集約の元データです。
 *  キーは（日付 + 店舗ID + 業務ID）。同じキーがあれば上書き（upsert）。
 *  既存の「履歴」タブがある場合は、その列見出しをそのまま使います。
 * ============================================================
 */

/** 履歴シートの見出しを用意（既存見出しは残す） */
function setupLogSheet_() {
  const sh = logSheet_(true);
  if (sh.getLastRow() === 0) {
    const headers = LOG_FIELDS.map(function (f) { return f.header; });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    resolveColumns_(sh, LOG_FIELDS);   // 足りない列だけ追加
  }
  sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

/**
 * 履歴の日付セルを 'yyyy/MM/dd' 文字列に正規化。
 * 「2026/8/1」のような0埋め無しの文字列も揃えます
 * （月間集計は文字列の大小で期間を絞るため、桁が揃っている必要があります）。
 */
function logDateStr_(v) {
  if (isDate_(v)) return dateStr_(v);
  const s = String(v == null ? '' : v).trim();
  const m = s.match(/^(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return s;
  return m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2);
}

/**
 * 複数件をまとめて upsert（1回の読み書きで済ませる）。
 * @param {string} date 'yyyy/MM/dd'
 * @param {Array} records [{taskId,name,phase,done,time,staff,min,memo}]
 */
function writeHistoryRows_(date, records) {
  if (!records || !records.length) return;
  const sh = setupLogSheet_();
  const col = resolveColumns_(sh, LOG_FIELDS);
  const width = Math.max(sh.getLastColumn(), col._width);
  const c = cfg_();
  const now = nowTimeStr_();

  const last = sh.getLastRow();
  let vals = [];
  const index = {};
  if (last >= 2) {
    vals = sh.getRange(2, 1, last - 1, width).getValues();
    for (let i = 0; i < vals.length; i++) {
      const tid = String(vals[i][col.taskId - 1] || '').trim();
      if (!tid) continue;
      const k = logDateStr_(vals[i][col.date - 1]) + '|' +
                String(vals[i][col.storeId - 1] || '').trim() + '|' + tid;
      index[k] = i;
    }
  }

  const fill = function (row, rec) {
    row[col.date - 1]    = date;
    row[col.storeId - 1] = c.STORE_ID;
    row[col.store - 1]   = c.STORE_NAME;
    row[col.taskId - 1]  = rec.taskId;
    row[col.name - 1]    = rec.name;
    row[col.phase - 1]   = rec.phase;
    row[col.done - 1]    = rec.done ? 'TRUE' : 'FALSE';
    row[col.time - 1]    = rec.time || '';
    row[col.staff - 1]   = rec.staff || '';
    row[col.min - 1]     = rec.min || 0;
    row[col.memo - 1]    = rec.memo || '';
    row[col.updated - 1] = now;
    return row;
  };

  const appends = [];
  let dirty = false;
  records.forEach(function (rec) {
    const k = date + '|' + c.STORE_ID + '|' + rec.taskId;
    if (index[k] !== undefined) {
      fill(vals[index[k]], rec);       // 他の列（独自に追加した列）は保持される
      dirty = true;
    } else {
      const row = [];
      for (let i = 0; i < width; i++) row.push('');
      appends.push(fill(row, rec));
      index[k] = -1;                   // 同一実行内の重複追記を防ぐ
    }
  });

  if (dirty && vals.length) sh.getRange(2, 1, vals.length, width).setValues(vals);
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, width).setValues(appends);
}

/** 1件だけ upsert */
function writeHistoryRow_(date, rec) { writeHistoryRows_(date, [rec]); }

/** 履歴の1行 → 扱いやすいオブジェクトに */
function logRowToObj_(date, r, col) {
  const g = function (k) { return col[k] ? r[col[k] - 1] : ''; };
  return {
    date: date,
    taskId: String(g('taskId') || '').trim(),
    name: g('name'),
    phase: g('phase'),
    done: toBool_(g('done')),
    time: timeStr_(g('time')),
    staff: g('staff') || '',
    memo: g('memo') || '',
    min: toMinutes_(g('min')),
  };
}

/**
 * 期間内（この店舗分）の履歴を、日付ごとにまとめて返す。
 * 月間集計のように何日分も見るときは、これで1回だけ読み込みます。
 * @param {string} from 'yyyy/MM/dd'（含む・省略可）
 * @param {string} to   'yyyy/MM/dd'（含む・省略可）
 * @return {Object} { 'yyyy/MM/dd': [ログ, ...], ... }
 */
function getLogsGroupedByDate_(from, to) {
  const out = {};
  const sh = logSheet_(false);
  if (!sh || sh.getLastRow() < 2) return out;
  const col = resolveColumns_(sh, LOG_FIELDS, 1, false);
  if (!col.date || !col.taskId) return out;
  const c = cfg_();
  const width = Math.max(sh.getLastColumn(), 1);
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();

  vals.forEach(function (r) {
    const ds = logDateStr_(r[col.date - 1]);
    if (!ds) return;
    if (from && ds < String(from)) return;      // 'yyyy/MM/dd' は文字列比較で日付順になる
    if (to && ds > String(to)) return;
    const sid = String((col.storeId ? r[col.storeId - 1] : '') || '').trim();
    if (sid && sid !== String(c.STORE_ID)) return;
    (out[ds] = out[ds] || []).push(logRowToObj_(ds, r, col));
  });
  return out;
}

/** 指定日の履歴（この店舗分） */
function getLogByDate_(date) {
  return getLogsGroupedByDate_(date, date)[String(date)] || [];
}

/** 指定日の完了統計 */
function getDailyStats_(date) {
  return statsFromLogs_(date, getLogByDate_(date));
}

/**
 * 読み込み済みの履歴から、その日の完了統計を作る。
 * @param {string} date 'yyyy/MM/dd'
 * @param {Array} logs  その日の履歴（getLogsGroupedByDate_ の1日分）
 */
function statsFromLogs_(date, logs) {
  logs = logs || [];
  const doneMap = {}, timeMap = {};
  logs.forEach(function (l) { doneMap[l.taskId] = l.done; timeMap[l.taskId] = l.time; });

  // 当日 …… 母数は「今日やるべき業務」。まだ触っていない業務も未完了として数える
  // 過去日 …… 確定済みの履歴をそのまま使う
  let base;
  if (String(date) === todayStr_()) {
    base = getScheduledTasks_(parseDate_(date)).map(function (t) {
      return { id: t.id, name: t.name, phase: t.phase, min: t.min, done: doneMap[t.id] === true };
    });
  } else {
    base = logs.map(function (l) {
      return { id: l.taskId, name: l.name, phase: l.phase, min: l.min, done: l.done };
    });
  }

  const byPhase = {};
  PHASE_ORDER.forEach(function (p) { byPhase[p] = { total: 0, done: 0 }; });
  const notDone = [];
  let done = 0;
  let lastMin = -1;             // 最後に完了した業務の時刻（＝100%に達した時刻）

  base.forEach(function (t) {
    if (!byPhase[t.phase]) byPhase[t.phase] = { total: 0, done: 0 };
    byPhase[t.phase].total++;
    if (t.done) {
      done++;
      byPhase[t.phase].done++;
      const mm = timeToMin_(timeMap[t.id]);
      if (mm > lastMin) lastMin = mm;
    } else {
      notDone.push(t.phase + '：' + t.name);
    }
  });

  const total = base.length;
  const full = total > 0 && done === total;
  return {
    date: date,
    total: total,
    done: done,
    pct: total ? Math.round((done / total) * 100) : 0,
    byPhase: byPhase,
    notDone: notDone,
    // 全業務が完了した日だけ、その達成時刻を 'HH:mm' で返す（時刻未記録なら空）
    doneAt: full && lastMin >= 0 ? minToTimeStr_(lastMin) : '',
    isFull: full,
  };
}
