/**
 * ============================================================
 *  Log.gs  ―  日次ログ（履歴）への読み書き
 * ------------------------------------------------------------
 *  「日次ログ」シートは分析・集計・本部集約の元データです。
 *  キーは（日付 + 業務ID）。同じキーがあれば上書き（upsert）。
 * ============================================================
 */

const LOG_COLS = 10; // 日付,店舗ID,店舗名,業務ID,業務名,区分,完了,完了時刻,目安,更新時刻

/** 1件を upsert */
function writeLog_(date, taskId, name, phase, done, time, min) {
  const sh = getOrCreateSheet_(CONFIG.SHEETS.LOG);
  const last = sh.getLastRow();
  const row = [
    date, CONFIG.STORE_ID, CONFIG.STORE_NAME, taskId, name, phase,
    done ? 'TRUE' : 'FALSE', time || '', min, nowTimeStr_(),
  ];

  if (last >= 2) {
    const keys = sh.getRange(2, 1, last - 1, 4).getValues(); // 日付, 店舗ID, 店舗名, 業務ID
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(date) && String(keys[i][3]) === String(taskId)) {
        sh.getRange(i + 2, 1, 1, LOG_COLS).setValues([row]);
        return;
      }
    }
  }
  sh.getRange(last + 1, 1, 1, LOG_COLS).setValues([row]);
}

/** 指定日のログ行を取得（この店舗分） */
function getLogByDate_(date) {
  const sh = ss_().getSheetByName(CONFIG.SHEETS.LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, LOG_COLS).getValues();
  return values
    .filter(function (r) { return String(r[0]) === String(date); })
    .map(function (r) {
      return {
        date: r[0], storeId: r[1], storeName: r[2], taskId: r[3],
        name: r[4], phase: r[5], done: r[6] === 'TRUE' || r[6] === true,
        time: r[7], min: Number(r[8]) || 0,
      };
    });
}

/** 指定日の完了統計（この店舗） */
function getDailyStats_(date) {
  const logs = getLogByDate_(date);
  const doneMap = {};
  logs.forEach(function (l) { doneMap[l.taskId] = l.done; });

  // 当日 : 総数は「本日スケジュールされた業務」を基準にし、完了状態をログで上書き。
  //        （まだ触っていない業務も未完了として正しく数える）
  // 過去日: 確定済みのログをそのまま用いる。
  let base;
  if (String(date) === todayStr_()) {
    base = getScheduledTasks_().map(function (t) {
      return { id: t.id, name: t.name, phase: t.phase, done: doneMap[t.id] === true };
    });
  } else {
    base = logs.map(function (l) {
      return { id: l.taskId, name: l.name, phase: l.phase, done: l.done };
    });
  }
  const total = base.length;

  let done = 0;
  const byPhase = {};
  PHASE_ORDER.forEach(function (p) { byPhase[p] = { total: 0, done: 0 }; });
  const notDone = [];

  base.forEach(function (t) {
    if (!byPhase[t.phase]) byPhase[t.phase] = { total: 0, done: 0 };
    byPhase[t.phase].total++;
    if (t.done) { done++; byPhase[t.phase].done++; }
    else notDone.push(t.phase + '：' + t.name);
  });

  return {
    date: date,
    total: total,
    done: done,
    pct: total ? Math.round((done / total) * 100) : 0,
    byPhase: byPhase,
    notDone: notDone,
  };
}
