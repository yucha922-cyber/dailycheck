/**
 * ============================================================
 *  Dashboard.gs  ―  ダッシュボード（任意）
 * ------------------------------------------------------------
 *  「設定」タブの『ダッシュボードを使う』が TRUE のときだけ
 *  作成・更新されます（既定は OFF ＝ タブを増やしません）。
 * ============================================================
 */

const DASH_COLS = 9;    // A〜I（H列＝100%達成時刻／I列＝帰宅時間）

function buildDashboard_() {
  const c = cfg_();
  const sh = sheetOf_('DASH', true);
  if (sh.getMaxColumns() < DASH_COLS) {
    sh.insertColumnsAfter(sh.getMaxColumns(), DASH_COLS - sh.getMaxColumns());
  }
  // 前回のレイアウトの結合が残っていると書き込めないので一度解除する
  try {
    sh.getRange(1, 1, sh.getMaxRows(), DASH_COLS).breakApart();
  } catch (err) {}
  sh.clear();
  const today = todayStr_();
  const s = getDailyStats_(today);

  sh.setColumnWidth(1, 20);
  sh.setColumnWidths(2, 6, 110);
  sh.setColumnWidth(8, 110);          // H列（100%達成時刻）
  sh.setColumnWidth(9, 110);          // I列（帰宅時間）

  sh.getRange(2, 2, 1, 8).merge()
    .setValue('📊 ' + c.STORE_NAME + '　業務ダッシュボード')
    .setFontSize(16).setFontWeight('bold').setFontColor('#263238');
  sh.getRange(3, 2, 1, 8).merge()
    .setValue('更新：' + todayLabel_() + ' ' + nowTimeStr_() +
              '　／　営業時間 ' + c.OPEN_TIME + '〜' + c.CLOSE_TIME)
    .setFontColor('#78909c');

  // --- 本日の完了率 ---
  sh.getRange(5, 2, 1, 2).merge().setValue('本日の完了率').setFontWeight('bold').setFontColor('#546e7a');
  sh.getRange(6, 2, 2, 2).merge()
    .setValue(s.pct + '%')
    .setFontSize(40).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontColor(s.pct >= 100 ? '#1e8e3e' : (s.pct >= 70 ? '#f9ab00' : '#d93025'));
  sh.getRange(8, 2, 1, 2).merge()
    .setValue('完了 ' + s.done + ' / ' + s.total + ' 業務')
    .setHorizontalAlignment('center').setFontColor('#546e7a');

  // --- 区分別 ---
  sh.getRange(5, 5, 1, 3).merge().setValue('区分別の進捗').setFontWeight('bold').setFontColor('#546e7a');
  let r = 6;
  PHASE_ORDER.forEach(function (p) {
    const bp = s.byPhase[p] || { total: 0, done: 0 };
    const pct = bp.total ? Math.round((bp.done / bp.total) * 100) : 0;
    sh.getRange(r, 5).setValue(p).setBackground(PHASE_COLOR[p]).setFontWeight('bold');
    sh.getRange(r, 6).setValue(bp.done + '/' + bp.total);
    sh.getRange(r, 7).setValue(bar_(pct) + ' ' + pct + '%').setFontFamily('monospace');
    r++;
  });

  // --- 想定作業時間 vs 目標 ---
  const est = estimateMinutes_();
  sh.getRange(10, 2, 1, 8).merge().setValue('⏱ 想定作業時間（施術以外）')
    .setFontWeight('bold').setFontColor('#546e7a');
  sh.getRange(11, 2, 1, 4).setValues([['区分', '想定(分)', '目標(分)', '判定']])
    .setFontWeight('bold').setBackground('#eceff1');
  [['開店前', est.before], ['閉店前', est.after]].forEach(function (row, i) {
    const ok = row[1] <= c.BUFFER_MIN;
    sh.getRange(12 + i, 2, 1, 4).setValues([[row[0], row[1], c.BUFFER_MIN, ok ? '✅ OK' : '⚠️ 超過']]);
    sh.getRange(12 + i, 5).setFontColor(ok ? '#1e8e3e' : '#d93025');
  });
  sh.getRange(14, 2, 1, 8).merge()
    .setValue('※「営業中」は施術と並行のため対象外。開店前・閉店前を目標時間内に収めるのが狙いです。')
    .setFontColor('#90a4ae').setFontSize(9);

  // --- 未完了リスト ---
  sh.getRange(16, 2, 1, 3).merge()
    .setValue('❗ 本日の未完了（' + s.notDone.length + '件）')
    .setFontWeight('bold').setFontColor('#d93025');
  if (s.notDone.length === 0) {
    sh.getRange(17, 2, 1, 3).merge().setValue('🎉 すべて完了しています！').setFontColor('#1e8e3e');
  } else {
    sh.getRange(17, 2, s.notDone.length, 1)
      .setValues(s.notDone.map(function (x) { return [x]; })).setFontColor('#c5221f');
  }

  // --- 1ヶ月間（当月分すべて）の推移 ---
  const month = new Date();
  sh.getRange(16, 5, 1, 5).merge()
    .setValue('📈 ' + monthLabel_(month) + 'の完了率（1ヶ月間）')
    .setFontWeight('bold').setFontColor('#546e7a');
  sh.getRange(17, 5, 1, 5).setValues([['日付', '完了率', '推移', '100%達成時刻', '帰宅時間']])
    .setFontWeight('bold').setBackground('#eceff1').setFontColor('#546e7a');

  const trend = getMonthTrend_(month);
  const TREND_START = 18;
  if (trend.length) {
    sh.getRange(TREND_START, 5, trend.length, 2)
      .setValues(trend.map(function (t) { return [t.label, t.has ? t.pct + '%' : '―']; }));
    sh.getRange(TREND_START, 7, trend.length, 1)
      .setValues(trend.map(function (t) { return [t.has ? bar_(t.pct) : '']; }))
      .setFontFamily('monospace');
    // H列：その日 100% に達した時刻（未達成・データ無しは「―」）
    sh.getRange(TREND_START, 8, trend.length, 1)
      .setValues(trend.map(function (t) { return [fullTimeCell_(t)]; }))
      .setHorizontalAlignment('center');
    // I列：「今日のチェック」の帰宅時間欄に入力された、院を出た時刻
    sh.getRange(TREND_START, 9, trend.length, 1)
      .setValues(trend.map(function (t) { return [t.leaveAt || '―']; }))
      .setHorizontalAlignment('center');
  }

  sh.getRange(1, 1).activate();
  return sh;
}

/** 「2026年8月」のような月ラベル */
function monthLabel_(d) {
  return d.getFullYear() + '年' + (d.getMonth() + 1) + '月';
}

/** H列に出す「100%達成時刻」の表示文字列 */
function fullTimeCell_(t) {
  if (!t.has) return '―';                 // その日の記録がまだ無い
  if (t.pct < 100) return '―';            // 未達成
  return t.doneAt || '時刻なし';           // 全完了だが完了時刻が記録されていない
}

/** 開店前・閉店前の想定作業時間合計 */
function estimateMinutes_() {
  const tasks = getScheduledTasks_(new Date());
  let before = 0, after = 0;
  tasks.forEach(function (t) {
    if (t.phase === '開店前') before += t.min;
    if (t.phase === '閉店前') after += t.min;
  });
  return { before: before, after: after };
}

/**
 * その月の日ごとの完了率（1日〜）。
 *  ・当月 …… 1日から「今日」まで（先の日付は出しません）
 *  ・過去月 … 1日から月末まで
 * 履歴シートは1回だけ読み、日付ごとに振り分けて集計します。
 * @param {Date=} base 対象月に含まれる任意の日（省略時は今日）
 */
function getMonthTrend_(base) {
  base = base || new Date();
  const y = base.getFullYear(), m = base.getMonth();
  const now = new Date();
  const isCurrentMonth = (y === now.getFullYear() && m === now.getMonth());
  const lastDay = isCurrentMonth ? now.getDate() : daysInMonth_(y, m);

  const grouped = getLogsGroupedByDate_(dateStr_(new Date(y, m, 1)),
                                        dateStr_(new Date(y, m, lastDay)));
  const out = [];
  for (let day = 1; day <= lastDay; day++) {
    const d = new Date(y, m, day);
    const ds = dateStr_(d);
    const logs = grouped[ds] || [];
    const st = statsFromLogs_(ds, logs);
    out.push({
      label: Utilities.formatDate(d, TZ, 'MM/dd') + '(' + WEEKDAY_JP[d.getDay()] + ')',
      pct: st.pct,
      has: logs.length > 0,
      doneAt: st.doneAt,
      leaveAt: st.leaveAt,
    });
  }
  return out;
}
