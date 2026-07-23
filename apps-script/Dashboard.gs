/**
 * ============================================================
 *  Dashboard.gs  ―  ダッシュボード（見える化）
 * ------------------------------------------------------------
 *  ・本日の完了率／区分別の進捗／未完了リスト
 *  ・直近7日間の完了率の推移
 *  ・開店前・閉店前の想定作業時間 vs 目標
 * ============================================================
 */

function buildDashboard_() {
  const sh = getOrCreateSheet_(CONFIG.SHEETS.DASH);
  sh.clear();
  const today = todayStr_();
  const s = getDailyStats_(today);

  sh.setColumnWidth(1, 20);
  sh.setColumnWidths(2, 6, 110);

  // タイトル
  sh.getRange(2, 2, 1, 6).merge()
    .setValue('📊 ' + CONFIG.STORE_NAME + '　業務ダッシュボード')
    .setFontSize(16).setFontWeight('bold').setFontColor('#263238');
  sh.getRange(3, 2, 1, 6).merge()
    .setValue('更新：' + todayLabel_() + ' ' + nowTimeStr_() +
              '　／　営業時間 ' + CONFIG.OPEN_TIME + '〜' + CONFIG.CLOSE_TIME)
    .setFontColor('#78909c');

  // --- 本日の完了率（大きく表示） ---
  sh.getRange(5, 2, 1, 2).merge().setValue('本日の完了率').setFontWeight('bold').setFontColor('#546e7a');
  sh.getRange(6, 2, 2, 2).merge()
    .setValue(s.pct + '%')
    .setFontSize(40).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
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
  sh.getRange(10, 2, 1, 6).merge().setValue('⏱ 想定作業時間（施術以外）').setFontWeight('bold').setFontColor('#546e7a');
  const timeRows = [
    ['開店前', est.before, CONFIG.BUFFER_MIN],
    ['閉店前', est.after, CONFIG.BUFFER_MIN],
  ];
  sh.getRange(11, 2, 1, 4).setValues([['区分', '想定(分)', '目標(分)', '判定']])
    .setFontWeight('bold').setBackground('#eceff1');
  timeRows.forEach(function (row, i) {
    const ok = row[1] <= row[2];
    sh.getRange(12 + i, 2, 1, 4).setValues([[row[0], row[1], row[2], ok ? '✅ OK' : '⚠️ 超過']]);
    sh.getRange(12 + i, 5).setFontColor(ok ? '#1e8e3e' : '#d93025');
  });
  sh.getRange(14, 2, 1, 6).merge()
    .setValue('※「営業中」は施術と並行のため対象外。開店前・閉店前を目標時間内に収めるのが狙いです。')
    .setFontColor('#90a4ae').setFontSize(9);

  // --- 未完了リスト ---
  sh.getRange(16, 2, 1, 6).merge()
    .setValue('❗ 本日の未完了（' + s.notDone.length + '件）').setFontWeight('bold').setFontColor('#d93025');
  if (s.notDone.length === 0) {
    sh.getRange(17, 2, 1, 6).merge().setValue('🎉 すべて完了しています！').setFontColor('#1e8e3e');
  } else {
    const nd = s.notDone.map(function (x) { return [x]; });
    sh.getRange(17, 2, nd.length, 1).setValues(nd).setFontColor('#c5221f');
  }

  // --- 直近7日間の推移 ---
  const trendStart = 16;
  sh.getRange(trendStart, 5, 1, 3).merge().setValue('📈 直近7日間の完了率').setFontWeight('bold').setFontColor('#546e7a');
  const trend = getTrend_(7);
  const trRows = trend.map(function (t) { return [t.label, t.pct]; });
  sh.getRange(trendStart + 1, 5, trRows.length, 2).setValues(trRows);
  for (let i = 0; i < trRows.length; i++) {
    sh.getRange(trendStart + 1 + i, 7)
      .setValue(bar_(trRows[i][1]) + ' ' + trRows[i][1] + '%').setFontFamily('monospace');
  }

  sh.getRange(1, 1).activate();
}

/** テキストの進捗バー */
function bar_(pct) {
  const n = Math.round((pct / 100) * 10);
  return '█'.repeat(n) + '░'.repeat(10 - n);
}

/** 開店前・閉店前の想定作業時間合計 */
function estimateMinutes_() {
  const tasks = getScheduledTasks_();
  let before = 0, after = 0;
  tasks.forEach(function (t) {
    if (t.phase === '開店前') before += t.min;
    if (t.phase === '閉店前') after += t.min;
  });
  return { before: before, after: after };
}

/** 直近 n 日の完了率 */
function getTrend_(n) {
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    const ds = dateStr_(d);
    const label = Utilities.formatDate(d, CONFIG.TZ, 'MM/dd') + '(' + WEEKDAY_JP[d.getDay()] + ')';
    const st = getDailyStats_(ds);
    // ログが無い過去日は空欄扱い（0%表示を避ける）
    const logs = getLogByDate_(ds);
    out.push({ label: label, pct: logs.length ? st.pct : 0 });
  }
  return out;
}
