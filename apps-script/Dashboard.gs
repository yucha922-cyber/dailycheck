/**
 * ============================================================
 *  Dashboard.gs  ―  ダッシュボード（任意）
 * ------------------------------------------------------------
 *  「設定」タブの『ダッシュボードを使う』が TRUE のときだけ
 *  作成・更新されます（既定は OFF ＝ タブを増やしません）。
 *
 *  画面の構成
 *   上 … 本日の完了率／区分別／想定作業時間／未完了リスト
 *   　　＋ 今月の完了率（1ヶ月間の推移）
 *   下 … 月次アーカイブ（先月と今月を横に並べた比較表）
 *        月が変わると、終わった月はそのままこの下の表に残ります。
 * ============================================================
 */

const DASH_COLS = 9;            // 上段が使う列 A〜I（H列＝100%達成時刻／I列＝帰宅時間）

/** 1ヶ月ぶんの表（日付〜帰宅時間）の見出しと幅 */
const TREND_HEAD = ['日付', '完了率', '推移', '100%達成時刻', '帰宅時間'];
const TREND_W = TREND_HEAD.length;   // 5列
const BLOCK_STEP = TREND_W + 1;      // 月と月の間に1列あける（B〜F・H〜L…）

/** 下段（月次アーカイブ）を置きはじめる行。上段と重なるときは自動で下にずれます */
const ARCHIVE_TOP_ROW = 30;

/** 上段「今月の推移」の位置 */
const TREND_HEAD_ROW = 17;
const TREND_TOP_ROW = 18;
const TREND_COL = 5;            // E列

function buildDashboard_() {
  const c = cfg_();
  const sh = sheetOf_('DASH', true);

  // 過去何ヶ月ぶんを今月と並べて比較するか（既定は1＝先月と今月）
  const compare = Math.max(0, Math.min(6, Math.round(c.COMPARE_MONTHS || 0)));
  const months = archiveMonths_(compare);          // 古い月 → 今月 の順
  const needCols = Math.max(DASH_COLS, archiveCol_(months.length - 1) + TREND_W - 1);

  if (sh.getMaxColumns() < needCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  }
  // 前回のレイアウトの結合が残っていると書き込めないので一度解除する
  // （比較する月数を減らしたときのために、シート全体を対象にします）
  try {
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  } catch (err) {}
  sh.clear();
  const today = todayStr_();
  const s = getDailyStats_(today);

  sh.setColumnWidth(1, 20);
  sh.setColumnWidths(2, needCols - 1, 110);

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
  let leftEnd = 17;
  if (s.notDone.length === 0) {
    sh.getRange(17, 2, 1, 3).merge().setValue('🎉 すべて完了しています！').setFontColor('#1e8e3e');
  } else {
    sh.getRange(17, 2, s.notDone.length, 1)
      .setValues(s.notDone.map(function (x) { return [x]; })).setFontColor('#c5221f');
    leftEnd = 16 + s.notDone.length;
  }

  // 履歴シートは1回だけ読み、上段（今月）と下段（アーカイブ）で使い回す
  const grouped = getLogsGroupedByDate_(dateStr_(months[0]), today);

  // --- 上段：今月の推移（1ヶ月間） ---
  const month = new Date();
  sh.getRange(16, TREND_COL, 1, TREND_W).merge()
    .setValue('📈 ' + monthLabel_(month) + 'の完了率（1ヶ月間）')
    .setFontWeight('bold').setFontColor('#546e7a');
  sh.getRange(TREND_HEAD_ROW, TREND_COL, 1, TREND_W).setValues([TREND_HEAD])
    .setFontWeight('bold').setBackground('#eceff1').setFontColor('#546e7a');

  const trend = getMonthTrend_(month, grouped);
  writeTrendRows_(sh, TREND_TOP_ROW, TREND_COL, trend);
  const trendEnd = TREND_TOP_ROW + trend.length - 1;

  // --- 下段：月次アーカイブ（先月 ⇔ 今月） ---
  if (compare > 0) {
    const top = Math.max(ARCHIVE_TOP_ROW, trendEnd + 2, leftEnd + 2);
    writeMonthArchive_(sh, top, months, grouped);
  }

  sh.getRange(1, 1).activate();
  return sh;
}

/* ============================================================
 *  月次アーカイブ（先月と今月を横に並べる）
 * ========================================================== */

/**
 * 今月と、その手前 count ヶ月ぶんの「月の初日」を古い順に返す。
 * count=1 なら [先月, 今月]。
 */
function archiveMonths_(count) {
  const now = new Date();
  const out = [];
  for (let i = count; i >= 0; i--) {
    out.push(new Date(now.getFullYear(), now.getMonth() - i, 1));   // 月をまたいでも年は自動で繰り上がる
  }
  return out;
}

/** アーカイブの i 番目（0始まり・古い月から）の表が始まる列。0→B, 1→H, 2→N … */
function archiveCol_(i) {
  return 2 + i * BLOCK_STEP;
}

/**
 * 月ごとの表を横に並べて書き出す。
 * @param {Sheet} sh
 * @param {number} startRow 見出しを置く行
 * @param {Array<Date>} months 古い月 → 今月 の順
 * @param {Object} grouped getLogsGroupedByDate_() の結果（読み込み済みの履歴）
 * @return {number} 使い終わった次の行
 */
function writeMonthArchive_(sh, startRow, months, grouped) {
  const lastCol = archiveCol_(months.length - 1) + TREND_W - 1;
  const wide = lastCol - 1;                   // B列から右端までの幅

  const headRow = startRow + 4;               // 「日付・完了率…」の見出し
  const topRow = headRow + 1;                 // 日ごとのデータの1行目
  ensureRows_(sh, topRow + 31);

  sh.getRange(startRow, 2, 1, wide).merge()
    .setValue('🗂 月次アーカイブ（先月と今月の比較）')
    .setFontWeight('bold').setFontColor('#546e7a');
  sh.getRange(startRow + 1, 2, 1, wide).merge()
    .setValue('※ 左が過去の月、右端が今月です。月が変わると、終わった月の結果はそのままこの表に残ります。')
    .setFontColor('#90a4ae').setFontSize(9);

  const blocks = months.map(function (d) {
    const tr = getMonthTrend_(d, grouped);
    return { date: d, trend: tr, sum: monthSummary_(tr) };
  });

  let used = 0;
  blocks.forEach(function (blk, i) {
    const col = archiveCol_(i);
    const isCurrent = (i === blocks.length - 1);
    const isPrev = (i === blocks.length - 2);

    sh.getRange(startRow + 2, col, 1, TREND_W).merge()
      .setValue(monthLabel_(blk.date) + (isCurrent ? '（今月）' : (isPrev ? '（先月）' : '')))
      .setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground(isCurrent ? '#e8f0fe' : '#eceff1')
      .setFontColor('#263238');
    sh.getRange(startRow + 3, col, 1, TREND_W).merge()
      .setValue(monthSummaryText_(blk.sum, i > 0 ? blocks[i - 1].sum : null))
      .setFontSize(9).setFontColor('#546e7a');
    sh.getRange(headRow, col, 1, TREND_W).setValues([TREND_HEAD])
      .setFontWeight('bold').setBackground('#eceff1').setFontColor('#546e7a');

    // 記録が1日も無い過去の月は、日付を31行並べても意味が無いので1行で知らせる
    if (!blk.sum.days && !isCurrent) {
      sh.getRange(topRow, col, 1, TREND_W).merge()
        .setValue('この月の記録はありません').setFontColor('#90a4ae');
      used = Math.max(used, 1);
      return;
    }
    writeTrendRows_(sh, topRow, col, blk.trend);
    used = Math.max(used, blk.trend.length);
  });

  return topRow + used;
}

/**
 * 1ヶ月ぶんの日次データ（日付・完了率・推移・100%達成時刻・帰宅時間）を書き出す。
 * 上段の「今月の推移」と下段のアーカイブで共通です。
 */
function writeTrendRows_(sh, row, col, trend) {
  if (!trend || !trend.length) return;
  sh.getRange(row, col, trend.length, 2)
    .setValues(trend.map(function (t) { return [t.label, t.has ? t.pct + '%' : '―']; }));
  sh.getRange(row, col + 2, trend.length, 1)
    .setValues(trend.map(function (t) { return [t.has ? bar_(t.pct) : '']; }))
    .setFontFamily('monospace');
  // その日 100% に達した時刻（未達成・データ無しは「―」）
  sh.getRange(row, col + 3, trend.length, 1)
    .setValues(trend.map(function (t) { return [fullTimeCell_(t)]; }))
    .setHorizontalAlignment('center');
  // 「今日のチェック」の帰宅時間欄に入力された、院を出た時刻
  sh.getRange(row, col + 4, trend.length, 1)
    .setValues(trend.map(function (t) { return [t.leaveAt || '―']; }))
    .setHorizontalAlignment('center');
}

/** 1ヶ月ぶんの推移から、その月のまとめを作る */
function monthSummary_(trend) {
  let days = 0, sum = 0, full = 0, leaveSum = 0, leaveDays = 0;
  (trend || []).forEach(function (t) {
    if (!t.has) return;                       // 記録の無い日は平均に入れない
    days++;
    sum += t.pct;
    if (t.pct >= 100) full++;
    const m = timeToMin_(t.leaveAt);
    if (m >= 0) { leaveSum += m; leaveDays++; }
  });
  return {
    days: days,                                        // 記録のあった日数
    avgPct: days ? Math.round(sum / days) : 0,         // 平均完了率
    fullDays: full,                                    // 100%達成の日数
    avgLeave: leaveDays ? minToTimeStr_(Math.round(leaveSum / leaveDays)) : '',
  };
}

/** まとめの1行（前月ぶんを渡すと「前月比」も付けます） */
function monthSummaryText_(sm, prev) {
  if (!sm.days) return '記録なし';
  let s = '平均完了率 ' + sm.avgPct + '%　／　100%達成 ' + sm.fullDays + '/' + sm.days + '日' +
          '　／　平均帰宅 ' + (sm.avgLeave || '―');
  if (prev && prev.days) {
    const d = sm.avgPct - prev.avgPct;
    s += '　／　前月比 ' + (d === 0 ? '±0' : (d > 0 ? '+' + d : String(d))) + 'pt';
  }
  return s;
}

/** 行が足りなければ追加する（既定のタブは1000行あるので通常は何もしません） */
function ensureRows_(sh, need) {
  const max = sh.getMaxRows();
  if (max >= need) return;
  sh.insertRowsAfter(max, need - max);
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
 * @param {Date=} base 対象月に含まれる任意の日（省略時は今日）
 * @param {Object=} grouped 読み込み済みの履歴（getLogsGroupedByDate_ の結果）。
 *                          省略時はこの中で履歴シートを読みます。
 *                          何ヶ月ぶんも作るときは、まとめて読んだものを渡してください。
 */
function getMonthTrend_(base, grouped) {
  base = base || new Date();
  const y = base.getFullYear(), m = base.getMonth();
  const now = new Date();
  const isCurrentMonth = (y === now.getFullYear() && m === now.getMonth());
  const lastDay = isCurrentMonth ? now.getDate() : daysInMonth_(y, m);

  if (!grouped) {
    grouped = getLogsGroupedByDate_(dateStr_(new Date(y, m, 1)),
                                    dateStr_(new Date(y, m, lastDay)));
  }
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
