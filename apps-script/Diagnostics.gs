/**
 * ============================================================
 *  Diagnostics.gs  ―  動作チェック（診断）
 * ------------------------------------------------------------
 *  「マスターを直したのに今日のチェックが変わらない」を
 *  自分で確認できるようにするための点検ツールです。
 *  メニュー「🩺 動作チェック（診断）」から実行します。
 * ============================================================
 */

function runDiagnostics() {
  clearCfgCache_();
  const lines = [];
  const ng = [];

  lines.push('■ バージョン：' + APP_VERSION);
  lines.push('■ 今日：' + todayLabel_());
  lines.push('');

  // --- タブの認識状況 ---
  lines.push('■ 使用中のタブ');
  const bind = [
    ['業務マスター', masterSheet_(false)],
    ['今日のチェック', checkSheet_(false)],
    ['履歴', logSheet_(false)],
    ['設定', settingsSheet_(false)],
  ];
  bind.forEach(function (b) {
    if (b[1]) lines.push('　✅ ' + b[0] + ' → 「' + b[1].getName() + '」');
    else { lines.push('　❌ ' + b[0] + ' → 見つかりません'); ng.push(b[0] + 'タブがありません（🏗 初期セットアップを実行）'); }
  });
  lines.push('');

  // --- マスターの読み取り ---
  const master = masterSheet_(false);
  if (master) {
    const col = resolveColumns_(master, MASTER_FIELDS, 1, false);
    const found = [];
    const missing = [];
    MASTER_FIELDS.forEach(function (f) {
      if (col[f.key]) found.push(f.header + '=' + colLetter_(col[f.key]) + '列');
      else missing.push(f.header);
    });
    lines.push('■ マスターの列の認識');
    lines.push('　' + found.join(' / '));
    if (missing.length) lines.push('　⚠ 未検出：' + missing.join('、') + '（初期セットアップで自動追加されます）');

    const tasks = getMasterTasks_();
    const todays = getScheduledTasks_(new Date());
    lines.push('　登録業務：' + tasks.length + '件　／　有効：' +
               tasks.filter(function (t) { return t.enabled; }).length + '件');
    lines.push('　今日の対象：' + todays.length + '件');
    if (!tasks.length) ng.push('マスターから業務を読み取れません（「業務名」の列見出しを確認）');
    if (tasks.length && !todays.length) ng.push('今日の対象が0件です（「頻度」「有効」を確認）');
    lines.push('');
  }

  // --- 今日のチェックの状態 ---
  const check = checkSheet_(false);
  if (check) {
    const ds = getCheckDate_(check);
    const rows = Math.max(check.getLastRow() - CHK_DATA_START + 1, 0);
    lines.push('■ 今日のチェック');
    lines.push('　表示中の日付：' + (ds || '（未生成）'));
    lines.push('　行数：' + rows + '件');
    if (ds && ds !== todayStr_()) ng.push('表示中の日付が今日ではありません（🔄 マスターの内容を反映 を実行）');
    lines.push('');
  }

  // --- 突き合わせ：マスターの今日の対象 vs 画面 ---
  if (master && check) {
    const want = getScheduledTasks_(new Date()).map(function (t) { return t.id; });
    const have = Object.keys(readCheckState_(check));
    const missingOnCheck = want.filter(function (id) { return have.indexOf(id) < 0; });
    const extraOnCheck = have.filter(function (id) { return want.indexOf(id) < 0; });
    lines.push('■ マスターと画面の一致');
    if (!missingOnCheck.length && !extraOnCheck.length) {
      lines.push('　✅ 一致しています');
    } else {
      if (missingOnCheck.length) lines.push('　⚠ 画面に出ていない業務：' + missingOnCheck.join(', '));
      if (extraOnCheck.length) lines.push('　⚠ マスターに無い業務が画面に：' + extraOnCheck.join(', '));
      ng.push('マスターと画面がずれています（🔄 マスターの内容を反映 を実行）');
    }
    lines.push('');
  }

  // --- トリガー ---
  lines.push('■ 自動化トリガー');
  const handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  const need = [
    ['onEditInstallable', 'マスター編集の即反映／✓の自動記録'],
    ['dailyMorningJob', '毎朝の当日分の用意'],
    ['nightlyCloseJob', '毎晩の確定'],
  ];
  need.forEach(function (t) {
    if (handlers.indexOf(t[0]) >= 0) lines.push('　✅ ' + t[1]);
    else { lines.push('　❌ ' + t[1] + ' 未設定'); ng.push('「⏰ 自動化トリガーを設定」が未実行です'); }
  });
  lines.push('');

  // --- 設定 ---
  const c = cfg_();
  lines.push('■ 設定');
  lines.push('　店舗：' + c.STORE_NAME + '（' + c.STORE_ID + '）　営業 ' + c.OPEN_TIME + '〜' + c.CLOSE_TIME);
  lines.push('　未完了メール：' + (c.ALERT_EMAIL || '（なし）'));
  lines.push('　本部集約ID：' + (c.CENTRAL_SS_ID ? '設定あり' : '（なし）'));
  lines.push('　ダッシュボード：' + (c.USE_DASHBOARD ? '使う' : '使わない'));
  lines.push('');

  // --- 履歴 ---
  const todayLogs = getLogByDate_(todayStr_());
  lines.push('■ 履歴');
  lines.push('　今日の記録：' + todayLogs.length + '件（完了 ' +
             todayLogs.filter(function (l) { return l.done; }).length + '件）');
  lines.push('');

  lines.push(ng.length ? '▼ 対応が必要です\n・' + uniq_(ng).join('\n・')
                       : '🎉 すべて正常です。マスターを編集すると今日のチェックに即反映されます。');

  const text = lines.join('\n');
  Logger.log(text);
  alert_('動作チェック結果', text);
  return text;
}

function uniq_(arr) {
  const seen = {}, out = [];
  arr.forEach(function (x) { if (!seen[x]) { seen[x] = true; out.push(x); } });
  return out;
}

/** 列番号 → A1 の列文字 */
function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}
