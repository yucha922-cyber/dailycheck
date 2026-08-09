/* GASコードを実際に走らせて挙動を検証する */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const mock = require('./gas-mock.js');

const SRC = path.join(__dirname, '..', 'apps-script');
const ORDER = ['Config', 'Common', 'Settings', 'Master', 'History', 'Checklist', 'Dashboard', 'CentralSync', 'Diagnostics', 'Menu', 'Setup', 'Triggers'];
const code = ORDER.map(n => fs.readFileSync(path.join(SRC, n + '.gs'), 'utf8')).join('\n');

const ctx = {
  Utilities: mock.Utilities, Logger: mock.Logger, SpreadsheetApp: mock.SpreadsheetApp,
  ScriptApp: mock.ScriptApp, MailApp: mock.MailApp, console,
};
vm.createContext(ctx);
vm.runInContext(code, ctx, { filename: 'bundle.gs' });

let failures = 0, checks = 0;
function ok(cond, label, extra) {
  checks++;
  if (cond) { console.log('  ✓ ' + label); }
  else { failures++; console.log('  ✗ ' + label + (extra !== undefined ? '  → ' + JSON.stringify(extra) : '')); }
}
function section(t) { console.log('\n== ' + t); }

/* ---------- 1. 頻度判定 ---------- */
section('頻度の判定');
const mon = new Date(2026, 6, 27);   // 2026/07/27 月曜
const sat = new Date(2026, 7, 1);    // 2026/08/01 土曜
const eom = new Date(2026, 6, 31);   // 7/31 金曜・月末
const t = (freq, enabled) => ({ freq, enabled: enabled !== false, name: 'x' });

ok(ctx.isScheduledOn_(t('毎日'), mon) === true, '毎日 → 月曜に出る');
ok(ctx.isScheduledOn_(t('毎回'), sat) === true, '毎回 → 土曜に出る');
ok(ctx.isScheduledOn_(t(''), mon) === true, '空欄 → 毎日扱い');
ok(ctx.isScheduledOn_(t('月'), mon) === true, '月 → 月曜に出る');
ok(ctx.isScheduledOn_(t('月'), sat) === false, '月 → 土曜には出ない');
ok(ctx.isScheduledOn_(t('月,木'), mon) === true, '月,木 → 月曜に出る');
ok(ctx.isScheduledOn_(t('火水'), mon) === false, '火水 → 月曜には出ない');
ok(ctx.isScheduledOn_(t('平日'), mon) === true, '平日 → 月曜に出る');
ok(ctx.isScheduledOn_(t('平日'), sat) === false, '平日 → 土曜には出ない');
ok(ctx.isScheduledOn_(t('土日'), sat) === true, '土日 → 土曜に出る');
ok(ctx.isScheduledOn_(t('月末'), eom) === true, '月末 → 7/31に出る');
ok(ctx.isScheduledOn_(t('月末'), mon) === false, '月末 → 7/27(月)には出ない ※「月」と誤判定しない');
ok(ctx.isScheduledOn_(t('毎月27日'), mon) === true, '毎月27日 → 27日に出る');
ok(ctx.isScheduledOn_(t('毎月15日'), mon) === false, '毎月15日 → 27日には出ない');
ok(ctx.isScheduledOn_(t('日'), new Date(2026, 6, 26)) === true, '日 → 日曜に出る');
ok(ctx.isScheduledOn_(t('日'), mon) === false, '日 → 月曜には出ない');
ok(ctx.isScheduledOn_(t('毎日', false), mon) === false, '有効OFF → 出ない');

/* ---------- 2. 見出しゆれの吸収 ---------- */
section('シート名・列見出しのゆれ吸収');
ok(ctx.norm_('各業務マスター') === ctx.norm_('各業務マスター '), '前後の空白を無視');
ok(ctx.norm_('目安（分）') === ctx.norm_('目安(分)'), '全角カッコ＝半角カッコ');

/* ---------- 3. 実シート想定のシナリオ ---------- */
section('ユーザーの既存タブ（各業務マスター/今日のチェック/履歴/設定）で動くか');

// ユーザーのシートを模した状態：タブは既にあるが中身は空
const sheets = ['各業務マスター', '今日のチェック', '履歴', '設定'].map(n => new mock.FakeSheet(n));
mock.SpreadsheetApp._ss = new mock.FakeSpreadsheet(sheets);

ctx.setupAll();

const ss = mock.SpreadsheetApp._ss;
ok(ss.getSheets().length === 4, '新しいタブを増やしていない（4タブのまま）', ss.getSheets().map(s => s.getName()));
ok(ss.getSheetByName('業務マスタ') === null, '「業務マスタ」等の重複タブを作っていない');

const master = ss.getSheetByName('各業務マスター');
const check = ss.getSheetByName('今日のチェック');
const log = ss.getSheetByName('履歴');
const settings = ss.getSheetByName('設定');

ok(master.getLastRow() === 14, 'マスターに初期業務13件が入った', master.getLastRow());
ok(String(check.getRange(1, 1).getValue()).indexOf(ctx.todayStr_()) >= 0, '今日のチェックに今日の日付が入る');
ok(settings.getLastRow() >= 10, '設定タブに項目が並ぶ', settings.getLastRow());

const rowsOf = sh => Math.max(sh.getLastRow() - 4, 0);
const before = rowsOf(check);
ok(before > 0, '今日のチェックに業務が並んだ（' + before + '件）');

/* --- マスターに業務を追加 → 反映されるか --- */
section('マスターに行を足す → 今日のチェックに出るか');
const mLast = master.getLastRow();
master.getRange(mLast + 1, 2, 1, 2).setValues([['開店前', 'タオル補充']]);   // 区分と業務名だけ入力
ctx.handleEdit_({ range: master.getRange(mLast + 1, 3, 1, 1) });

const names = () => check.getRange(5, 4, Math.max(check.getLastRow() - 4, 1), 1).getValues().map(r => String(r[0]));
ok(names().some(n => n.indexOf('タオル補充') >= 0), '追加した業務が今日のチェックに出た', names());
ok(String(master.getRange(mLast + 1, 1).getValue()).length > 0, 'IDが自動採番された', master.getRange(mLast + 1, 1).getValue());
ok(String(master.getRange(mLast + 1, 5).getValue()) === '毎日', '頻度の空欄が「毎日」で補完された');
ok(master.getRange(mLast + 1, 8).getValue() === true, '有効が自動でON');

/* --- チェックを付ける → 履歴に記録されるか --- */
section('✓ を付ける → 履歴に記録されるか');
const targetRow = 5;
check.getRange(targetRow, 1).setValue(true);
ctx.handleEdit_({ range: check.getRange(targetRow, 1, 1, 1) });
ok(String(check.getRange(targetRow, 6).getValue()).match(/^\d{2}:\d{2}$/) !== null, '完了時刻が自動で入る', check.getRange(targetRow, 6).getValue());
ok(log.getLastRow() >= 2, '履歴に行が追加された', log.getLastRow());
const logTaskId = String(log.getRange(2, 4).getValue());
ok(logTaskId === String(check.getRange(targetRow, 2).getValue()), '履歴の業務IDが一致');
ok(String(log.getRange(2, 7).getValue()) === 'TRUE', '履歴の完了が TRUE');
ok(String(check.getRange(2, 1).getValue()).indexOf('完了 1 /') >= 0, '進捗表示が更新された', check.getRange(2, 1).getValue());

/* --- マスターを再編集 → ✓ が消えないか --- */
section('マスター編集後も ✓ が保持されるか');
master.getRange(mLast + 2, 2, 1, 2).setValues([['閉店前', '戸締まり確認']]);
ctx.handleEdit_({ range: master.getRange(mLast + 2, 3, 1, 1) });
const doneCol = check.getRange(5, 1, Math.max(check.getLastRow() - 4, 1), 1).getValues().map(r => r[0]);
ok(doneCol.filter(v => v === true).length === 1, '付けた✓が再生成後も残っている', doneCol);
ok(names().some(n => n.indexOf('戸締まり確認') >= 0), '追加した業務も出ている');
ok(log.getRange(2, 7).getValue() === 'TRUE', '履歴が二重に増えていない（upsert）');

/* --- 頻度を変える → 消えるか --- */
section('頻度を「日」にする → 月曜なら消えるか');
const isSunday = new Date().getDay() === 0;
master.getRange(mLast + 1, 5).setValue(isSunday ? '月' : '日');
ctx.handleEdit_({ range: master.getRange(mLast + 1, 5, 1, 1) });
ok(!names().some(n => n.indexOf('タオル補充') >= 0), '対象外の曜日にすると今日のチェックから消える', names());

/* --- 有効OFF --- */
section('有効のチェックを外す → 消えるか');
master.getRange(mLast + 2, 8).setValue(false);
ctx.handleEdit_({ range: master.getRange(mLast + 2, 8, 1, 1) });
ok(!names().some(n => n.indexOf('戸締まり確認') >= 0), '有効OFFで今日のチェックから消える');

/* --- 設定タブ --- */
section('設定タブの変更が反映されるか');
let storeRow = 0;
const sv = settings.getRange(1, 1, settings.getLastRow(), 2).getValues();
for (let i = 0; i < sv.length; i++) if (String(sv[i][0]) === '店舗名') storeRow = i + 1;
ok(storeRow > 0, '設定タブに「店舗名」の行がある');
settings.getRange(storeRow, 2).setValue('渋谷店');
ctx.handleEdit_({ range: settings.getRange(storeRow, 2, 1, 1) });
ok(String(check.getRange(1, 1).getValue()).indexOf('渋谷店') >= 0, '店舗名の変更がチェック画面に反映', check.getRange(1, 1).getValue());

/* --- 確定 --- */
section('確定して履歴に保存');
ctx.finalizeToday();
const logRows = log.getLastRow() - 1;
ok(logRows === Math.max(check.getLastRow() - 4, 0), '今日の全業務が履歴に入る（未完了も含む）', [logRows, check.getLastRow() - 4]);

/* --- 診断 --- */
section('診断が動くか');
const rep = ctx.runDiagnostics();
ok(rep.indexOf('各業務マスター') >= 0, '診断が使用中タブ名を報告する');
ok(rep.indexOf('❌') >= 0, '未設定のトリガーを検出できる（テスト環境ではトリガー無し）');

/* ---------- 4. 既存の別名タブ・別名見出しでも動くか ---------- */
section('タブ名・列見出しが違う既存シートでも動くか');
const s2 = [new mock.FakeSheet('業務マスター'), new mock.FakeSheet('本日のチェック'), new mock.FakeSheet('日次ログ')];
mock.SpreadsheetApp._ss = new mock.FakeSpreadsheet(s2);
ctx.clearCfgCache_();
// 独自の見出し・独自データを先に入れておく
s2[0].getRange(1, 1, 1, 4).setValues([['カテゴリ', 'タスク名', '所要時間', '実施頻度']]);
s2[0].getRange(2, 1, 2, 4).setValues([['開店前', '換気', 3, '毎日'], ['閉店前', 'ゴミ出し', 5, '毎日']]);
ctx.setupAll();
const ss2 = mock.SpreadsheetApp._ss;
ok(ss2.getSheetByName('各業務マスター') === null, '別名タブがあれば新規タブを作らない', ss2.getSheets().map(s => s.getName()));
const chk2 = ss2.getSheetByName('本日のチェック');
const n2 = chk2.getRange(5, 4, Math.max(chk2.getLastRow() - 4, 1), 1).getValues().map(r => String(r[0]));
ok(n2.some(x => x.indexOf('換気') >= 0) && n2.some(x => x.indexOf('ゴミ出し') >= 0), '独自見出しの既存データを読めた', n2);
ok(String(s2[0].getRange(1, 1).getValue()) === 'カテゴリ', '既存の見出しを書き換えていない');
ok(String(s2[0].getRange(2, 2).getValue()) === '換気', '既存データを書き換えていない');
const hdr2 = s2[0].getRange(1, 1, 1, s2[0].getLastColumn()).getValues()[0];
ok(hdr2.indexOf('ID') >= 0 && hdr2.indexOf('有効') >= 0, '足りない列（ID/有効）は右端に追加された', hdr2);

/* ---------- 5. 必須/有効が ○× などの文字で運用されている場合 ---------- */
section('必須・有効が文字（○×）の既存マスターでも壊さないか');
const s3 = [new mock.FakeSheet('各業務マスター'), new mock.FakeSheet('今日のチェック')];
mock.SpreadsheetApp._ss = new mock.FakeSpreadsheet(s3);
ctx.clearCfgCache_();
s3[0].getRange(1, 1, 1, 5).setValues([['区分', '業務名', '頻度', '必須', '有効']]);
s3[0].getRange(2, 1, 3, 5).setValues([
  ['開店前', '換気', '毎日', '○', '○'],
  ['営業中', '発注', '毎日', '×', '×'],
  ['閉店前', '施錠', '毎日', '○', ''],     // 有効が空欄
]);
ctx.setupAll();
ok(String(s3[0].getRange(2, 4).getValue()) === '○', '○ をチェックボックスに書き換えていない', s3[0].getRange(2, 4).getValue());
ok(String(s3[0].getRange(3, 5).getValue()) === '×', '× をそのまま残している');
const chk3 = ss3names();
function ss3names() {
  const sh = mock.SpreadsheetApp._ss.getSheetByName('今日のチェック');
  return sh.getRange(5, 4, Math.max(sh.getLastRow() - 4, 1), 1).getValues().map(r => String(r[0]));
}
ok(chk3.some(x => x.indexOf('換気') >= 0), '有効=○ の業務は出る', chk3);
ok(!chk3.some(x => x.indexOf('発注') >= 0), '有効=× の業務は出ない');
ok(chk3.some(x => x.indexOf('施錠') >= 0), '有効が空欄の業務は出る（追加直後に消えない）');
ok(chk3.some(x => x.indexOf('★') >= 0), '必須=○ が ★ として表示される', chk3);

/* ---------- 6. 目安時間（「5分」などの文字）の取り込み ---------- */
section('目安時間を分に変換できるか');
ok(ctx.toMinutes_(5) === 5, '数値 5 → 5');
ok(ctx.toMinutes_('5分') === 5, '「5分」→ 5');
ok(ctx.toMinutes_('10分') === 10, '「10分」→ 10');
ok(ctx.toMinutes_('１５分') === 15, '全角「１５分」→ 15');
ok(ctx.toMinutes_('1時間30分') === 90, '「1時間30分」→ 90');
ok(ctx.toMinutes_('約5分程度') === 5, '「約5分程度」→ 5');
ok(ctx.toMinutes_('10 min') === 10, '「10 min」→ 10');
ok(ctx.toMinutes_('') === 0, '空欄 → 0');
ok(ctx.toMinutes_('なし') === 0, '読めない値 → 0');
ok(ctx.toMinutes_(new Date(2026, 0, 1, 0, 5)) === 5, '時刻セル 0:05 → 5');

section('実シートと同じ列構成（ID/区分/業務名/期限/頻度/担当/目安時間/…）で反映されるか');
const s4 = [new mock.FakeSheet('各業務マスター'), new mock.FakeSheet('今日のチェック'),
            new mock.FakeSheet('履歴'), new mock.FakeSheet('設定')];
mock.SpreadsheetApp._ss = new mock.FakeSpreadsheet(s4);
ctx.clearCfgCache_();
s4[0].getRange(1, 1, 1, 13).setValues([[
  'ID', '区分', '業務名', '期限', '頻度', '担当', '目安時間', '優先度', '自動化', '有効', 'カテゴリ', '必須', '備考']]);
s4[0].getRange(2, 1, 3, 13).setValues([
  [1, '開店前', '清掃',       '10:40', '毎日', '担当者', '5分',  '高', '×', true, '', false, ''],
  [2, '営業中', 'カルテ記入', '施術後', '毎回', '担当者', '15分', '高', '△', true, '', false, ''],
  [3, '閉店前', '売上確認',   '20:02', '毎日', '院長',   '2分',  '高', '○', true, '', false, ''],
]);
ctx.setupAll();
const chk4 = mock.SpreadsheetApp._ss.getSheetByName('今日のチェック');
const mins = chk4.getRange(5, 5, 3, 1).getValues().map(r => r[0]);
ok(JSON.stringify(mins) === JSON.stringify([5, 15, 2]), 'E列(目安)に 5 / 15 / 2 が入る', mins);
ok(String(s4[0].getRange(2, 7).getValue()) === '5分', 'マスター側の「5分」は書き換えない');
ok(String(s4[0].getRange(2, 4).getValue()) === '10:40', '期限など未使用の列は触らない');
const est = ctx.estimateMinutes_();
ok(est.before === 5 && est.after === 2, '想定作業時間の合計も正しく計算される', est);

/* ---------- 7. 時刻ユーティリティ ---------- */
section('時刻の変換');
ok(ctx.timeToMin_('20:15') === 1215, '「20:15」→ 1215分');
ok(ctx.timeToMin_('9:05') === 545, '1桁の「9:05」も読める');
ok(ctx.timeToMin_(new Date(2026, 0, 1, 7, 30)) === 450, '時刻セルも読める');
ok(ctx.timeToMin_('') === -1, '空欄 → -1');
ok(ctx.timeToMin_('あとで') === -1, '読めない値 → -1');
ok(ctx.timeToMin_('99:99') === -1, 'あり得ない時刻 → -1');
ok(ctx.minToTimeStr_(545) === '09:05', '545分 → 「09:05」（0埋め）');
ok(ctx.logDateStr_('2026/8/1') === '2026/08/01', '履歴の「2026/8/1」を 0埋めに揃える');
ok(ctx.logDateStr_(new Date(2026, 7, 1)) === '2026/08/01', '日付セルも 0埋めに揃える');

/* ---------- 8. ダッシュボードの1ヶ月間の完了率 ---------- */
section('ダッシュボード：1ヶ月間の完了率と 100%達成時刻（H列）');
const s5 = ['各業務マスター', '今日のチェック', '履歴', '設定'].map(n => new mock.FakeSheet(n));
mock.SpreadsheetApp._ss = new mock.FakeSpreadsheet(s5);
ctx.clearCfgCache_();
ctx.setupAll();

const now = new Date();
const dayOf = n => ctx.dateStr_(new Date(now.getFullYear(), now.getMonth(), n));

// 1日＝全業務完了（最後の完了は 20:15）／2日＝半分だけ完了
ctx.writeHistoryRows_(dayOf(1), [
  { taskId: 'T1', name: '清掃', phase: '開店前', done: true, time: '10:40', min: 10 },
  { taskId: 'T2', name: '売上確認', phase: '閉店前', done: true, time: '20:15', min: 5 },
  { taskId: 'T3', name: '現金確認', phase: '閉店前', done: true, time: '19:50', min: 3 },
]);
ctx.writeHistoryRows_(dayOf(2), [
  { taskId: 'T1', name: '清掃', phase: '開店前', done: true, time: '10:30', min: 10 },
  { taskId: 'T2', name: '売上確認', phase: '閉店前', done: false, time: '', min: 5 },
]);

const trend = ctx.getMonthTrend_();
ok(trend.length === now.getDate(), '当月1日〜今日の日数ぶんの行が出る（' + trend.length + '行）', trend.length);
ok(trend[0].label.indexOf(('0' + (now.getMonth() + 1)).slice(-2) + '/01') === 0, '先頭は当月1日', trend[0].label);
const todayLabel = ('0' + (now.getMonth() + 1)).slice(-2) + '/' + ('0' + now.getDate()).slice(-2);
ok(trend[trend.length - 1].label.indexOf(todayLabel) === 0, '最後は今日（未来の日付は出さない）', trend[trend.length - 1].label);

const past = now.getDate() >= 3;   // 1〜2日はまだ「今日」なので集計方法が変わる
if (past) {
  ok(trend[0].pct === 100, '1日は完了率 100%', trend[0].pct);
  ok(trend[0].doneAt === '20:15', '1日の 100%達成時刻は最後の完了時刻 20:15', trend[0].doneAt);
  ok(ctx.fullTimeCell_(trend[0]) === '20:15', 'H列に 20:15 が出る');
  ok(trend[1].pct === 50, '2日は完了率 50%', trend[1].pct);
  ok(ctx.fullTimeCell_(trend[1]) === '―', '未達成の日の H列は「―」', ctx.fullTimeCell_(trend[1]));
} else {
  console.log('  （今日が月初のため 1日・2日の判定はスキップ）');
}
ok(ctx.fullTimeCell_({ has: false, pct: 0, doneAt: '' }) === '―', '記録の無い日の H列は「―」');
ok(ctx.fullTimeCell_({ has: true, pct: 100, doneAt: '' }) === '時刻なし', '全完了でも時刻未記録なら「時刻なし」');

ctx.buildDashboard_();
const dash = mock.SpreadsheetApp._ss.getSheetByName('ダッシュボード');
ok(dash !== null, 'ダッシュボードタブが作られる');
const dashTitle = String(dash.getRange(16, 5).getValue());
ok(dashTitle.indexOf((now.getMonth() + 1) + '月の完了率') >= 0, '見出しが「◯月の完了率（1ヶ月間）」になっている', dashTitle);
ok(String(dash.getRange(17, 8).getValue()) === '100%達成時刻', 'H列の見出しが「100%達成時刻」', dash.getRange(17, 8).getValue());
const hCol = dash.getRange(18, 8, trend.length, 1).getValues().map(r => String(r[0]));
ok(hCol.length === trend.length && hCol.every(v => v !== ''), 'H列が日数ぶん埋まっている', hCol.slice(0, 3));
if (past) ok(hCol[0] === '20:15', 'H列の1日目が 20:15', hCol[0]);
ok(String(dash.getRange(18, 5).getValue()) === trend[0].label, 'データは18行目から（1日）', dash.getRange(18, 5).getValue());

/* ---------- 9. 帰宅時間の欄 ---------- */
section('今日のチェックの「帰宅時間」欄（ID 14）');
const s6 = ['各業務マスター', '今日のチェック', '履歴', '設定'].map(n => new mock.FakeSheet(n));
mock.SpreadsheetApp._ss = new mock.FakeSpreadsheet(s6);
ctx.clearCfgCache_();
ctx.setupAll();

const chk6 = mock.SpreadsheetApp._ss.getSheetByName('今日のチェック');
const log6 = mock.SpreadsheetApp._ss.getSheetByName('履歴');
const lastRow6 = chk6.getLastRow();
ok(String(chk6.getRange(lastRow6, 2).getValue()) === '14', '最終行の ID が 14', chk6.getRange(lastRow6, 2).getValue());
ok(String(chk6.getRange(lastRow6, 4).getValue()) === '帰宅時間', '最終行の業務名が「帰宅時間」', chk6.getRange(lastRow6, 4).getValue());
ok(chk6.getRange(lastRow6, 1).getValue() === '', '帰宅時間の行にチェックボックスは無い');
ok(chk6.checkboxCells[lastRow6 + ',1'] !== true, '完了列のチェックボックス対象外', chk6.checkboxCells[lastRow6 + ',1']);
ok(String(chk6.getRange(lastRow6, 4).getNote()).indexOf('院を出た時刻') >= 0, '入力方法のメモが付く');

const taskCount6 = lastRow6 - 5;   // 帰宅時間の行を除いた業務数
const status6 = String(chk6.getRange(2, 1).getValue());
ok(status6.indexOf('完了 0 / ' + taskCount6 + '（') >= 0, '帰宅時間は完了率の母数に入らない（' + taskCount6 + '件）', status6);

// 帰宅時間を入力 → 履歴に残り、完了率は変わらない
chk6.getRange(lastRow6, 6).setValue('20:30');
ctx.handleEdit_({ range: chk6.getRange(lastRow6, 6, 1, 1) });
ok(String(chk6.getRange(lastRow6, 6).getValue()) === '20:30', '入力した帰宅時間が消えない');
ok(String(chk6.getRange(2, 1).getValue()).indexOf('完了 0 / ' + taskCount6 + '（') >= 0, '帰宅時間を入れても完了率は変わらない');
ok(String(chk6.getRange(2, 1).getValue()).indexOf('帰宅 20:30') >= 0, '上部に帰宅時間が表示される', chk6.getRange(2, 1).getValue());

const logIds6 = log6.getRange(2, 4, Math.max(log6.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]));
ok(logIds6.indexOf('14') >= 0, '履歴に帰宅時間（ID 14）が記録される', logIds6);
const stats6 = ctx.getDailyStats_(ctx.todayStr_());
ok(stats6.total === taskCount6, '完了統計の母数に帰宅時間が入らない', stats6.total);
ok(stats6.leaveAt === '20:30', '完了統計から帰宅時間が取れる', stats6.leaveAt);
ok(!stats6.notDone.some(x => x.indexOf('帰宅時間') >= 0), '未完了リストに帰宅時間が出ない', stats6.notDone);

// 業務の ✓ を付けても帰宅時間は消えない
chk6.getRange(5, 1).setValue(true);
ctx.handleEdit_({ range: chk6.getRange(5, 1, chk6.getLastRow() - 4, 1) });
ok(String(chk6.getRange(lastRow6, 6).getValue()) === '20:30', '✓の一括編集でも帰宅時間が消えない', chk6.getRange(lastRow6, 6).getValue());

// 確定 → 全業務が 100% になった日として扱えるか
section('全業務完了の日：100%達成時刻と帰宅時間が別々に出るか');
const rows6 = chk6.getLastRow() - 5;
for (let i = 0; i < rows6; i++) {
  chk6.getRange(5 + i, 1).setValue(true);
  chk6.getRange(5 + i, 6).setValue('1' + ('0' + (i % 10)).slice(-1) + ':00');
}
chk6.getRange(9, 6).setValue('21:45');    // いちばん遅い完了
ctx.handleEdit_({ range: chk6.getRange(5, 1, rows6, 6) });
const stats7 = ctx.getDailyStats_(ctx.todayStr_());
ok(stats7.pct === 100, '全業務完了で 100%', stats7.pct);
ok(stats7.doneAt === '21:45', '100%達成時刻は最後の完了時刻', stats7.doneAt);
ok(stats7.leaveAt === '20:30', '帰宅時間は 100%達成時刻とは別に保持される', stats7.leaveAt);

ctx.buildDashboard_();
const dash2 = mock.SpreadsheetApp._ss.getSheetByName('ダッシュボード');
ok(String(dash2.getRange(17, 9).getValue()) === '帰宅時間', 'I列の見出しが「帰宅時間」', dash2.getRange(17, 9).getValue());
const todayRow = 18 + new Date().getDate() - 1;
ok(String(dash2.getRange(todayRow, 8).getValue()) === '21:45', 'H列（今日）に 100%達成時刻', dash2.getRange(todayRow, 8).getValue());
ok(String(dash2.getRange(todayRow, 9).getValue()) === '20:30', 'I列（今日）に帰宅時間', dash2.getRange(todayRow, 9).getValue());
ok(String(dash2.getRange(18, 9).getValue()) === '―', '記録の無い日の I列は「―」', dash2.getRange(18, 9).getValue());

section('診断が帰宅時間の行を「マスターに無い業務」と誤検出しないか');
const rep2 = ctx.runDiagnostics();
ok(rep2.indexOf('マスターに無い業務が画面に') < 0, '帰宅時間の行は突き合わせの対象外', rep2.split('\n').filter(l => l.indexOf('マスターに無い') >= 0));

console.log('\n----------------------------------------');
console.log(failures === 0 ? `全 ${checks} 件パス` : `${failures} / ${checks} 件 失敗`);
process.exit(failures === 0 ? 0 : 1);
