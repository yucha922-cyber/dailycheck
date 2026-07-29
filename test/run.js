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

console.log('\n----------------------------------------');
console.log(failures === 0 ? `全 ${checks} 件パス` : `${failures} / ${checks} 件 失敗`);
process.exit(failures === 0 ? 0 : 1);
