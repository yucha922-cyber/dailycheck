/**
 * ============================================================
 *  CentralSync.gs  ―  本部（多店舗）集約
 * ------------------------------------------------------------
 *  各店舗のスプレッドシートから、本部の集約スプレッドシートへ
 *  「その日1行のサマリ」を送ります。
 *
 *  前提：
 *   ・「設定」タブの『本部集約スプレッドシートID』を入力
 *   ・そのスプレッドシートに、この Googleアカウントが編集権限を持つこと
 *   ・集約側に「店舗別サマリ」シートが無ければ自動作成します
 * ============================================================
 */

const CENTRAL_SHEET = '店舗別サマリ';
const CENTRAL_HEADER = ['日付', '店舗ID', '店舗名', '完了率', '完了数', '総数', '未完了数', '未完了業務', '更新時刻'];

/** 指定日のサマリを本部へ upsert 送信 */
function pushDailySummaryToCentral_(date) {
  const c = cfg_();
  if (!c.CENTRAL_SS_ID) return;

  let central;
  try {
    central = SpreadsheetApp.openById(c.CENTRAL_SS_ID);
  } catch (err) {
    try { ss_().toast('本部スプレッドシートを開けません。IDと共有権限を確認してください。', 'エラー', 8); } catch (e2) {}
    return;
  }

  let sh = central.getSheetByName(CENTRAL_SHEET);
  if (!sh) {
    sh = central.insertSheet(CENTRAL_SHEET);
    sh.getRange(1, 1, 1, CENTRAL_HEADER.length).setValues([CENTRAL_HEADER])
      .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }

  const s = getDailyStats_(date);
  const row = [
    date, c.STORE_ID, c.STORE_NAME, s.pct + '%',
    s.done, s.total, s.notDone.length, s.notDone.join(' / '), nowTimeStr_(),
  ];

  const last = sh.getLastRow();
  if (last >= 2) {
    const keys = sh.getRange(2, 1, last - 1, 2).getValues();
    for (let i = 0; i < keys.length; i++) {
      const d = isDate_(keys[i][0]) ? dateStr_(keys[i][0]) : String(keys[i][0]).trim();
      if (d === String(date) && String(keys[i][1]).trim() === String(c.STORE_ID)) {
        sh.getRange(i + 2, 1, 1, CENTRAL_HEADER.length).setValues([row]);
        return;
      }
    }
  }
  sh.getRange(last + 1, 1, 1, CENTRAL_HEADER.length).setValues([row]);
}
