/**
 * ============================================================
 *  CentralSync.gs  ―  本部（100店舗）集約
 * ------------------------------------------------------------
 *  各店舗のスプレッドシートから、本部の集約スプレッドシートへ
 *  「その日1行のサマリ」を送信します。
 *
 *  前提：
 *   ・本部の集約スプレッドシートID を Config.gs の CENTRAL_SS_ID に設定
 *   ・そのスプレッドシートに、この Googleアカウントが編集権限を持つこと
 *   ・集約側に「店舗別サマリ」シートが無ければ自動作成します
 *
 *  100店舗の全体像は docs/02_マルチ店舗展開.md を参照。
 * ============================================================
 */

const CENTRAL_SHEET = '店舗別サマリ';
const CENTRAL_COLS = 9; // 日付,店舗ID,店舗名,完了率,完了数,総数,未完了数,未完了業務,更新時刻

/** 指定日のサマリを本部へ upsert 送信 */
function pushDailySummaryToCentral_(date) {
  if (!CONFIG.CENTRAL_SS_ID) return;
  const s = getDailyStats_(date);
  let central;
  try {
    central = SpreadsheetApp.openById(CONFIG.CENTRAL_SS_ID);
  } catch (err) {
    ss_().toast('本部スプレッドシートを開けません。IDと共有権限を確認してください。', 'エラー', 8);
    return;
  }

  let sh = central.getSheetByName(CENTRAL_SHEET);
  if (!sh) {
    sh = central.insertSheet(CENTRAL_SHEET);
    sh.getRange(1, 1, 1, CENTRAL_COLS).setValues([[
      '日付', '店舗ID', '店舗名', '完了率', '完了数', '総数', '未完了数', '未完了業務', '更新時刻',
    ]]).setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }

  const row = [
    date, CONFIG.STORE_ID, CONFIG.STORE_NAME, s.pct + '%',
    s.done, s.total, s.notDone.length, s.notDone.join(' / '), nowTimeStr_(),
  ];

  const last = sh.getLastRow();
  if (last >= 2) {
    const keys = sh.getRange(2, 1, last - 1, 2).getValues(); // 日付, 店舗ID
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(date) && String(keys[i][1]) === String(CONFIG.STORE_ID)) {
        sh.getRange(i + 2, 1, 1, CENTRAL_COLS).setValues([row]);
        return;
      }
    }
  }
  sh.getRange(last + 1, 1, 1, CENTRAL_COLS).setValues([row]);
}
