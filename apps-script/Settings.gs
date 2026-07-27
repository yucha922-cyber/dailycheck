/**
 * ============================================================
 *  Settings.gs  ―  「設定」タブ
 * ------------------------------------------------------------
 *  店舗名・営業時間・通知先などをシート上で編集できるようにします。
 *  コードを触らずに運用を変えられるので、100店舗展開でも
 *  「同じコード＋各店の設定タブ」だけで済みます。
 * ============================================================
 */

let _cfgCache = null;

/** 設定を読み込む（1実行内はキャッシュ） */
function cfg_() {
  if (_cfgCache) return _cfgCache;
  const out = {};
  SETTING_DEFS.forEach(function (d) { out[d.key] = d.def; });

  const sh = settingsSheet_(false);
  if (sh && sh.getLastRow() >= 1) {
    const last = sh.getLastRow();
    const width = Math.max(sh.getLastColumn(), 2);
    const vals = sh.getRange(1, 1, last, width).getValues();
    const index = {};
    SETTING_DEFS.forEach(function (d) { index[norm_(d.label)] = d; });

    for (let r = 0; r < vals.length; r++) {
      const label = norm_(vals[r][0]);
      if (!label) continue;
      const d = index[label];
      if (!d) continue;
      let v = vals[r][1];
      if (v === '' || v === null) continue;
      out[d.key] = castSetting_(d, v);
    }
  }
  _cfgCache = out;
  return out;
}

/** 設定値を型に合わせて変換 */
function castSetting_(d, v) {
  if (typeof d.def === 'number') {
    const n = Number(String(v).replace(/[^\d.\-]/g, ''));
    return isNaN(n) ? d.def : n;
  }
  if (typeof d.def === 'boolean') return toBool_(v);
  if (d.key === 'OPEN_TIME' || d.key === 'CLOSE_TIME') return timeStr_(v);
  return String(v).trim();
}

/** キャッシュを捨てる（設定タブを編集したとき） */
function clearCfgCache_() { _cfgCache = null; }

/**
 * 「設定」タブを用意する。
 *  ・既にある項目は値を保持（絶対に上書きしない）
 *  ・足りない項目だけ追記する
 */
function setupSettingsSheet_() {
  const sh = settingsSheet_(true);
  const last = sh.getLastRow();

  // 既存の項目名を集める
  const existing = {};
  if (last >= 1) {
    const vals = sh.getRange(1, 1, last, 1).getValues();
    for (let r = 0; r < vals.length; r++) {
      const n = norm_(vals[r][0]);
      if (n) existing[n] = r + 1;
    }
  }

  // 見出しが無ければ付ける
  let startRow = last + 1;
  if (!existing[norm_('項目')] && last === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['項目', '値', '説明']])
      .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    startRow = 2;
  }

  // 足りない項目を追記
  const add = [];
  SETTING_DEFS.forEach(function (d) {
    if (existing[norm_(d.label)]) return;
    add.push([d.label, d.def, d.hint]);
  });
  if (add.length) {
    sh.getRange(startRow, 1, add.length, 3).setValues(add);
    // 真偽値の行はチェックボックスに
    for (let i = 0; i < add.length; i++) {
      if (typeof add[i][1] === 'boolean') sh.getRange(startRow + i, 2).insertCheckboxes();
    }
  }

  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 420);
  sh.getRange(1, 3, Math.max(sh.getLastRow(), 1), 1).setFontColor('#78909c').setFontSize(9);
  clearCfgCache_();
  return sh;
}
