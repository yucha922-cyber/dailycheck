/**
 * ============================================================
 *  Common.gs  ―  共通ユーティリティ
 * ------------------------------------------------------------
 *  ここでの一番の役割は「既存のタブ・既存の列を壊さずに使う」こと。
 *   ・タブ …… 別名リストで探し、見つかればそれを使う（新規作成しない）
 *   ・列 …… 見出し文字で探し、無い列だけ右端に追加する
 * ============================================================
 */

const WEEKDAY_JP = ['日', '月', '火', '水', '木', '金', '土'];
const TZ = 'Asia/Tokyo';

/** スプレッドシート本体 */
function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/* ============================================================
 *  タブの解決
 * ========================================================== */

/** 比較用に正規化（全角/半角・空白・大文字小文字・記号ゆれを吸収） */
function norm_(s) {
  return String(s == null ? '' : s)
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[（）]/g, function (c) { return c === '（' ? '(' : ')'; })
    .replace(/[\s　_\-・]/g, '')
    .toLowerCase()
    .trim();
}

/**
 * 定義済みのシートを取得。
 * @param {string} defKey  SHEET_DEFS のキー（MASTER / CHECK / LOG / SETTINGS / DASH）
 * @param {boolean} create 見つからないとき作成するか
 * @return {Sheet|null}
 */
function sheetOf_(defKey, create) {
  const def = SHEET_DEFS[defKey];
  const ss = ss_();
  const sheets = ss.getSheets();

  // 1) 別名の順に、既存タブを探す（＝既にあるタブを最優先で使う）
  for (let a = 0; a < def.aliases.length; a++) {
    const target = norm_(def.aliases[a]);
    for (let i = 0; i < sheets.length; i++) {
      if (norm_(sheets[i].getName()) === target) return sheets[i];
    }
  }
  // 2) 部分一致（「今日のチェック（新）」のような名前も拾う）
  for (let a = 0; a < def.aliases.length; a++) {
    const target = norm_(def.aliases[a]);
    if (target.length < 3) continue;
    for (let i = 0; i < sheets.length; i++) {
      if (norm_(sheets[i].getName()).indexOf(target) >= 0) return sheets[i];
    }
  }
  if (!create) return null;
  return ss.insertSheet(def.canonical);
}

/** よく使うシートの取得ショートカット */
function masterSheet_(create)   { return sheetOf_('MASTER',   create !== false); }
function checkSheet_(create)    { return sheetOf_('CHECK',    create !== false); }
function logSheet_(create)      { return sheetOf_('LOG',      create !== false); }
function settingsSheet_(create) { return sheetOf_('SETTINGS', create !== false); }

/* ============================================================
 *  列の解決（見出しベース）
 * ========================================================== */

/**
 * 見出し行から、フィールド定義に対応する列番号を求める。
 * 見つからないフィールドは（createMissing=true なら）右端に列を追加。
 *
 * @param {Sheet} sh
 * @param {Array} fields   MASTER_FIELDS / LOG_FIELDS
 * @param {number} headerRow 見出しの行番号（既定1）
 * @param {boolean} createMissing 足りない列を追加するか
 * @return {Object} { key: 列番号(1始まり), ... }  ＋ _headerRow, _width
 */
function resolveColumns_(sh, fields, headerRow, createMissing) {
  headerRow = headerRow || 1;
  createMissing = createMissing !== false;

  let width = Math.max(sh.getLastColumn(), 1);
  let headers = sh.getRange(headerRow, 1, 1, width).getValues()[0];

  const map = { _headerRow: headerRow };
  const used = {};

  fields.forEach(function (f) {
    const cands = [f.header].concat(f.aliases || []).map(norm_);
    for (let c = 0; c < headers.length; c++) {
      const h = norm_(headers[c]);
      if (!h || used[c + 1]) continue;
      if (cands.indexOf(h) >= 0) {
        map[f.key] = c + 1;
        used[c + 1] = true;
        return;
      }
    }
  });

  // 完全一致で見つからなかったものは部分一致で再挑戦
  fields.forEach(function (f) {
    if (map[f.key]) return;
    const cands = [f.header].concat(f.aliases || []).map(norm_).filter(function (x) { return x.length >= 2; });
    for (let c = 0; c < headers.length; c++) {
      const h = norm_(headers[c]);
      if (!h || used[c + 1]) continue;
      for (let k = 0; k < cands.length; k++) {
        if (h === cands[k] || h.indexOf(cands[k]) >= 0 || cands[k].indexOf(h) >= 0) {
          map[f.key] = c + 1;
          used[c + 1] = true;
          return;
        }
      }
    }
  });

  // それでも無い列は右端に追加
  if (createMissing) {
    let next = width;
    // 末尾の空見出しを詰める
    while (next > 0 && !String(headers[next - 1] || '').trim()) next--;
    fields.forEach(function (f) {
      if (map[f.key]) return;
      next++;
      if (next > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), next - sh.getMaxColumns());
      sh.getRange(headerRow, next).setValue(f.header);
      map[f.key] = next;
      used[next] = true;
    });
    width = Math.max(width, next);
  }

  map._width = width;
  return map;
}

/* ============================================================
 *  日付・時刻
 * ========================================================== */

/** 日付セルかどうか（Sheetsが返す Date を判定） */
function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

function today_()          { return new Date(); }
function todayStr_()       { return dateStr_(new Date()); }
function dateStr_(d)       { return Utilities.formatDate(d, TZ, 'yyyy/MM/dd'); }
function nowTimeStr_()     { return Utilities.formatDate(new Date(), TZ, 'HH:mm'); }
function weekdayChar_(d)   { return WEEKDAY_JP[d.getDay()]; }
function dateLabel_(d)     { return dateStr_(d) + '(' + weekdayChar_(d) + ')'; }
function todayLabel_()     { return dateLabel_(new Date()); }

/** 'yyyy/MM/dd' 文字列 → Date（不正なら今日） */
function parseDate_(s) {
  if (isDate_(s)) return s;
  const m = String(s || '').match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (!m) return new Date();
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** セルの値を 'HH:mm' 表記に（Date でも文字列でもOK） */
function timeStr_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, TZ, 'HH:mm');
  return String(v == null ? '' : v).trim();
}

/** 'HH:mm'（'9:05' のような1桁も可）→ 0時からの分数。読めなければ -1 */
function timeToMin_(v) {
  const m = timeStr_(v).match(/(\d{1,2})\s*[:：]\s*(\d{1,2})/);
  if (!m) return -1;
  const h = Number(m[1]), mi = Number(m[2]);
  if (h > 23 || mi > 59) return -1;
  return h * 60 + mi;
}

/** 分数 → 'HH:mm' */
function minToTimeStr_(n) {
  if (n === null || n === undefined || n < 0) return '';
  const h = Math.floor(n / 60), m = Math.round(n % 60);
  return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
}

/** その月の日数 */
function daysInMonth_(y, m) { return new Date(y, m + 1, 0).getDate(); }

/**
 * 目安時間を「分」の数値にする。
 *  5        → 5
 *  '5分'     → 5
 *  '10 min' → 10
 *  '1時間30分' → 90
 *  '約5分程度' → 5
 *  時刻セル（0:05）→ 5
 *  空欄・読めない値 → 0
 */
function toMinutes_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  if (isDate_(v)) return v.getHours() * 60 + v.getMinutes();

  let s = String(v)
    .replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[．，,]/g, function (c) { return c === '．' ? '.' : ''; })
    .replace(/\s/g, '');

  let total = 0, hit = false;
  const h = s.match(/(\d+(?:\.\d+)?)(?:時間|hours?|hrs?|h)/i);
  if (h) { total += parseFloat(h[1]) * 60; hit = true; s = s.replace(h[0], ''); }
  const m = s.match(/(\d+(?:\.\d+)?)(?:分|minutes?|mins?|m)/i);
  if (m) { total += parseFloat(m[1]); hit = true; }
  if (hit) return Math.round(total);

  const n = parseFloat(s.replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

/** その日が月末か */
function isMonthEnd_(d) {
  const n = new Date(d.getFullYear(), d.getMonth() + 1, 0);
  return d.getDate() === n.getDate();
}

/** true/false のゆるい判定（チェックボックス・文字列どちらでも） */
function toBool_(v) {
  if (v === true) return true;
  if (v === false) return false;
  const s = String(v == null ? '' : v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === '○' || s === '◯' || s === '✓' ||
         s === 'yes' || s === 'y' || s === 'はい' || s === '必須' || s === '有効';
}

/* ============================================================
 *  頻度の判定
 * ========================================================== */

/**
 * その業務を「その日」に出すか判定する。
 *
 *  毎日 / 毎回 / 日次      → 毎日
 *  （空欄）                → 毎日（追加直後に消えないように）
 *  平日                    → 月〜金
 *  土日 / 週末             → 土・日
 *  月 火 水 木 金 土 日     → その曜日（「月,木」「月水金」「毎週火」も可）
 *  月末 / 末日             → 月末日
 *  毎月15日 / 15日         → 毎月その日（複数可「1日,15日」）
 *
 * @param {Object} task getMasterTasks_() の1件
 * @param {Date} date 判定する日
 */
function isScheduledOn_(task, date) {
  if (task.enabled === false) return false;
  const raw = String(task.freq == null ? '' : task.freq).trim();
  if (!raw) return true;                       // 未入力は「毎日」扱い
  if (/毎日|毎回|日次|随時/.test(raw)) return true;

  const dow = date.getDay();
  const dom = date.getDate();
  let s = raw;

  // 月末（「月」が曜日と誤判定されないよう先に処理）
  if (/月末|末日/.test(s)) {
    if (isMonthEnd_(date)) return true;
    s = s.replace(/月末|末日/g, '');
  }

  // 「15日」「毎月1日,15日」（曜日の「日」と区別するため数字付きのみ）
  const days = s.match(/\d{1,2}\s*日/g);
  if (days) {
    for (let i = 0; i < days.length; i++) {
      if (parseInt(days[i], 10) === dom) return true;
    }
    s = s.replace(/\d{1,2}\s*日/g, '');
  }

  // 「毎月」だけ書かれていたら 1日 とみなす
  if (/毎月/.test(s) && !days) {
    if (dom === 1) return true;
  }
  s = s.replace(/毎月|毎週|隔週/g, '');

  if (/平日/.test(s)) {
    if (dow >= 1 && dow <= 5) return true;
    s = s.replace(/平日/g, '');
  }
  if (/土日|週末/.test(s)) {
    if (dow === 0 || dow === 6) return true;
    s = s.replace(/土日|週末/g, '');
  }

  // 曜日
  return s.indexOf(WEEKDAY_JP[dow]) >= 0;
}

/** その日に実施すべき業務の一覧（区分順に並べる） */
function getScheduledTasks_(date) {
  date = date || new Date();
  return getMasterTasks_()
    .filter(function (t) { return t.name && isScheduledOn_(t, date); })
    .sort(function (a, b) {
      const pa = PHASE_ORDER.indexOf(a.phase), pb = PHASE_ORDER.indexOf(b.phase);
      const na = pa < 0 ? 99 : pa, nb = pb < 0 ? 99 : pb;
      if (na !== nb) return na - nb;
      return a.row - b.row;   // 同じ区分内はマスターの行順
    });
}
