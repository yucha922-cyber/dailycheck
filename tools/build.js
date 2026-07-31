/**
 * apps-script/*.gs を 1ファイルにまとめて dist/DailyCheck.gs を作ります。
 * Apps Script エディタに「1回コピペするだけ」で導入できるようにするためのものです。
 *
 *   node tools/build.js
 */
const fs = require('fs');
const path = require('path');

const ORDER = ['Config', 'Common', 'Settings', 'Master', 'History', 'Checklist',
               'Dashboard', 'CentralSync', 'Diagnostics', 'Menu', 'Setup', 'Triggers'];

const src = path.join(__dirname, '..', 'apps-script');
const out = path.join(__dirname, '..', 'dist', 'DailyCheck.gs');

const version = (fs.readFileSync(path.join(src, 'Config.gs'), 'utf8')
  .match(/APP_VERSION\s*=\s*'([^']+)'/) || [, '?'])[1];

const header = [
  '/**',
  ' * ============================================================',
  ' *  サロン日次業務チェック  ―  一括貼り付け用ファイル',
  ' *  version ' + version,
  ' * ------------------------------------------------------------',
  ' *  このファイル1つを Apps Script エディタの「コード.gs」に',
  ' *  そのまま貼り付ければ導入できます。',
  ' *  （apps-script/ の各ファイルを結合した自動生成ファイルです。',
  ' *    修正は apps-script/ 側で行い、node tools/build.js で再生成してください）',
  ' * ============================================================',
  ' */',
  '',
].join('\n');

const body = ORDER.map(function (n) {
  const code = fs.readFileSync(path.join(src, n + '.gs'), 'utf8').replace(/\s+$/, '');
  return '// ============================================================\n' +
         '// ' + n + '.gs\n' +
         '// ============================================================\n\n' + code;
}).join('\n\n');

fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, header + body + '\n');
console.log('生成しました: dist/DailyCheck.gs  (' + ORDER.length + 'ファイル / ' +
            (header + body).split('\n').length + '行)');
