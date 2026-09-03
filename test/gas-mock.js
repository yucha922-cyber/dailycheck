/* Google Apps Script の最小モック（テスト用） */

function pad(n, w) { return ('000' + n).slice(-w); }

const Utilities = {
  formatDate: function (d, tz, fmt) {
    const map = {
      'yyyy/MM/dd': pad(d.getFullYear(), 4) + '/' + pad(d.getMonth() + 1, 2) + '/' + pad(d.getDate(), 2),
      'HH:mm': pad(d.getHours(), 2) + ':' + pad(d.getMinutes(), 2),
      'MM/dd': pad(d.getMonth() + 1, 2) + '/' + pad(d.getDate(), 2),
    };
    if (!(fmt in map)) throw new Error('unsupported format ' + fmt);
    return map[fmt];
  },
};

const Logger = { log: function () {} };

class FakeSheet {
  constructor(name, maxRows, maxCols) {
    this.name = name;
    this.maxRows = maxRows || 200;
    this.maxCols = maxCols || 26;
    this.grid = [];
    this.notes = [];
    this.checkboxCells = {};
    for (let r = 0; r < this.maxRows; r++) {
      this.grid.push(new Array(this.maxCols).fill(''));
      this.notes.push(new Array(this.maxCols).fill(''));
    }
  }
  getName() { return this.name; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  getLastRow() {
    for (let r = this.maxRows - 1; r >= 0; r--) {
      for (let c = 0; c < this.maxCols; c++) if (this.grid[r][c] !== '' && this.grid[r][c] !== null) return r + 1;
    }
    return 0;
  }
  getLastColumn() {
    for (let c = this.maxCols - 1; c >= 0; c--) {
      for (let r = 0; r < this.maxRows; r++) if (this.grid[r][c] !== '' && this.grid[r][c] !== null) return c + 1;
    }
    return 0;
  }
  getRange(r, c, nr, nc) { return makeRange(this, r, c, nr === undefined ? 1 : nr, nc === undefined ? 1 : nc); }
  clear() {
    for (let r = 0; r < this.maxRows; r++)
      for (let c = 0; c < this.maxCols; c++) { this.grid[r][c] = ''; this.notes[r][c] = ''; }
    this.checkboxCells = {};
    return this;
  }
  insertRowsAfter(after, n) {
    this.maxRows += n;
    for (let i = 0; i < n; i++) {
      this.grid.push(new Array(this.maxCols).fill(''));
      this.notes.push(new Array(this.maxCols).fill(''));
    }
    return this;
  }
  insertColumnsAfter(after, n) {
    this.maxCols += n;
    for (let r = 0; r < this.maxRows; r++) {
      for (let i = 0; i < n; i++) { this.grid[r].push(''); this.notes[r].push(''); }
    }
    return this;
  }
  setFrozenRows() { return this; }
  setColumnWidth() { return this; }
  setColumnWidths() { return this; }
  setRowHeight() { return this; }
  activate() { return this; }
}

function makeRange(sheet, row, col, nr, nc) {
  const target = {
    getRow: () => row,
    getColumn: () => col,
    getNumRows: () => nr,
    getNumColumns: () => nc,
    getSheet: () => sheet,
    getValues: () => {
      const out = [];
      for (let r = 0; r < nr; r++) {
        const line = [];
        for (let c = 0; c < nc; c++) line.push(sheet.grid[row - 1 + r][col - 1 + c]);
        out.push(line);
      }
      return out;
    },
    getValue: () => sheet.grid[row - 1][col - 1],
    setValues: (vals) => {
      if (vals.length !== nr) throw new Error('row count mismatch: got ' + vals.length + ' want ' + nr);
      for (let r = 0; r < nr; r++) {
        if (vals[r].length !== nc) throw new Error('col count mismatch: got ' + vals[r].length + ' want ' + nc);
        for (let c = 0; c < nc; c++) sheet.grid[row - 1 + r][col - 1 + c] = vals[r][c];
      }
      return proxy;
    },
    setValue: (v) => { sheet.grid[row - 1][col - 1] = v; return proxy; },
    setNotes: (notes) => {
      for (let r = 0; r < nr; r++) for (let c = 0; c < nc; c++) sheet.notes[row - 1 + r][col - 1 + c] = notes[r][c];
      return proxy;
    },
    setNote: (note) => {
      for (let r = 0; r < nr; r++) for (let c = 0; c < nc; c++) sheet.notes[row - 1 + r][col - 1 + c] = note;
      return proxy;
    },
    clearNote: () => {
      for (let r = 0; r < nr; r++) for (let c = 0; c < nc; c++) sheet.notes[row - 1 + r][col - 1 + c] = '';
      return proxy;
    },
    getNote: () => sheet.notes[row - 1][col - 1],
    insertCheckboxes: () => {
      for (let r = 0; r < nr; r++) for (let c = 0; c < nc; c++) {
        const cell = sheet.grid[row - 1 + r][col - 1 + c];
        sheet.checkboxCells[(row + r) + ',' + (col + c)] = true;
        if (cell === '') sheet.grid[row - 1 + r][col - 1 + c] = false;
      }
      return proxy;
    },
    getCell: (r, c) => makeRange(sheet, row + r - 1, col + c - 1, 1, 1),
    merge: () => proxy,
    breakApart: () => proxy,
  };
  // 未実装のメソッド（書式系）は自分自身を返すだけ
  const proxy = new Proxy(target, {
    get(t, prop) {
      if (prop in t) return t[prop];
      if (typeof prop === 'string') return () => proxy;
      return undefined;
    },
  });
  return proxy;
}

class FakeSpreadsheet {
  constructor(sheets) { this.sheets = sheets || []; }
  getSheets() { return this.sheets; }
  getSheetByName(n) { return this.sheets.filter(s => s.getName() === n)[0] || null; }
  insertSheet(n) { const s = new FakeSheet(n); this.sheets.push(s); return s; }
  setActiveSheet(s) { return s; }
  moveActiveSheet() {}
  toast() {}
}

const SpreadsheetApp = {
  _ss: null,
  getActiveSpreadsheet() { return SpreadsheetApp._ss; },
  getActive() { return SpreadsheetApp._ss; },
  getUi() { throw new Error('no ui'); },
  newDataValidation() {
    const b = {
      requireValueInList: () => b, setAllowInvalid: () => b, setHelpText: () => b, build: () => ({}),
    };
    return b;
  },
  openById() { throw new Error('no access'); },
};

const ScriptApp = {
  getProjectTriggers: () => [],
  newTrigger: () => { throw new Error('not in test'); },
  deleteTrigger: () => {},
};

const MailApp = { sendEmail: () => {} };

module.exports = { FakeSheet, FakeSpreadsheet, SpreadsheetApp, Utilities, Logger, ScriptApp, MailApp };
