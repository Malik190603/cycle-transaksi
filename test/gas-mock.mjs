// Tiruan layanan Google Apps Script (SpreadsheetApp, CacheService, LockService, dst) di memori,
// cukup untuk menjalankan backend/src/*.gs di Node. Meniru perilaku Sheets yang penting:
// string angka menjadi angka dan string tanggal yyyy-MM-dd menjadi Date, kecuali kolom berformat teks.
import crypto from 'node:crypto';

export function createGas(now = () => new Date()) {
  const books = new Map();
  let seq = 0;
  const colIdx = (s) => s.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);

  class Sheet {
    constructor(book, name) { this.book = book; this.name = name; this.rows = []; this.textCols = new Set(); this.frozen = 0; }
    getName() { return this.name; }
    getParent() { return this.book; }
    getLastRow() { for (let i = this.rows.length - 1; i >= 0; i--) if ((this.rows[i] || []).some((v) => v !== '' && v != null)) return i + 1; return 0; }
    getLastColumn() { return this.rows.reduce((m, r) => { for (let i = (r || []).length - 1; i >= 0; i--) if (r[i] !== '' && r[i] != null) return Math.max(m, i + 1); return m; }, 0); }
    getMaxRows() { return Math.max(1000, this.rows.length); }
    getMaxColumns() { return 26; }
    setFrozenRows(n) { this.frozen = n; return this; }
    conv(v, c) {
      if (v === null || v === undefined) return '';
      if (typeof v !== 'string' || this.textCols.has(c)) return v;
      if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return new Date(v + 'T00:00:00+07:00');
      if (v.trim() !== '' && /^-?\d+(\.\d+)?$/.test(v.trim())) return Number(v);
      return v;
    }
    cell(r, c) { const row = this.rows[r - 1]; const v = row ? row[c - 1] : ''; return v === undefined || v === null ? '' : (v instanceof Date ? new Date(v.getTime()) : v); }
    put(r, c, v) { while (this.rows.length < r) this.rows.push([]); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = this.conv(v, c); }
    appendRow(vals) { const r = this.getLastRow() + 1; vals.forEach((v, i) => this.put(r, i + 1, v)); return this; }
    getRange(a, b, c, d) {
      if (typeof a === 'string') {
        let m = a.match(/^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/);
        if (!m) throw new Error('A1 tidak didukung tiruan: ' + a);
        const c1 = colIdx(m[1]), r1 = m[2] ? +m[2] : 1, c2 = m[3] ? colIdx(m[3]) : c1, open = !m[4] && (m[3] || !m[2]);
        const r2 = m[4] ? +m[4] : (open ? Math.max(this.getMaxRows(), r1) : r1);
        return new Range(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1, open);
      }
      if (!(a >= 1) || !(b >= 1)) throw new Error(`Range tidak valid (${a}, ${b}) di sheet ${this.name}`);
      const nr = c === undefined ? 1 : c, nc = d === undefined ? 1 : d;
      if (!(nr >= 1) || !(nc >= 1)) throw new Error(`The number of rows/columns in the range must be at least 1 (${nr}x${nc}) di sheet ${this.name}`);
      return new Range(this, a, b, nr, nc, false);
    }
    getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()), false); }
    deleteRow(r) { if (r >= 1 && r <= this.rows.length) this.rows.splice(r - 1, 1); return this; }
    deleteRows(r, n) { this.rows.splice(r - 1, n); return this; }
    insertRowBefore(r) { this.rows.splice(r - 1, 0, []); return this; }
    insertRowsAfter(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r, 0, []); return this; }
    clearContents() { this.rows = []; return this; }
    clear() { this.rows = []; return this; }
    autoResizeColumns() { return this; } setColumnWidth() { return this; } hideSheet() { return this; }
  }
  class Range {
    constructor(sheet, r, c, nr, nc, open) { Object.assign(this, { sheet, r, c, nr, nc, open }); }
    getSheet() { return this.sheet; }
    getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
    getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sheet.cell(this.r + i, this.c + j)); out.push(row); } return out; }
    getValue() { return this.sheet.cell(this.r, this.c); }
    getDisplayValues() { return this.getValues().map((r) => r.map(String)); }
    setValues(v) {
      if (!Array.isArray(v) || v.length !== this.nr || v.some((row) => !Array.isArray(row) || row.length !== this.nc)) throw new Error(`The number of rows/columns in the data does not match the range (${v && v.length}x${v && v[0] && v[0].length} vs ${this.nr}x${this.nc}) di sheet ${this.sheet.name}`);
      v.forEach((row, i) => row.forEach((x, j) => this.sheet.put(this.r + i, this.c + j, x))); return this;
    }
    setValue(x) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.put(this.r + i, this.c + j, x); return this; }
    clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) if (this.sheet.rows[this.r + i - 1]) this.sheet.put(this.r + i, this.c + j, ''); return this; }
    clear() { return this.clearContent(); }
    setNumberFormat(f) { if (f === '@' && this.open) for (let j = 0; j < this.nc; j++) this.sheet.textCols.add(this.c + j); return this; }
    setFontWeight() { return this; } setBackground() { return this; } setFontColor() { return this; } setHorizontalAlignment() { return this; } setWrap() { return this; } setNote() { return this; }
  }
  class Book {
    constructor(name) { this.id = 'SS' + (++seq) + crypto.randomBytes(6).toString('hex'); this.name = name; this.sheets = [new Sheet(this, 'Sheet1')]; books.set(this.id, this); }
    getId() { return this.id; } getName() { return this.name; } getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id + '/edit'; }
    getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
    insertSheet(n) { if (this.getSheetByName(n)) throw new Error('Sheet sudah ada: ' + n); const s = new Sheet(this, n); this.sheets.push(s); return s; }
    getSheets() { return this.sheets.slice(); }
    deleteSheet(s) { this.sheets = this.sheets.filter((x) => x !== s); }
    rename(n) { this.name = n; }
  }
  const active = new Book('Cycle Transaksi - DB');
  const cache = new Map();
  const cacheApi = {
    get(k) { const e = cache.get(k); if (!e) return null; if (e.exp < now().getTime()) { cache.delete(k); return null; } return e.v; },
    put(k, v, ttl) { cache.set(k, { v: String(v), exp: now().getTime() + (ttl || 600) * 1000 }); },
    remove(k) { cache.delete(k); }, removeAll(ks) { ks.forEach((k) => cache.delete(k)); },
    getAll(ks) { const o = {}; ks.forEach((k) => { const v = cacheApi.get(k); if (v !== null) o[k] = v; }); return o; },
    putAll(o, ttl) { Object.keys(o).forEach((k) => cacheApi.put(k, o[k], ttl)); }
  };
  const props = new Map();
  const triggers = [];
  const fmt = (date, tz, f) => {
    const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', { timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }).formatToParts(date).map((x) => [x.type, x.value]));
    return f.replace(/yyyy|MM|dd|HH|mm|ss/g, (t) => ({ yyyy: p.year, MM: p.month, dd: p.day, HH: p.hour, mm: p.minute, ss: p.second }[t]));
  };
  const text = (s) => ({ content: s, mime: null, setMimeType(m) { this.mime = m; return this; }, getContent() { return this.content; } });
  const html = (s) => ({ content: s, getContent() { return this.content; }, setTitle() { return this; }, addMetaTag() { return this; }, setXFrameOptionsMode() { return this; }, evaluate() { return this; } });
  return {
    books, active, cache, triggers,
    globals: {
      SpreadsheetApp: {
        getActiveSpreadsheet: () => active,
        openById: (id) => { const b = books.get(id); if (!b) throw new Error('Spreadsheet tidak ditemukan: ' + id); return b; },
        create: (n) => new Book(n), flush() {}, getUi() { throw new Error('Cannot call SpreadsheetApp.getUi() from this context.'); }
      },
      CacheService: { getScriptCache: () => cacheApi },
      LockService: { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {}, hasLock: () => true }) },
      PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (props.has(k) ? props.get(k) : null), setProperty: (k, v) => props.set(k, String(v)), deleteProperty: (k) => props.delete(k) }) },
      Utilities: { formatDate: fmt, getUuid: () => crypto.randomUUID(), sleep() {} },
      ScriptApp: {
        getProjectTriggers: () => triggers.slice(), deleteTrigger: (t) => { const i = triggers.indexOf(t); if (i >= 0) triggers.splice(i, 1); },
        newTrigger: (fn) => { const t = { fn, getHandlerFunction: () => fn }; const b = { timeBased: () => b, everyMinutes: (n) => { t.every = n; return b; }, create: () => { triggers.push(t); return t; } }; return b; }
      },
      ContentService: { createTextOutput: text, MimeType: { JSON: 'application/json' } },
      HtmlService: { createHtmlOutput: html, createTemplateFromFile: () => html(''), createHtmlOutputFromFile: () => html(''), XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' } },
      Logger: { log() {} }, Session: { getActiveUser: () => ({ getEmail: () => '' }) }, console
    }
  };
}
