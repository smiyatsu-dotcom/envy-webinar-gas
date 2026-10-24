// 単体テスト用の、GASの小さな偽物（スプレッドシート・ドライブ・プロパティ）。
// ★本物のGoogleには一切つながない。実在の案件（ホスピタソン社など）のIDも使わない。すべて偽のデータ。
// 書き込みは全部 log に積むので、「旧の流れの動きが変わっていないか」を、変更前のコードと突き合わせて比べられる。
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const root = path.join(__dirname, '..');
// The macOS shim named `git` can hang under the test executor. Prefer the
// already-installed Command Line Tools binary, with the ordinary PATH fallback
// retained for other environments.
const gitExecutable = fs.existsSync('/Library/Developer/CommandLineTools/usr/bin/git') ? '/Library/Developer/CommandLineTools/usr/bin/git' : 'git';

function a1ToRC(a1) {
  const m = String(a1).match(/^([A-Z]+)(\d+)$/);
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return [Number(m[2]), c];
}
function rcToA1(r, c) {
  let s = '';
  while (c > 0) { const k = (c - 1) % 26; s = String.fromCharCode(65 + k) + s; c = Math.floor((c - 1) / 26); }
  return s + r;
}
function norm(v) {
  if (v instanceof Date || (v && v.constructor && v.constructor.name === 'Date')) return 'DATE';
  return v;
}

function makeEnv() {
  const log = [];
  const props = {};
  let seq = 0;

  // ---------------- スプレッドシート ----------------
  function makeSheet(key, name) {
    const cells = {};    // 'r,c' -> {v, note, cb, f}
    const cell = (r, c) => (cells[r + ',' + c] = cells[r + ',' + c] || { v: '' });
    const sheet = {
      key, cells,
      getName: () => name,
      setName: (n) => { name = n; return sheet; },
      getRange(a, b, nr, nc) {
        let r0, c0, r1, c1;
        if (typeof a === 'string') {
          const parts = a.split(':');
          [r0, c0] = a1ToRC(parts[0]);
          [r1, c1] = parts[1] ? a1ToRC(parts[1]) : [r0, c0];
        } else { r0 = a; c0 = b; r1 = a + (nr || 1) - 1; c1 = b + (nc || 1) - 1; }
        const rng = {
          getValue: () => cell(r0, c0).v,
          setValue: (v) => { cell(r0, c0).v = v; log.push(['set', key, rcToA1(r0, c0), norm(v)]); return rng; },
          getValues: () => { const out = []; for (let r = r0; r <= r1; r++) { const row = []; for (let c = c0; c <= c1; c++) row.push(cell(r, c).v); out.push(row); } return out; },
          setValues: (vals) => { vals.forEach((row, i) => row.forEach((v, j) => { cell(r0 + i, c0 + j).v = v; })); log.push(['setValues', key, rcToA1(r0, c0), vals.map(r => r.map(norm))]); return rng; },
          getNote: () => cell(r0, c0).note || '',
          setNote: (n) => { cell(r0, c0).note = n; log.push(['note', key, rcToA1(r0, c0)]); return rng; },
          getDataValidation: () => (cell(r0, c0).cb ? {} : null),
          insertCheckboxes: () => { cell(r0, c0).cb = true; if (cell(r0, c0).v === '') cell(r0, c0).v = false; log.push(['cb', key, rcToA1(r0, c0)]); return rng; },
          getFormula: () => cell(r0, c0).f || '',
          setFormula: (f) => { cell(r0, c0).f = f; log.push(['formula', key, rcToA1(r0, c0), f]); return rng; },
          setWrap: () => rng, setFontWeight: () => rng, setBackground: () => rng, setFontColor: () => rng,
          setVerticalAlignment: () => rng, merge: () => rng
        };
        return rng;
      },
      getDataRange() {
        let mr = 0, mc = 0;
        Object.keys(cells).forEach(k => { const [r, c] = k.split(',').map(Number); if (cells[k].v !== '' && cells[k].v !== undefined) { mr = Math.max(mr, r); mc = Math.max(mc, c); } });
        return sheet.getRange(1, 1, Math.max(mr, 1), Math.max(mc, 1));
      },
      appendRow(vals) {
        let mr = 0;
        Object.keys(cells).forEach(k => { const r = Number(k.split(',')[0]); if (cells[k].v !== '') mr = Math.max(mr, r); });
        vals.forEach((v, j) => { cell(mr + 1, j + 1).v = v; });
        log.push(['append', key, vals.map(norm)]);
      },
      deleteRow(r) {
        const next = {};
        Object.keys(cells).forEach(k => { const [rr, c] = k.split(',').map(Number); if (rr < r) next[k] = cells[k]; else if (rr > r) next[(rr - 1) + ',' + c] = cells[k]; });
        Object.keys(cells).forEach(k => delete cells[k]); Object.assign(cells, next);
        log.push(['deleteRow', key, r]);
      },
      setRowHeight: () => sheet, setColumnWidth: () => sheet, setFrozenRows: () => sheet
    };
    return sheet;
  }
  const spreadsheets = {};
  function makeSpreadsheet(id, tabs) {
    const ss = { id, sheets: {}, getId: () => id, getUrl: () => 'https://docs.google.com/spreadsheets/d/' + id + '/edit' };
    tabs.forEach(t => { ss.sheets[t] = makeSheet(id + '/' + t, t); });
    ss.getSheets = () => Object.values(ss.sheets);
    ss.getSheetByName = (n) => ss.sheets[n] || null;
    spreadsheets[id] = ss;
    return ss;
  }
  const SpreadsheetApp = {
    openById: (id) => { if (!spreadsheets[id]) throw new Error('Exception: No item with the given ID could be found'); return spreadsheets[id]; }
  };

  // ---------------- ドライブ ----------------
  const files = {};
  function makeFile(name, content, mime, parent) {
    const id = 'F' + (++seq) + 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
    const f = {
      id, name, content: content || '', mime: mime || 'text/plain', parent, trashed: false, updated: new Date(Date.UTC(2026, 8, 26, 1, 0, seq)),
      getId: () => id, getName: () => f.name, getMimeType: () => f.mime, getUrl: () => 'https://drive/' + id, isTrashed: () => f.trashed,
      getLastUpdated: () => f.updated,
      getBlob: () => ({ getDataAsString: () => f.content }),
      setName: (n) => { log.push(['rename', f.parent.path + '/' + f.name, n]); f.name = n; return f; },
      setTrashed: (t) => { f.trashed = !!t; log.push(['trash', f.parent.path + '/' + f.name]); return f; },
      setContent: (c) => { f.content = c; log.push(['content', f.parent.path + '/' + f.name]); return f; },
      moveTo: (folder) => { f.parent = folder; return f; },
      makeCopy: (n, folder) => { log.push(['copy', f.parent.path + '/' + f.name, folder.path + '/' + n]); return folder.addFile(n, f.content, f.mime); }
    };
    files[id] = f;
    return f;
  }
  function iter(arr) { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; }
  const folders = {};
  function makeFolder(name, parent, id) {
    id = id || ('D' + (++seq) + 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxx');
    const d = {
      id, name, parent, path: (parent ? parent.path + '/' : '') + name, subs: [], files: [],
      getId: () => id, getName: () => name, getUrl: () => 'https://drive/folders/' + id,
      getParents: () => iter(parent ? [parent] : []),
      getFolders: () => iter(d.subs.slice()),
      getFiles: () => iter(d.files.filter(f => !f.trashed && f.parent === d)),
      getFoldersByName: (n) => iter(d.subs.filter(s => s.name === n)),
      getFilesByName: (n) => iter(d.files.filter(f => !f.trashed && f.parent === d && f.name === n)),
      createFolder: (n) => { log.push(['mkdir', d.path + '/' + n]); return d.addFolder(n); },
      createFile: (n, c, m) => { log.push(['create', d.path + '/' + n, String(c).slice(0, 80)]); return d.addFile(n, c, m); },
      addFolder: (n, fid) => { const s = makeFolder(n, d, fid); d.subs.push(s); return s; },
      addFile: (n, c, m) => { const f = makeFile(n, c, m, d); d.files.push(f); return f; }
    };
    folders[id] = d;
    return d;
  }
  const DriveApp = {
    getFolderById: (id) => { if (!folders[id]) throw new Error('Exception: No item with the given ID could be found'); return folders[id]; },
    getFileById: (id) => { if (!files[id]) throw new Error('Exception: No item with the given ID could be found'); return files[id]; },
    createFile: (n, c, m) => { const f = makeFile(n, c, m, { path: '(root)', files: [] }); return f; }
  };

  const PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = String(v); return this; },
      deleteProperty: (k) => { delete props[k]; },
      getProperties: () => Object.assign({}, props)
    })
  };
  const LockService = { getScriptLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {} }) };
  const MimeType = { GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet', GOOGLE_DOCS: 'application/vnd.google-apps.document', GOOGLE_SLIDES: 'application/vnd.google-apps.presentation', PLAIN_TEXT: 'text/plain' };
  const Logger = { log: () => {} };
  const ScriptApp = { getService: () => ({ getUrl: () => '' }), getProjectTriggers: () => [] };
  const Utilities = {
    DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
    computeDigest: (_algorithm, text) => [...crypto.createHash('sha256').update(String(text), 'utf8').digest()].map(n => n > 127 ? n - 256 : n)
  };

  return { log, props, spreadsheets, folders, files, makeSpreadsheet, makeFolder, globals: { SpreadsheetApp, DriveApp, PropertiesService, LockService, MimeType, Logger, ScriptApp, Utilities } };
}

/** webinar-gas の .js を読み込む。rev を渡すと、そのコミットの中身で読む（変更前との比べ用） */
function loadGas(env, rev) {
  const ctx = Object.assign({ console }, env ? env.globals : {});
  vm.createContext(ctx);
  let names;
  if (rev) {
    names = execFileSync(gitExecutable, ['-C', root, 'ls-tree', '--name-only', rev], { encoding: 'utf8' }).split('\n').filter(f => /\.js$/.test(f)).sort();
  } else {
    names = fs.readdirSync(root).filter(f => /\.js$/.test(f)).sort();
  }
  names.forEach(f => {
    const raw = rev ? execFileSync(gitExecutable, ['-C', root, 'show', rev + ':' + f], { encoding: 'utf8' }) : fs.readFileSync(path.join(root, f), 'utf8');
    vm.runInContext(raw.replace(/^const /gm, 'var '), ctx, { filename: (rev ? rev + ':' : '') + f });
  });
  return ctx;
}

module.exports = { makeEnv, loadGas, a1ToRC, rcToA1 };
