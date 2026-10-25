// Fixed No.40 map-only repair and current-J read audit.  This test uses only
// fake Drive/Slides objects; it cannot contact or modify Google.
const assert = require('assert');
const { makeEnv, loadGas } = require('./fake_gas.cjs');

const CASE = '隔離通し診断テスト社';
const HEARING = '1ur8hPUElZMrKuitnDo1j1wzr54EkmRY9efNMaEQFErs';
const FOLDER = '1Ko7cbJ2nW2C4uQbZfDnZjW6mfz_qmHZ0';
const DECK = '1wLXXeU95iculuFJ_NOQ5oSZ3PSSWXLRa_L6m2GuX_vM';
const MAP = '【スライド対応】隔離通し診断テスト社.json';
const GEOMETRY = '【配置指定】隔離通し診断テスト社.json';
const SCRIPT = '【台本JSON】隔離通し診断テスト社.txt';
const TEXT = '【台本】隔離通し診断テスト社_ヒアリングシート.txt';
const IMPORT_RECORD = '【不確定】J台本取込記録.json';

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('ok  ' + name); }
  catch (error) { console.log('NG  ' + name + '\n    ' + (error && error.stack || error)); process.exitCode = 1; }
}
function iterator(items) { let index = 0; return { hasNext: () => index < items.length, next: () => items[index++] }; }
function rows(G) {
  return Array.from({ length: 106 }, (_, index) => {
    const no = index + 1;
    return { no, daihon: G.FIXED_DAIHON[no] || 'note-' + no };
  });
}
function shape(id, text, geometry) {
  let currentText = text;
  return {
    getObjectId: () => id, getPageElementType: () => 'SHAPE', getTitle: () => '', getDescription: () => '',
    getLeft: () => geometry.left, getTop: () => geometry.top, getWidth: () => geometry.width, getHeight: () => geometry.height,
    asShape: () => ({ getText: () => ({ asString: () => currentText }) }),
    setText: value => { currentText = value; }
  };
}
function wordArt(id, text) {
  let currentText = text;
  return {
    getObjectId: () => id, getPageElementType: () => 'WORD_ART',
    asWordArt: () => ({ getRenderedText: () => currentText }),
    setText: value => { currentText = value; }
  };
}
function slide(id, note, elements) {
  let currentNote = note;
  const currentElements = elements || [];
  return {
    elements: currentElements,
    setNote: value => { currentNote = value; },
    getObjectId: () => id,
    getPageElements: () => currentElements,
    getNotesPage: () => ({ getSpeakerNotesShape: () => ({ getText: () => ({ asString: () => currentNote }) }) })
  };
}
function fixture() {
  const env = makeEnv();
  const G = loadGas(env);
  env.props.INPUT_FOLDER_ID = 'normal-input';
  env.props.OUTPUT_FOLDER_ID = 'normal-output';
  env.props.TEMPLATE_SLIDE_ID = 'template';
  env.props.ASSETS_FOLDER_ID = 'assets';
  env.props.MGMT_SHEET_ID = 'normal-ledger';
  const parent = env.makeFolder('【隔離試験】配置永続化', null, '1GUmxo4J9cj0WYdh4IPj2lBi3RSxiMw73');
  const root = parent.addFolder('【隔離通し】診断から資料音声_20261023', '1oQSEwyopPcvX20xuxNDHy5Kh4TfvtCBO');
  root.addFolder('【出力】', '1Mxch8d4SEbddMfx-C2slbmQWWa1Uz16u');
  const output = root.addFolder('【出力】生成物', '18-xhLrskU8M9DvW2Rurqg5ST_NrOiSM2');
  root.addFolder('【素材】投影資料', 'isolation-assets');
  const folder = output.addFolder(CASE, FOLDER);
  const ledger = env.makeSpreadsheet('16UgS3scd3X2D2Qqhzd1WfJuGkuNcOt5WbekQZLyKDpg', ['台帳']);
  ledger.getSheets()[0].appendRow(['ヒアリングファイルID', '案件名', 'ステータス', '構成シートURL', 'スライドURL', 'メッセージ', '更新日時', '流れ']);
  ledger.getSheets()[0].appendRow([HEARING, CASE + '_ヒアリングシート', '完了', '', '', '', '', '新']);
  const before = { left: 40.44094488188976, top: 37.21555118110236, width: 639.1181102362204, height: 47.40944881889764 };
  const after = { ...before, top: 40 };
  const pages = Array.from({ length: 106 }, (_, index) => ({ slideNo: index + 1, pageId: 'page-' + (index + 1) }));
  const map = { schema: 'consultation-slide-map/v2', caseName: CASE, presentationId: DECK, presentationRevision: 'test', pages };
  const geometry = { schema: 'consultation-slide-geometry/v2', caseName: CASE, records: [{ status: 'active', slideNo: 40, selector: { type: 'SHAPE', title: '', description: '', templateGeometry: before }, after }] };
  const scriptRows = rows(G);
  const scriptText = JSON.stringify(scriptRows);
  const notes = scriptRows.map(row => row.daihon);
  const mapFile = folder.addFile(MAP, JSON.stringify(map));
  folder.addFile(GEOMETRY, JSON.stringify(geometry));
  folder.addFile(SCRIPT, scriptText);
  folder.addFile(TEXT, notes.join('\n\n') + '\n');
  const oldGetFile = G.DriveApp.getFileById;
  G.DriveApp.getFileById = id => {
    if (id === ledger.getId()) return { getParents: () => iterator([root]) };
    if (id === DECK) return { getMimeType: () => G.MimeType.GOOGLE_SLIDES, isTrashed: () => false, getParents: () => iterator([folder]) };
    return oldGetFile(id);
  };
  G.隔離通し_J台本_SHA256 = G.隔離通し_J台本_sha256_(scriptText);
  folder.addFile(IMPORT_RECORD, JSON.stringify({ status: 'completed', automaticRetry: false, folderId: FOLDER, targetScriptSha256: G.隔離通し_J台本_SHA256 }));
  G.隔離通し_J_No40配置対応修復_.mapSha256 = G.隔離通し_J_No40配置対応修復_sha256_(mapFile.content);
  G.隔離通し_J_No40配置対応修復_.geometrySha256 = G.隔離通し_J_No40配置対応修復_sha256_(JSON.stringify(geometry));
  const slides = pages.map((page, index) => slide(page.pageId, notes[index], index === 39 ? [shape('durable-heading', '永続配置の見出し', after)] : []));
  G.SlidesApp = { openById: id => { assert.strictEqual(id, DECK); return { getSlides: () => slides, saveAndClose: () => { throw new Error('read/repair must not save deck'); } }; } };
  G.processImports_ = () => { throw new Error('repair must not import'); };
  G.processNewFlow_ = () => { throw new Error('repair must not advance flow'); };
  G.スライドを組む_ = () => { throw new Error('repair must not regenerate'); };
  return { env, G, folder, mapFile, slides };
}

test('fixed No.7 is a natural participant action without emoji or production metadata', () => {
  const { G } = fixture();
  assert.strictEqual(G.FIXED_DAIHON[7], '例えば、皆さん経営していてこんなお悩みはありませんか？\n当てはまる方は、挙手マークを押してください。');
  assert.doesNotMatch(G.FIXED_DAIHON[7], /🖐️|リアクション|挙手ありがとうございます|避け/u);
});

test('current J audit is read-only and reports the missing No.40 selector without using the historical deck audit', () => {
  const { env, G } = fixture();
  const before = env.log.length;
  const result = G.隔離通し_J現行資料を読取検品();
  assert.strictEqual(result.ok, false);
  assert.strictEqual(result.mode, 'read-only-current-j-strict-audit');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.scriptNotes)), {
    expectedRows: 106, comparedRows: 106, mismatchCount: 0, mismatchNos: [],
    emptyNoteCount: 0, emptyNoteNos: [], unreadableNoteCount: 0, unreadableNoteNos: []
  });
  assert.strictEqual(result.fixedScript.expectedCount, 9);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.fixedScript.fixedNos)), [2, 7, 33, 38, 49, 97, 101, 105, 106]);
  assert.match(result.fixedScript.baselineSha256, /^[a-f0-9]{64}$/);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.fixedScript.scriptMismatchNos)), []);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.fixedScript.noteMismatchNos)), []);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.zoomOccurrences)), { notes: 0, visibleSlides: 0, visibleTextReadErrors: 0 });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.geometry)), [{ slideNo: 40, mapSelector: false, selectorAfter: false, currentDeckMatches: 0, retained: false }]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.actions)), { imported: false, regenerated: false, mapRepaired: false, delivered: false, notified: false, audioStarted: false });
  assert.strictEqual(env.log.length, before);
});

test('strict audit compares all 106 JSON rows to their mapped notes without removing wording, and reports only safe evidence', () => {
  const { env, G, slides } = fixture();
  // Line-ending style alone is intentionally tolerated.
  slides[1].setNote(G.FIXED_DAIHON[2].replace(/\n/g, '\r\n'));
  // Empty and altered notes remain mismatches; no filtering or narration-text
  // removal can turn either into a pass.
  slides[4].setNote('');
  slides[5].setNote('note-6（挙手ありがとうございます。）');
  slides[7].setNote('different-8 Zoom');
  slides[8].setNote('LABEL: note-9');
  slides[8].elements.push(wordArt('visible-zoom', 'ZOOM'));
  const before = env.log.length;
  const result = G.隔離通し_J現行資料を読取検品();
  assert.strictEqual(result.ok, false);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.scriptNotes)), {
    expectedRows: 106, comparedRows: 106, mismatchCount: 4, mismatchNos: [5, 6, 8, 9],
    emptyNoteCount: 1, emptyNoteNos: [5], unreadableNoteCount: 0, unreadableNoteNos: []
  });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.zoomOccurrences)), { notes: 1, visibleSlides: 1, visibleTextReadErrors: 0 });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.noteDifferenceSummary)), {
    'exact-match': 101, 'line-ending-only-match': 1, 'empty-note': 1,
    'note-is-prefix-of-expected': 0, 'note-has-extra-prefix': 1, 'note-has-extra-suffix': 1,
    'content-difference': 1, 'unreadable-note': 0
  });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.noteDifferences)), [
    { no: 5, matches: false, category: 'empty-note', actualLength: 0, expectedLength: 6, firstDifferentCharacter: 1, actualExcerpt: '', expectedExcerpt: 'note-5' },
    { no: 6, matches: false, category: 'note-has-extra-suffix', actualLength: 21, expectedLength: 6, firstDifferentCharacter: 7, actualExcerpt: 'note-6（挙手ありがとうございます。）', expectedExcerpt: 'note-6' },
    { no: 8, matches: false, category: 'content-difference', actualLength: 16, expectedLength: 6, firstDifferentCharacter: 1, actualExcerpt: 'different-8 Zoom', expectedExcerpt: 'note-8' },
    { no: 9, matches: false, category: 'note-has-extra-prefix', actualLength: 13, expectedLength: 6, firstDifferentCharacter: 1, actualExcerpt: 'LABEL: note-9', expectedExcerpt: 'note-9' }
  ]);
  assert.strictEqual(result.normalGeneratorExpectation.order[0], 'resolveDaihon_-fixed-before-imported-JSON');
  assert.ok(result.noteDifferences.every(item => String(item.actualExcerpt || '').length <= 32 && String(item.expectedExcerpt || '').length <= 32));
  const serialized = JSON.stringify(result);
  assert.doesNotMatch(serialized, /visible-zoom/);
  assert.match(result.normalizedNotesSha256, /^[a-f0-9]{64}$/);
  assert.match(result.visibleSlideTextSha256, /^[a-f0-9]{64}$/);
  assert.strictEqual(env.log.length, before);
});

test('strict audit checks each fixed row against both the JSON and current note baseline', () => {
  const { G } = fixture();
  G.FIXED_DAIHON[2] = 'changed fixed baseline';
  const result = G.隔離通し_J現行資料を読取検品();
  assert.strictEqual(result.ok, false);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.fixedScript.scriptMismatchNos)), [2]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.fixedScript.noteMismatchNos)), [2]);
});

test('strict audit uses normal-generator fixed precedence for notes while separately reporting raw JSON divergence', () => {
  const { G, folder } = fixture();
  const scriptFile = folder.getFilesByName(SCRIPT).next();
  const rows = JSON.parse(scriptFile.content);
  rows.find(row => row.no === 105).daihon = 'raw-105-different';
  rows.find(row => row.no === 106).daihon = 'raw-106-different';
  scriptFile.content = JSON.stringify(rows);
  G.隔離通し_J台本_SHA256 = G.隔離通し_J台本_sha256_(scriptFile.content);
  const record = folder.getFilesByName(IMPORT_RECORD).next();
  record.content = JSON.stringify({ status: 'completed', automaticRetry: false, folderId: FOLDER, targetScriptSha256: G.隔離通し_J台本_SHA256 });
  const result = G.隔離通し_J現行資料を読取検品();
  assert.strictEqual(result.scriptNotes.mismatchCount, 0);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.normalGeneratorExpectation.fixedSourceDivergenceNos)), [105, 106]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.fixedScript.scriptMismatchNos)), [105, 106]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.fixedScript.noteMismatchNos)), []);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.noteDifferences)), []);
});

test('No.40 map repair preserves exact input, resolves only the durable after-state heading, and fails closed on retry', () => {
  const { env, G, folder, mapFile } = fixture();
  const result = G.隔離通し_J_No40配置対応を一度修復する();
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.mode, 'fixed-map-only-repair');
  const repaired = JSON.parse(mapFile.content);
  assert.deepStrictEqual(repaired.pages[39].targetSelector, {
    type: 'SHAPE', title: '', description: '', templateGeometry: { left: 40.44094488188976, top: 40, width: 639.1181102362204, height: 47.40944881889764 }, role: 'heading', text: '永続配置の見出し'
  });
  assert.strictEqual(folder.getFilesByName('【保全】J_No40配置対応修復前.json').next().content.includes('"slideNo":40'), true);
  assert.strictEqual(JSON.parse(folder.getFilesByName('【不確定】J_No40配置対応修復記録.json').next().content).status, 'completed');
  const auditLogLength = env.log.length;
  const audit = G.隔離通し_J現行資料を読取検品();
  assert.strictEqual(audit.ok, true);
  assert.strictEqual(audit.geometry[0].retained, true);
  assert.strictEqual(env.log.length, auditLogLength);
  const retryLogLength = env.log.length;
  assert.throws(() => G.隔離通し_J_No40配置対応を一度修復する(), /自動再実行せず/);
  assert.strictEqual(env.log.length, retryLogLength);
});

test('repair source cannot call import, flow, or slide regeneration', () => {
  const source = require('fs').readFileSync(require('path').join(__dirname, '..', '56_geometry_isolation.js'), 'utf8');
  const body = source.slice(source.indexOf('function 隔離通し_J_No40配置対応を一度修復する'), source.indexOf('function runIsolationJNo40MapRepair'));
  assert.doesNotMatch(body, /processImports_\s*\(/);
  assert.doesNotMatch(body, /processNewFlow_\s*\(/);
  assert.doesNotMatch(body, /スライドを組む_\s*\(/);
  assert.doesNotMatch(body, /saveAndClose\s*\(/);
});

test('strict audit source has no mutation, import, regeneration, delivery, or notification path', () => {
  const source = require('fs').readFileSync(require('path').join(__dirname, '..', '58_isolated_case_audit.js'), 'utf8');
  const body = source.slice(source.indexOf('function 隔離通し_J現行資料を読取検品'), source.indexOf('function runIsolationCurrentJReadAudit'));
  assert.doesNotMatch(body, /setContent\s*\(/);
  assert.doesNotMatch(body, /createFile\s*\(/);
  assert.doesNotMatch(body, /processImports_\s*\(/);
  assert.doesNotMatch(body, /processNewFlow_\s*\(/);
  assert.doesNotMatch(body, /スライドを組む_\s*\(/);
  assert.doesNotMatch(body, /saveAndClose\s*\(/);
  assert.doesNotMatch(body, /notify\w*\s*\(/i);
});

console.log('pass ' + passed);
