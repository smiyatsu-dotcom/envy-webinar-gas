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
function rows() { return JSON.stringify(Array.from({ length: 106 }, (_, index) => ({ no: index + 1, daihon: 'note-' + (index + 1) }))); }
function shape(id, text, geometry) {
  return {
    getObjectId: () => id, getPageElementType: () => 'SHAPE', getTitle: () => '', getDescription: () => '',
    getLeft: () => geometry.left, getTop: () => geometry.top, getWidth: () => geometry.width, getHeight: () => geometry.height,
    asShape: () => ({ getText: () => ({ asString: () => text }) })
  };
}
function slide(id, note, elements) {
  return { getObjectId: () => id, getPageElements: () => elements || [], getNotesPage: () => ({ getSpeakerNotesShape: () => ({ getText: () => ({ asString: () => note }) }) }) };
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
  const scriptText = rows();
  const notes = Array.from({ length: 106 }, (_, index) => 'note-' + (index + 1));
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
  return { env, G, folder, mapFile };
}

test('current J audit is read-only and reports the missing No.40 selector without using the historical deck audit', () => {
  const { env, G } = fixture();
  const before = env.log.length;
  const result = G.隔離通し_J現行資料を読取検品();
  assert.strictEqual(result.ok, false);
  assert.strictEqual(result.mode, 'read-only-current-j-artifact-audit');
  assert.strictEqual(result.notesTxtMatch, true);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.geometry)), [{ slideNo: 40, mapSelector: false, selectorAfter: false, currentDeckMatches: 0, retained: false }]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.actions)), { imported: false, regenerated: false, mapRepaired: false, delivered: false, notified: false, audioStarted: false });
  assert.strictEqual(env.log.length, before);
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

console.log('pass ' + passed);
