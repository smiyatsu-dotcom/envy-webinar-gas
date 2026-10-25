// No.7 source-and-sheet writer regression tests. Uses fake GAS services only.
const assert = require('assert');
const { makeEnv, loadGas } = require('./fake_gas.cjs');

const CASE = '隔離通し診断テスト社';
const HEARING = '1ur8hPUElZMrKuitnDo1j1wzr54EkmRY9efNMaEQFErs';
const FOLDER = '1Ko7cbJ2nW2C4uQbZfDnZjW6mfz_qmHZ0';
const SOURCE_ID = '1g08BmnhKnzdtepHhUNFkAr2mhjww2cnC';
const SOURCE_NAME = '【台本JSON】隔離通し診断テスト社.txt';
const OLD = '例えば、皆さん経営していてこんなお悩みはありませんか？\nもし、当てはまるものがあれば、挙手🖐️マークでリアクションしていただけると嬉しいです！';
const NEW = '例えば、皆さん経営していてこんなお悩みはありませんか？\n当てはまる方は、挙手マークを押してください。';

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('ok  ' + name); }
  catch (error) { console.log('NG  ' + name + '\n    ' + (error && error.stack || error)); process.exitCode = 1; }
}
function iterator(items) { let i = 0; return { hasNext: () => i < items.length, next: () => items[i++] }; }
function scriptRows() { return Array.from({ length: 106 }, (_, i) => ({ no: i + 1, daihon: i === 6 ? OLD : 'line-' + (i + 1) })); }
function otherRows(rows) { return rows.filter(row => row.no !== 7); }

function fixture({ staleSource = false, staleCell = false } = {}) {
  const env = makeEnv();
  const G = loadGas(env);
  const output = env.makeFolder('【出力】生成物', null, 'test-output');
  const folder = output.addFolder(CASE, FOLDER);
  const sourceRows = scriptRows();
  if (staleSource) sourceRows[6].daihon = 'not-the-approved-old-text';
  const sourceText = JSON.stringify(sourceRows);
  const source = folder.addFile(SOURCE_NAME, sourceText);
  source.getId = () => SOURCE_ID;
  source.getParents = () => iterator([folder]);
  const ledger = env.makeSpreadsheet('ledger', ['台帳']);
  ledger.getSheets()[0].appendRow(['hearing', 'name', 'status', 'sheet', 'slide', 'message', 'updated', 'flow']);
  const kousei = env.makeSpreadsheet('kousei-test-12345678901234567890', ['設定', 'イントロ', '問題教育', 'ストーリー', 'オファー']);
  const settings = kousei.getSheetByName('設定');
  settings.getRange('B1').setValue(CASE + '_ヒアリングシート');
  settings.getRange('B2').setValue(false);
  settings.getRange('B4').setValue(HEARING);
  settings.getRange('B22').setValue(false);
  const intro = kousei.getSheetByName('イントロ');
  intro.appendRow(['No.', '基本構成', '小分類', '構成テキスト', '台本', 'アイコン', '背景']);
  intro.appendRow([7, '', '', '', staleCell ? 'manually-changed' : OLD, '', '']);
  ledger.getSheets()[0].appendRow([HEARING, CASE + '_ヒアリングシート', '完了', kousei.getUrl(), '', '', '', '新']);
  G.隔離通し_生成_生成物rootID = output.getId();
  G.隔離通し_生成_案件名 = CASE;
  G.隔離通し_生成_ヒアリングID = HEARING;
  G.隔離通し_生成設定を作る_ = () => ({ OUTPUT_FOLDER_ID: output.getId() });
  G.withIsolationTestConfig_ = (_cfg, fn) => fn();
  G.caseFolderIfExists_ = () => folder;
  G.隔離通し_生成の親を確かめる_ = () => {};
  G.mgmtSheet_ = () => ledger.getSheets()[0];
  G.registry_ = () => ({ [HEARING]: 2 });
  const realGetFile = G.DriveApp.getFileById;
  G.DriveApp.getFileById = id => id === SOURCE_ID ? source : realGetFile(id);
  G.隔離通し_No7挙手文言修正_.expectedSourceSha256 = G.隔離通し_J台本_sha256_(sourceText);
  G.processImports_ = () => { throw new Error('must not import'); };
  G.processNewFlow_ = () => { throw new Error('must not progress flow'); };
  G.スライドを組む_ = () => { throw new Error('must not regenerate'); };
  return { env, G, folder, source, sourceRows, kousei };
}

test('preflight is read-only and returns the fixed source, old SHA, E cell, and other-105 fingerprint', () => {
  const { env, G, source, kousei } = fixture();
  const beforeLog = env.log.length, beforeSource = source.content, beforeCell = kousei.getSheetByName('イントロ').getRange('E2').getValue();
  const result = G.隔離通し_No7挙手文言修正を検査する();
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.mode, 'read-only-fixed-no7-preflight');
  assert.strictEqual(result.sourceFileId, SOURCE_ID);
  assert.strictEqual(result.sheetCell, 'イントロ!E2');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.actions)), { sourceWritten: false, sheetWritten: false, imported: false, regenerated: false, audioStarted: false, delivered: false, notified: false });
  assert.strictEqual(source.content, beforeSource);
  assert.strictEqual(kousei.getSheetByName('イントロ').getRange('E2').getValue(), beforeCell);
  assert.strictEqual(env.log.length, beforeLog);
});

test('updates exactly No.7 in source JSON and sheet E cell; all 105 other rows are stable', () => {
  const { env, G, folder, source, sourceRows, kousei } = fixture();
  const result = G.隔離通し_No7挙手文言を一度反映する();
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.mode, 'fixed-no7-source-and-sheet-only');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.actions)), { imported: false, regenerated: false, audioStarted: false, delivered: false, notified: false });
  const afterRows = JSON.parse(source.content);
  assert.strictEqual(afterRows.find(row => row.no === 7).daihon, NEW);
  assert.deepStrictEqual(afterRows.filter(row => row.no !== 7), otherRows(sourceRows));
  assert.strictEqual(kousei.getSheetByName('イントロ').getRange('E2').getValue(), NEW);
  const record = JSON.parse(folder.getFilesByName('【不確定】No7挙手文言修正記録.json').next().content);
  assert.strictEqual(record.status, 'completed');
  assert.strictEqual(record.sheetCell, 'イントロ!E2');
  assert.strictEqual(folder.getFilesByName('【保全】No7挙手文言修正前台本.txt').next().content, JSON.stringify(sourceRows));
  const beforeRetry = env.log.length;
  assert.throws(() => G.隔離通し_No7挙手文言を一度反映する(), /自動再実行せず/);
  assert.strictEqual(env.log.length, beforeRetry);
});

test('fails before any write when either the source old text or sheet old text differs', () => {
  for (const options of [{ staleSource: true }, { staleCell: true }]) {
    const { env, G, source } = fixture(options);
    const beforeLog = env.log.length, beforeText = source.content;
    assert.throws(() => G.隔離通し_No7挙手文言を一度反映する(), /旧文|旧SHA-256/);
    assert.strictEqual(source.content, beforeText);
    assert.strictEqual(env.log.length, beforeLog);
  }
});

test('records uncertain and blocks retry if the source changes but sheet writer cannot complete', () => {
  const { G, folder, source } = fixture();
  G.直すセルを書く_ = () => ({ 書いた: [], 違う: ['forced sheet failure'] });
  assert.throws(() => G.隔離通し_No7挙手文言を一度反映する(), /結果が不確定/);
  assert.strictEqual(JSON.parse(source.content).find(row => row.no === 7).daihon, NEW);
  assert.strictEqual(JSON.parse(folder.getFilesByName('【不確定】No7挙手文言修正記録.json').next().content).status, 'uncertain');
  assert.throws(() => G.隔離通し_No7挙手文言を一度反映する(), /自動再実行せず/);
});

test('writer source never invokes import, flow, regeneration, audio, delivery, or notification', () => {
  const source = require('fs').readFileSync(require('path').join(__dirname, '..', '56_geometry_isolation.js'), 'utf8');
  const body = source.slice(source.indexOf('function 隔離通し_No7挙手文言を一度反映する'), source.indexOf('function 隔離通しスライド失敗を読む'));
  assert.doesNotMatch(body, /processImports_\s*\(/);
  assert.doesNotMatch(body, /processNewFlow_\s*\(/);
  assert.doesNotMatch(body, /スライドを組む_\s*\(/);
  assert.doesNotMatch(body, /generateSlides_\s*\(/);
  assert.doesNotMatch(body, /notify\w*\s*\(/i);
});

console.log('pass ' + passed);
