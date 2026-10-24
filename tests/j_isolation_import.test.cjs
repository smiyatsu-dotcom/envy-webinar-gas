// J最終台本の固定隔離取込。Google本体には接続せず、固定対象以外を
// processImports_ に渡せないことを確認する。
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { makeEnv, loadGas } = require('./fake_gas.cjs');

const root = path.join(__dirname, '..');
const CASE = '隔離通し診断テスト社';
const HEARING = '1ur8hPUElZMrKuitnDo1j1wzr54EkmRY9efNMaEQFErs';
const FOLDER = '1Ko7cbJ2nW2C4uQbZfDnZjW6mfz_qmHZ0';
const FILE = '【台本JSON】隔離通し診断テスト社.txt';

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('ok  ' + name); }
  catch (error) { console.log('NG  ' + name + '\n    ' + (error && error.stack || error)); process.exitCode = 1; }
}
function exactRows(prefix) {
  return JSON.stringify(Array.from({ length: 106 }, (_, i) => ({ no: i + 1, daihon: prefix + (i + 1) })));
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
  const rootFolder = parent.addFolder('【隔離通し】診断から資料音声_20261023', '1oQSEwyopPcvX20xuxNDHy5Kh4TfvtCBO');
  rootFolder.addFolder('【出力】', '1Mxch8d4SEbddMfx-C2slbmQWWa1Uz16u');
  const output = rootFolder.addFolder('【出力】生成物', '18-xhLrskU8M9DvW2Rurqg5ST_NrOiSM2');
  rootFolder.addFolder('【素材】投影資料', 'isolation-assets');
  const folder = output.addFolder(CASE, FOLDER);
  const oldText = exactRows('old-');
  const nextText = exactRows('next-');
  const script = folder.addFile(FILE, oldText);
  const ledger = env.makeSpreadsheet('16UgS3scd3X2D2Qqhzd1WfJuGkuNcOt5WbekQZLyKDpg', ['台帳']);
  const ls = ledger.getSheets()[0];
  ls.appendRow(['ヒアリングファイルID', '案件名', 'ステータス', '構成シートURL', 'スライドURL', 'メッセージ', '更新日時', '流れ']);
  const kouseiId = '1YvJszV8Mz1tN1DVZ9xOIY5D7XGW754QQFVU_oyi34AQ';
  ls.appendRow([HEARING, CASE + '_ヒアリングシート', '完了', 'https://docs.google.com/spreadsheets/d/' + kouseiId + '/edit', '', '', '', '新']);
  const kousei = env.makeSpreadsheet(kouseiId, ['設定']);
  const st = kousei.getSheetByName('設定');
  st.getRange('B1').setValue(CASE + '_ヒアリングシート');
  st.getRange('B2').setValue(false);
  st.getRange('B4').setValue(HEARING);
  st.getRange('B22').setValue(false);
  const oldGetFile = G.DriveApp.getFileById;
  G.DriveApp.getFileById = id => id === ledger.getId() ? { getParents: () => ({ hasNext: (() => { let n = 0; return () => n++ === 0; })(), next: () => rootFolder }) } : oldGetFile(id);
  // The test changes only vm globals. The source still pins the real hashes.
  G.隔離通し_J台本_旧SHA256 = G.隔離通し_J台本_sha256_(oldText);
  G.隔離通し_J台本_SHA256 = G.隔離通し_J台本_sha256_(nextText);
  return { env, G, folder, script, ledger, oldText, nextText };
}

test('source pins the supplied old and next script hashes plus the literal isolation folder', () => {
  const source = fs.readFileSync(path.join(root, '56_geometry_isolation.js'), 'utf8');
  assert.match(source, /1Ko7cbJ2nW2C4uQbZfDnZjW6mfz_qmHZ0/);
  assert.match(source, /a3d6e2cd4eaaaf57dacf64818ca33486cac14d907b6de562a48868da9404918a/);
  assert.match(source, /7eb361136315b71a1505d50e94cbeef12d28feb1846631ec67370072e5331645/);
  assert.match(source, /processImports_\(\)/);
  assert.doesNotMatch(source.slice(source.indexOf('function 隔離通し_J台本を正規取込する')), /processNewFlow_\(\)/);
});

test('read-only preflight identifies exactly one fixed folder, row, and unapproved sheet', () => {
  const { G } = fixture();
  const result = G.隔離通し_J台本取込を検査();
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result)), {
    ok: true, caseName: CASE, folderId: FOLDER, managementRow: 2,
    kouseiSheetId: '1YvJszV8Mz1tN1DVZ9xOIY5D7XGW754QQFVU_oyi34AQ',
    currentScriptSha256: G.隔離通し_J台本_旧SHA256,
    expectedNextScriptSha256: G.隔離通し_J台本_SHA256,
    deliveryOrApproval: false, mode: 'read-only-preflight'
  });
});

test('exact 106-row script reaches only processImports_ and remains unapproved/undelivered', () => {
  const { G, script, ledger, nextText } = fixture();
  let imports = 0;
  G.processNewFlow_ = () => { throw new Error('new-flow must not run'); };
  G.processImports_ = () => {
    imports++;
    assert.strictEqual(script.content, nextText);
    ledger.getSheets()[0].getRange('C2').setValue('レビュー待ち');
  };
  const result = G.隔離通し_J台本を正規取込する(nextText);
  assert.strictEqual(imports, 1);
  assert.strictEqual(result.status, 'レビュー待ち');
  assert.strictEqual(result.mode, 'processImports-only');
  assert.strictEqual(ledger.getSheets()[0].getRange('C2').getValue(), 'レビュー待ち');
});

test('wrong payload or another importable ledger row fails before the fixed script is replaced', () => {
  const { G, script, ledger, oldText } = fixture();
  assert.throws(() => G.隔離通し_J台本を正規取込する(oldText), /SHA-256/);
  assert.strictEqual(script.content, oldText);
  ledger.getSheets()[0].appendRow(['other', '他社_ヒアリングシート', 'AI出力待ち', 'https://docs.google.com/spreadsheets/d/other/edit', '', '', '', '新']);
  assert.throws(() => G.隔離通し_J台本取込を検査(), /固定の新フロー1件|固定fixture以外/);
  assert.strictEqual(script.content, oldText);
});

console.log('pass ' + passed);
