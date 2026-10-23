// 固定隔離通し生成の安全境界。Google 本体へは接続しない。
const assert = require('assert');
const { makeEnv, loadGas } = require('./fake_gas.cjs');

let pass = 0;
function test(name, fn) {
  try { fn(); pass++; console.log('ok  ' + name); }
  catch (e) { console.log('NG  ' + name + '\n    ' + (e && e.stack || e)); process.exitCode = 1; }
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
  const startBox = root.addFolder('【出力】', '1Mxch8d4SEbddMfx-C2slbmQWWa1Uz16u');
  const output = root.addFolder('【出力】生成物', '18-xhLrskU8M9DvW2Rurqg5ST_NrOiSM2');
  const assets = root.addFolder('【素材】投影資料', 'isolation-assets');
  const ss = env.makeSpreadsheet('16UgS3scd3X2D2Qqhzd1WfJuGkuNcOt5WbekQZLyKDpg', ['シート1']);
  ss.getSheets()[0].appendRow(['ヒアリングファイルID', '案件名', 'ステータス', '構成シートURL', 'スライドURL', 'メッセージ', '更新日時', '流れ']);
  ss.getSheets()[0].appendRow(['1ur8hPUElZMrKuitnDo1j1wzr54EkmRY9efNMaEQFErs', '隔離通し診断テスト社_ヒアリングシート', 'AI出力待ち', '', '', '', '', '新']);
  const ledgerFile = { getParents: () => ({ hasNext: (() => { let n = 0; return () => n++ === 0; })(), next: () => root }) };
  const oldFile = G.DriveApp.getFileById;
  G.DriveApp.getFileById = id => id === ss.getId() ? ledgerFile : oldFile(id);
  return { env, G, root, startBox, output, assets, ss };
}

test('withIsolationTestConfig_: 実行中だけ設定を差し替え、Propertiesは書かない', () => {
  const { env, G } = fixture();
  const before = Object.assign({}, env.props);
  const value = G.withIsolationTestConfig_({ ISOLATION_TEST: true, OUTPUT_FOLDER_ID: 'isolated' }, () => G.getConfig_().OUTPUT_FOLDER_ID);
  assert.strictEqual(value, 'isolated');
  assert.deepStrictEqual(env.props, before);
  assert.throws(() => G.withIsolationTestConfig_({}, () => {}), /形が違います/);
});

test('隔離通し生成設定: 入出力・素材・台帳を固定照合し、通常設定と混ぜない', () => {
  const { G, startBox, assets } = fixture();
  const cfg = G.隔離通し_生成設定を作る_();
  assert.strictEqual(cfg.ISOLATION_TEST, true);
  assert.strictEqual(cfg.OUTPUT_FOLDER_ID, '18-xhLrskU8M9DvW2Rurqg5ST_NrOiSM2');
  assert.strictEqual(cfg.INPUT_FOLDER_ID, startBox.getId());
  assert.strictEqual(cfg.ASSETS_FOLDER_ID, assets.getId());
  assert.strictEqual(cfg.MGMT_SHEET_ID, '16UgS3scd3X2D2Qqhzd1WfJuGkuNcOt5WbekQZLyKDpg');
  assert.strictEqual(cfg.ISOLATION_START_BOX_ID, '1Mxch8d4SEbddMfx-C2slbmQWWa1Uz16u');
  assert.strictEqual(cfg.TEMPLATE_SLIDE_ID, 'template');
  assert.notStrictEqual(cfg.INPUT_FOLDER_ID, 'normal-input');
  assert.notStrictEqual(cfg.OUTPUT_FOLDER_ID, 'normal-output');
  assert.notStrictEqual(cfg.ASSETS_FOLDER_ID, 'assets');
  assert.notStrictEqual(cfg.MGMT_SHEET_ID, 'normal-ledger');
});

test('隔離通し生成設定: 通常設定が固定素材rootを指す混線はfail-closedする', () => {
  const { env, G } = fixture();
  env.props.ASSETS_FOLDER_ID = 'isolation-assets';
  assert.throws(() => G.隔離通し_生成設定を作る_(), /通常設定と隔離fixture/);
});

test('隔離通し入口: 正規開始JSONを本文不変で案件フォルダへ渡し、通常処理を一度呼ぶ', () => {
  const { G, startBox, output } = fixture();
  const startName = '【開始】隔離通し診断テスト社.json';
  const startFile = startBox.addFile(startName, JSON.stringify({
    案件名: '隔離通し診断テスト社',
    ヒアリングシートID: '1ur8hPUElZMrKuitnDo1j1wzr54EkmRY9efNMaEQFErs',
    セル: Array(81).fill({})
  }));
  const source = startFile.content;
  let startCalls = 0;
  G.processStartOrders_ = () => { startCalls++; };
  const result = G.隔離通し生成を開始する();
  assert.strictEqual(result.phase, 'start');
  assert.strictEqual(result.handed, true);
  assert.strictEqual(startFile.content, source);
  assert.strictEqual(startBox.getFilesByName(startName).hasNext(), false);
  const destination = output.getFoldersByName('隔離通し診断テスト社').next();
  assert.strictEqual(startFile.parent, destination);
  assert.strictEqual(startCalls, 1);
  assert.strictEqual(G.SpreadsheetApp.openById('16UgS3scd3X2D2Qqhzd1WfJuGkuNcOt5WbekQZLyKDpg').getSheets()[0].getDataRange().getValues().length, 1);
});

test('隔離通し進行: 開始を触らず、正規の取り込み後に新フローを進める', () => {
  const { G } = fixture();
  const calls = [];
  G.processImports_ = () => calls.push('imports');
  G.processNewFlow_ = () => calls.push('new-flow');
  G.processStartOrders_ = () => calls.push('start');
  const result = G.隔離通し生成を進める();
  assert.strictEqual(result.phase, 'advance');
  assert.deepStrictEqual(calls, ['imports', 'new-flow']);
});

console.log('pass ' + pass);
