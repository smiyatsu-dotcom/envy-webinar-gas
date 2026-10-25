// Contract tests for the fixed J No.7 regeneration.  All Drive/Slides data
// below is fake; no Google request can be made from this test.
const assert = require('assert');
const { makeEnv, loadGas } = require('./fake_gas.cjs');

const CASE = '隔離通し診断テスト社';
const HEARING = '1ur8hPUElZMrKuitnDo1j1wzr54EkmRY9efNMaEQFErs';
const FOLDER = '1Ko7cbJ2nW2C4uQbZfDnZjW6mfz_qmHZ0';
const SOURCE = '1g08BmnhKnzdtepHhUNFkAr2mhjww2cnC';
const MAP_ID = '1DukH9lGhK0T5W-PmkBBWUCVqUyJOwUtc';
const PRIOR = 'prior-fixed-deck';
const NEXT = 'new-fixed-deck';
const OLD7 = '例えば、皆さん経営していてこんなお悩みはありませんか？\nもし、当てはまるものがあれば、挙手🖐️マークでリアクションしていただけると嬉しいです！';
const REV1 = '2026-10-25T12:00:00.000Z';
const REV2 = '2026-10-25T13:00:00.000Z';

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('ok  ' + name); }
  catch (error) { console.log('NG  ' + name + '\n    ' + (error && error.stack || error)); process.exitCode = 1; }
}
function iterator(items) { let i = 0; return { hasNext: () => i < items.length, next: () => items[i++] }; }
function page(id, note) { return { getObjectId: () => id, getPageElements: () => [], getNotesPage: () => ({ getSpeakerNotesShape: () => ({ getText: () => ({ asString: () => note }) }) }) }; }
function rows(G) { return Array.from({ length: 106 }, (_, index) => ({ no: index + 1, daihon: G.FIXED_DAIHON[index + 1] || 'note-' + (index + 1) })); }
function mapFor(deckId, revision) { return { schema: 'consultation-slide-map/v2', caseName: CASE, presentationId: deckId, presentationRevision: revision, pages: Array.from({ length: 106 }, (_, index) => ({ slideNo: index + 1, pageId: deckId + '-p' + (index + 1) })) }; }

function fixture() {
  const env = makeEnv(), G = loadGas(env);
  env.props.INPUT_FOLDER_ID = 'normal-input'; env.props.OUTPUT_FOLDER_ID = 'normal-output'; env.props.TEMPLATE_SLIDE_ID = 'template'; env.props.ASSETS_FOLDER_ID = 'assets'; env.props.MGMT_SHEET_ID = 'normal-ledger';
  const parent = env.makeFolder('【隔離試験】配置永続化', null, '1GUmxo4J9cj0WYdh4IPj2lBi3RSxiMw73');
  const root = parent.addFolder('【隔離通し】診断から資料音声_20261023', '1oQSEwyopPcvX20xuxNDHy5Kh4TfvtCBO');
  root.addFolder('【出力】', '1Mxch8d4SEbddMfx-C2slbmQWWa1Uz16u');
  const output = root.addFolder('【出力】生成物', '18-xhLrskU8M9DvW2Rurqg5ST_NrOiSM2'); root.addFolder('【素材】投影資料', 'isolation-assets');
  const folder = output.addFolder(CASE, FOLDER);
  const ledger = env.makeSpreadsheet('16UgS3scd3X2D2Qqhzd1WfJuGkuNcOt5WbekQZLyKDpg', ['台帳']);
  const kousei = env.makeSpreadsheet('kousei-test-12345678901234567890', ['設定', '問題教育']);
  const settings = kousei.getSheetByName('設定');
  settings.getRange('B1').setValue(CASE + '_ヒアリングシート'); settings.getRange('B2').setValue(false); settings.getRange('B4').setValue(HEARING); settings.getRange('B22').setValue(false);
  kousei.getSheetByName('問題教育').getRange('E2').setValue(G.FIXED_DAIHON[7]);
  ledger.getSheets()[0].appendRow(['hearing', 'name', 'status', 'sheet', 'slide', 'message', 'updated', 'flow']);
  ledger.getSheets()[0].appendRow([HEARING, CASE + '_ヒアリングシート', '完了', kousei.getUrl(), '', '', '', '新']);
  const c = G.ISOLATED_J_REGEN_AUDIT_;
  c.folderId = FOLDER; c.sourceFileId = SOURCE; c.kouseiSheetId = kousei.getId(); c.mapFileId = MAP_ID; c.oldNo7 = OLD7; c.newNo7 = G.FIXED_DAIHON[7];
  const afterRows = rows(G), beforeRows = JSON.parse(JSON.stringify(afterRows)); beforeRows[6].daihon = OLD7;
  const sourceText = JSON.stringify(afterRows), beforeText = JSON.stringify(beforeRows);
  c.sourceSha256Before = G.隔離通し_J台本_sha256_(beforeText); c.sourceSha256After = G.隔離通し_J台本_sha256_(sourceText); c.other105Sha256 = G.隔離通し_J台本_sha256_(JSON.stringify(afterRows.filter(row => row.no !== 7)));
  const source = folder.addFile(c.sourceFileName, sourceText); source.getId = () => SOURCE;
  const snapshot = folder.addFile(c.writerSnapshotFileName, beforeText);
  folder.addFile(c.importRecordFileName, JSON.stringify({ status: 'completed', automaticRetry: false, folderId: FOLDER, targetScriptSha256: c.sourceSha256Before }));
  folder.addFile(c.writerRecordFileName, JSON.stringify({ schema: 'isolation-j-no7-hand-raise/v1', status: 'completed', automaticRetry: false, folderId: FOLDER, sourceFileId: SOURCE, kouseiSheetId: kousei.getId(), managementRow: 2, sourceSha256Before: c.sourceSha256Before, sourceSha256After: c.sourceSha256After, other105Sha256: c.other105Sha256, backupFileId: snapshot.getId(), sheetCell: c.sheetCell }));
  const oldMap = mapFor(PRIOR, REV1), mapFile = folder.addFile(c.mapFileName, JSON.stringify(oldMap)); mapFile.getId = () => MAP_ID;
  folder.addFile(c.geometryFileName, JSON.stringify({ schema: 'consultation-slide-geometry/v2', caseName: CASE, records: [] }));
  const decks = {
    [PRIOR]: { revision: REV1, slides: oldMap.pages.map((item, index) => page(item.pageId, beforeRows[index].daihon)) },
    [NEXT]: { revision: REV2, slides: mapFor(NEXT, REV2).pages.map((item, index) => page(item.pageId, afterRows[index].daihon)) }
  };
  const originalFile = G.DriveApp.getFileById;
  G.DriveApp.getFileById = id => {
    if (id === ledger.getId()) return { getParents: () => iterator([root]) };
    if (decks[id]) return { getMimeType: () => G.MimeType.GOOGLE_SLIDES, isTrashed: () => false, getParents: () => iterator([folder]), getLastUpdated: () => new Date(decks[id].revision) };
    return originalFile(id);
  };
  G.SlidesApp = { openById: id => { if (!decks[id]) throw new Error('unexpected deck ' + id); return { getSlides: () => decks[id].slides }; } };
  G.隔離通し_J台本_SHA256 = c.sourceSha256Before;
  return { env, G, folder, mapFile, oldMap, decks, afterRows };
}
function completeRecord(f) {
  const c = f.G.ISOLATED_J_REGEN_AUDIT_, nextMap = mapFor(NEXT, REV2), nextText = JSON.stringify(nextMap);
  f.mapFile.content = nextText;
  f.folder.addFile(c.regenerationRecordFileName, JSON.stringify({ schema: 'isolation-j-no7-regeneration/v1', status: 'completed', automaticRetry: false, folderId: FOLDER, sourceSha256: c.sourceSha256After, mapFileId: MAP_ID, previousDeckId: PRIOR, previousMapSha256: f.G.隔離通し_J台本_sha256_(JSON.stringify(f.oldMap)), newDeckId: NEXT, newMapSha256: f.G.隔離通し_J台本_sha256_(nextText), newDeckRevision: REV2 }));
}

test('preflight accepts only the fixed source/writer chain and the prior map deck that still has the snapshot script', () => {
  const f = fixture(), before = f.env.log.length;
  const result = f.G.隔離通し_J台本更新後の再組版前を検査する();
  assert.strictEqual(result.ok, true); assert.strictEqual(result.previousDeckId, PRIOR); assert.strictEqual(result.mapFileId, MAP_ID); assert.strictEqual(result.slideCount, 106);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.actions)), { regenerated: false, imported: false, audioStarted: false, delivered: false, notified: false });
  assert.strictEqual(f.env.log.length, before);
});

test('post audit requires a completed fixed execution record and validates only its new deck/map pair', () => {
  const f = fixture();
  assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /実行記録がありません/);
  completeRecord(f);
  const result = f.G.隔離通し_J台本更新後の再組版を読取検品する();
  assert.strictEqual(result.ok, true); assert.strictEqual(result.newDeckId, NEXT); assert.strictEqual(result.mapFileId, MAP_ID); assert.strictEqual(result.slideCount, 106);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.actions)), { regenerated: false, imported: false, audioStarted: false, delivered: false, notified: false });
});

test('the no-argument guarded wrapper records the builder result once, and a failed build becomes uncertain', () => {
  {
    const f = fixture();
    f.G.スライドを組む_ = () => { f.mapFile.content = JSON.stringify(mapFor(NEXT, REV2)); return true; };
    const result = f.G.隔離通し_J台本更新後のスライドを一度だけ組み直す();
    assert.strictEqual(result.ok, true); assert.strictEqual(result.newDeckId, NEXT); assert.strictEqual(result.mapFileId, MAP_ID);
    const record = JSON.parse(f.folder.getFilesByName(f.G.ISOLATED_J_REGEN_AUDIT_.regenerationRecordFileName).next().content);
    assert.strictEqual(record.status, 'completed'); assert.strictEqual(record.newDeckId, NEXT);
    assert.strictEqual(f.G.隔離通し_J台本更新後の再組版を読取検品する().ok, true);
  }
  {
    const f = fixture();
    f.G.スライドを組む_ = () => { throw new Error('builder failed'); };
    assert.throws(() => f.G.隔離通し_J台本更新後のスライドを一度だけ組み直す(), /結果が不確定/);
    const record = JSON.parse(f.folder.getFilesByName(f.G.ISOLATED_J_REGEN_AUDIT_.regenerationRecordFileName).next().content);
    assert.strictEqual(record.status, 'uncertain');
    assert.throws(() => f.G.隔離通し_J台本更新後のスライドを一度だけ組み直す(), /既にあります/);
  }
});

test('the real normal builder and generateSlides_ retain the fixed isolation config, then restore it after failure', () => {
  const f = fixture(), c = f.G.ISOLATED_J_REGEN_AUDIT_, observed = [];
  // This deliberately reaches the real スライドを組む_ and real generateSlides_.
  // Only their downstream Drive/Sheet writers are stopped after generateSlides_
  // has obtained getConfig_ and chosen the case folder.
  f.G.新_先に組む回_ = true;
  f.G.upsertRow_ = cfg => { observed.push({ step: 'ledger-write', cfg: Object.assign({}, cfg) }); };
  f.G.setKouseiStatus_ = () => {};
  f.G.知らせを積む_ = () => { throw new Error('notification must not be reached'); };
  f.G.止まったを知らせる_ = () => { throw new Error('notification must not be reached'); };
  f.G.readKouseiSheet_ = id => {
    observed.push({ step: 'read-kousei', id, cfg: Object.assign({}, f.G.getConfig_()) });
    return { name: CASE, ryoukin: '単一', rows: {} };
  };
  f.G.caseFolder_ = (cfg, name) => {
    observed.push({ step: 'case-folder', name, cfg: Object.assign({}, cfg), active: Object.assign({}, f.G.getConfig_()) });
    throw new Error('stop after checking isolated output root');
  };

  assert.throws(() => f.G.隔離通し_J台本更新後のスライドを一度だけ組み直す(), /結果が不確定/);
  assert.deepStrictEqual(observed.map(item => item.step), ['ledger-write', 'read-kousei', 'case-folder', 'ledger-write']);
  observed.forEach(item => {
    const cfg = item.cfg;
    assert.strictEqual(cfg.ISOLATION_TEST, true);
    assert.strictEqual(cfg.INPUT_FOLDER_ID, '1Mxch8d4SEbddMfx-C2slbmQWWa1Uz16u');
    assert.strictEqual(cfg.OUTPUT_FOLDER_ID, '18-xhLrskU8M9DvW2Rurqg5ST_NrOiSM2');
    assert.strictEqual(cfg.ASSETS_FOLDER_ID, 'isolation-assets');
    assert.strictEqual(cfg.MGMT_SHEET_ID, '16UgS3scd3X2D2Qqhzd1WfJuGkuNcOt5WbekQZLyKDpg');
    assert.notStrictEqual(cfg.OUTPUT_FOLDER_ID, f.env.props.OUTPUT_FOLDER_ID);
    assert.notStrictEqual(cfg.MGMT_SHEET_ID, f.env.props.MGMT_SHEET_ID);
  });
  const final = observed.find(item => item.step === 'case-folder');
  assert.strictEqual(final.name, CASE);
  assert.deepStrictEqual(final.active, final.cfg);
  const record = JSON.parse(f.folder.getFilesByName(c.regenerationRecordFileName).next().content);
  assert.strictEqual(record.status, 'uncertain');
  assert.strictEqual(record.automaticRetry, false);
  assert.strictEqual(f.G.ISOLATION_TEST_CONFIG_, null);
  assert.strictEqual(f.G.getConfig_().OUTPUT_FOLDER_ID, 'normal-output');
  assert.strictEqual(f.G.getConfig_().ASSETS_FOLDER_ID, 'assets');
  assert.strictEqual(f.G.getConfig_().MGMT_SHEET_ID, 'normal-ledger');
});

test('preflight and post audit fail closed for map ID substitution, stale prior notes, uncertain execution, and unrecorded new deck', () => {
  {
    const f = fixture(); f.mapFile.getId = () => 'other-map';
    assert.throws(() => f.G.隔離通し_J台本更新後の再組版前を検査する(), /対応表ID/);
  }
  {
    const f = fixture(); f.decks[PRIOR].slides[6] = page(PRIOR + '-p7', 'not-the-snapshot');
    assert.throws(() => f.G.隔離通し_J台本更新後の再組版前を検査する(), /notes/);
  }
  {
    const f = fixture(); completeRecord(f);
    const record = f.folder.getFilesByName(f.G.ISOLATED_J_REGEN_AUDIT_.regenerationRecordFileName).next(); const body = JSON.parse(record.content); body.status = 'uncertain'; record.content = JSON.stringify(body);
    assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /completed実行記録/);
  }
  {
    const f = fixture(); completeRecord(f);
    const record = f.folder.getFilesByName(f.G.ISOLATED_J_REGEN_AUDIT_.regenerationRecordFileName).next(); const body = JSON.parse(record.content); body.newDeckId = 'unrecorded'; record.content = JSON.stringify(body);
    assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /hash|deck|対応表/);
  }
});

test('contract source has no arbitrary deck argument and both audits are read-only', () => {
  const source = require('fs').readFileSync(require('path').join(__dirname, '..', '59_j_regeneration_audit.js'), 'utf8');
  const pre = source.slice(source.indexOf('function 隔離通し_J台本更新後の再組版前を検査する'), source.indexOf('function isolatedJRegenerationRecord_'));
  const writer = source.slice(source.indexOf('function 隔離通し_J台本更新後のスライドを一度だけ組み直す'), source.indexOf('function 隔離通し_J台本更新後の再組版を読取検品する'));
  const post = source.slice(source.indexOf('function 隔離通し_J台本更新後の再組版を読取検品する'), source.length);
  assert.doesNotMatch(pre, /setContent\s*\(|createFile\s*\(|スライドを組む_\s*\(/);
  assert.doesNotMatch(post, /setContent\s*\(|createFile\s*\(|スライドを組む_\s*\(/);
  assert.match(writer, /function\s+隔離通し_J台本更新後のスライドを一度だけ組み直す\s*\(\s*\)/);
  assert.match(writer, /withIsolationTestConfig_\s*\(\s*target\.cfg\s*,/);
  assert.doesNotMatch(writer, /processNewFlow_|知らせを積む_|止まったを知らせる_/);
  assert.doesNotMatch(source, /function\s+隔離通し_J台本更新後の再組版前を検査する\s*\([^)]*[^\s)]/);
  assert.doesNotMatch(source, /function\s+隔離通し_J台本更新後の再組版を読取検品する\s*\([^)]*[^\s)]/);
});

console.log('pass ' + passed);
