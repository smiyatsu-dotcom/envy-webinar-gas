// Contract tests for the fixed J No.7 regeneration. All Drive/Slides data is fake.
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
// Map and direct-Drive observations deliberately differ by 817ms. The test
// proves that content evidence—not a timestamp tolerance—governs acceptance.
const REV1 = '2026-10-25T12:00:00.000Z';
const REV1_DRIVE = '2026-10-25T11:59:59.183Z';
const REV2 = '2026-10-25T13:00:00.000Z';
const REV2_DRIVE = '2026-10-25T12:59:59.183Z';
const G40 = { left: 1, top: 2, width: 3, height: 4 };

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('ok  ' + name); }
  catch (error) { console.log('NG  ' + name + '\n    ' + (error && error.stack || error)); process.exitCode = 1; }
}
function iterator(items) { let i = 0; return { hasNext: () => i < items.length, next: () => items[i++] }; }
function shape(id, text, geometry = G40) {
  const state = { text, geometry: Object.assign({}, geometry) };
  return {
    getObjectId: () => id, getPageElementType: () => 'SHAPE', getTitle: () => '', getDescription: () => '',
    getLeft: () => state.geometry.left, getTop: () => state.geometry.top, getWidth: () => state.geometry.width, getHeight: () => state.geometry.height,
    asShape: () => ({ getText: () => ({ asString: () => state.text }) }),
    setGeometry: value => { state.geometry = Object.assign({}, value); }, setText: value => { state.text = value; }
  };
}
function image(id, geometry = { left: 0, top: 0, width: 10, height: 10 }) {
  const state = { geometry: Object.assign({}, geometry) };
  return {
    getObjectId: () => id, getPageElementType: () => 'IMAGE', getTitle: () => 'decorative', getDescription: () => 'fixture image',
    getLeft: () => state.geometry.left, getTop: () => state.geometry.top, getWidth: () => state.geometry.width, getHeight: () => state.geometry.height,
    setGeometry: value => { state.geometry = Object.assign({}, value); }
  };
}
function page(id, note, elements = [], skipped = false) {
  const state = { note, elements, skipped };
  return {
    getObjectId: () => id, getPageElements: () => state.elements,
    getNotesPage: () => ({ getSpeakerNotesShape: () => ({ getText: () => ({ asString: () => state.note }) }) }),
    isSkipped: () => state.skipped, setNote: value => { state.note = value; }, setSkipped: value => { state.skipped = value === true; }
  };
}
function rows(G) { return Array.from({ length: 106 }, (_, index) => ({ no: index + 1, daihon: G.FIXED_DAIHON[index + 1] || 'note-' + (index + 1) })); }
function mapFor(deckId, revision) {
  return { schema: 'consultation-slide-map/v2', caseName: CASE, presentationId: deckId, presentationRevision: revision,
    pages: Array.from({ length: 106 }, (_, index) => {
      const item = { slideNo: index + 1, pageId: deckId + '-p' + (index + 1) };
      if (index === 39) item.targetSelector = { type: 'SHAPE', title: '', description: '', text: 'heading-40', templateGeometry: Object.assign({}, G40) };
      return item;
    }) };
}
function slidesFor(map, scriptRows) {
  return map.pages.map((item, index) => {
    const elements = [];
    if (index === 0) elements.push(image(item.pageId + '-image'));
    if (index === 39) elements.push(shape(item.pageId + '-heading', 'heading-40'));
    return page(item.pageId, index === 103 ? '' : scriptRows[index].daihon, elements, index === 103);
  });
}

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
  const oldMap = mapFor(PRIOR, REV1), nextMap = mapFor(NEXT, REV2), mapFile = folder.addFile(c.mapFileName, JSON.stringify(oldMap)); mapFile.getId = () => MAP_ID;
  folder.addFile(c.geometryFileName, JSON.stringify({ schema: 'consultation-slide-geometry/v2', caseName: CASE, records: [{ status: 'active', slideNo: 40, after: Object.assign({}, G40) }] }));
  const decks = { [PRIOR]: { revision: REV1_DRIVE, slides: slidesFor(oldMap, beforeRows) }, [NEXT]: { revision: REV2_DRIVE, slides: slidesFor(nextMap, afterRows) } };
  const originalFile = G.DriveApp.getFileById;
  G.DriveApp.getFileById = id => {
    if (id === ledger.getId()) return { getParents: () => iterator([root]) };
    if (decks[id]) return { getMimeType: () => G.MimeType.GOOGLE_SLIDES, isTrashed: () => false, getParents: () => iterator([folder]), getLastUpdated: () => new Date(decks[id].revision) };
    return originalFile(id);
  };
  G.SlidesApp = { openById: id => { if (!decks[id]) throw new Error('unexpected deck ' + id); return { getSlides: () => decks[id].slides }; } };
  G.隔離通し_J台本_SHA256 = c.sourceSha256Before;
  return { env, G, folder, mapFile, oldMap, nextMap, decks, afterRows };
}
function mapHash(f, map) { return f.G.隔離通し_J台本_sha256_(JSON.stringify(map)); }
function completeRecord(f) {
  const c = f.G.ISOLATED_J_REGEN_AUDIT_, nextText = JSON.stringify(f.nextMap), beforeRows = JSON.parse(JSON.stringify(f.afterRows));
  beforeRows[6].daihon = OLD7;
  const previous = f.G.isolatedJRegenMapDeck_(f.folder, beforeRows, mapHash(f, f.oldMap), PRIOR, 'before');
  f.mapFile.content = nextText;
  assert.strictEqual(f.G.isolatedJRegenHash_(f.mapFile.getBlob().getDataAsString('UTF-8')), mapHash(f, f.nextMap));
  const after = f.G.isolatedJRegenMapDeck_(f.folder, f.afterRows, mapHash(f, f.nextMap), NEXT, 'after');
  f.folder.addFile(c.regenerationRecordFileName, JSON.stringify({ schema: 'isolation-j-no7-regeneration/v2', status: 'completed', automaticRetry: false, folderId: FOLDER, sourceSha256: c.sourceSha256After, mapFileId: MAP_ID, previousDeckId: PRIOR, previousMapSha256: mapHash(f, f.oldMap), previousDeckFingerprint: previous.fingerprint, previousTimestampDiagnostic: previous.timestampDiagnostic, newDeckId: NEXT, newMapSha256: mapHash(f, f.nextMap), newDeckFingerprint: after.fingerprint, newTimestampDiagnostic: after.timestampDiagnostic }));
}
function record(f) { return f.folder.getFilesByName(f.G.ISOLATED_J_REGEN_AUDIT_.regenerationRecordFileName).next(); }

test('preflight accepts source/writer chain with timestamp diagnostic mismatch and returns content fingerprint', () => {
  const f = fixture(), before = f.env.log.length, result = f.G.隔離通し_J台本更新後の再組版前を検査する();
  assert.strictEqual(result.ok, true); assert.strictEqual(result.previousDeckId, PRIOR); assert.strictEqual(result.mapFileId, MAP_ID); assert.strictEqual(result.slideCount, 106);
  assert.strictEqual(result.timestampDiagnostic.mapPresentationRevision, REV1); assert.strictEqual(result.timestampDiagnostic.deckLastUpdated, REV1_DRIVE); assert.notStrictEqual(result.timestampDiagnostic.mapPresentationRevision, result.timestampDiagnostic.deckLastUpdated);
  assert.strictEqual(result.previousDeckFingerprint.pageCount, 106); assert.strictEqual(result.previousDeckFingerprint.metadataOnlyElements, 1);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.previousDeckFingerprint.metadataOnly)), [{ type: 'IMAGE', count: 1, limitation: 'image-pixels-and-dynamic-content-url-not-readable-via-SlidesApp' }]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.actions)), { regenerated: false, imported: false, audioStarted: false, delivered: false, notified: false }); assert.strictEqual(f.env.log.length, before);
});

test('post audit requires completed record and validates only its new deck/map fingerprint', () => {
  const f = fixture(); assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /実行記録がありません/); completeRecord(f);
  const result = f.G.隔離通し_J台本更新後の再組版を読取検品する();
  assert.strictEqual(result.ok, true); assert.strictEqual(result.newDeckId, NEXT); assert.strictEqual(result.mapFileId, MAP_ID); assert.strictEqual(result.slideCount, 106); assert.strictEqual(result.deckFingerprint.elementCount, 2); assert.strictEqual(result.timestampDiagnostic.deckLastUpdated, REV2_DRIVE);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.actions)), { regenerated: false, imported: false, audioStarted: false, delivered: false, notified: false });
});

test('guarded wrapper records builder result once, and failed build becomes uncertain', () => {
  { const f = fixture(); f.G.スライドを組む_ = () => { f.mapFile.content = JSON.stringify(f.nextMap); return true; };
    const result = f.G.隔離通し_J台本更新後のスライドを一度だけ組み直す(); assert.strictEqual(result.ok, true); assert.strictEqual(result.newDeckId, NEXT);
    const body = JSON.parse(record(f).content); assert.strictEqual(body.schema, 'isolation-j-no7-regeneration/v2'); assert.strictEqual(body.status, 'completed'); assert.strictEqual(body.newDeckFingerprint.pageCount, 106); assert.strictEqual(f.G.隔離通し_J台本更新後の再組版を読取検品する().ok, true); }
  { const f = fixture(); f.G.スライドを組む_ = () => { throw new Error('builder failed'); };
    assert.throws(() => f.G.隔離通し_J台本更新後のスライドを一度だけ組み直す(), /結果が不確定/); assert.strictEqual(JSON.parse(record(f).content).status, 'uncertain'); assert.throws(() => f.G.隔離通し_J台本更新後のスライドを一度だけ組み直す(), /既にあります/); }
});

test('real normal builder keeps fixed isolation config and restores it after failure', () => {
  const f = fixture(), observed = []; f.G.新_先に組む回_ = true;
  f.G.upsertRow_ = cfg => { observed.push({ step: 'ledger-write', cfg: Object.assign({}, cfg) }); }; f.G.setKouseiStatus_ = () => {};
  f.G.知らせを積む_ = () => { throw new Error('notification must not be reached'); }; f.G.止まったを知らせる_ = () => { throw new Error('notification must not be reached'); };
  f.G.readKouseiSheet_ = id => { observed.push({ step: 'read-kousei', id, cfg: Object.assign({}, f.G.getConfig_()) }); return { name: CASE, ryoukin: '単一', rows: {} }; };
  f.G.caseFolder_ = (cfg, name) => { observed.push({ step: 'case-folder', name, cfg: Object.assign({}, cfg), active: Object.assign({}, f.G.getConfig_()) }); throw new Error('stop after checking isolated output root'); };
  assert.throws(() => f.G.隔離通し_J台本更新後のスライドを一度だけ組み直す(), /結果が不確定/); assert.deepStrictEqual(observed.map(item => item.step), ['ledger-write', 'read-kousei', 'case-folder', 'ledger-write']);
  observed.forEach(item => { const cfg = item.cfg; assert.strictEqual(cfg.ISOLATION_TEST, true); assert.strictEqual(cfg.INPUT_FOLDER_ID, '1Mxch8d4SEbddMfx-C2slbmQWWa1Uz16u'); assert.strictEqual(cfg.OUTPUT_FOLDER_ID, '18-xhLrskU8M9DvW2Rurqg5ST_NrOiSM2'); assert.strictEqual(cfg.ASSETS_FOLDER_ID, 'isolation-assets'); assert.strictEqual(cfg.MGMT_SHEET_ID, '16UgS3scd3X2D2Qqhzd1WfJuGkuNcOt5WbekQZLyKDpg'); });
  const final = observed.find(item => item.step === 'case-folder'); assert.strictEqual(final.name, CASE); assert.deepStrictEqual(final.active, final.cfg); assert.strictEqual(JSON.parse(record(f).content).status, 'uncertain'); assert.strictEqual(f.G.ISOLATION_TEST_CONFIG_, null); assert.strictEqual(f.G.getConfig_().OUTPUT_FOLDER_ID, 'normal-output');
});

test('post audit fail-closes for map swap, note, fixed skip, geometry, add/remove element, and malformed fingerprint', () => {
  { const f = fixture(); completeRecord(f); f.mapFile.content = JSON.stringify(Object.assign({}, f.nextMap, { presentationRevision: 'different' })); assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /対応表hash/); }
  { const f = fixture(); completeRecord(f); f.decks[NEXT].slides[8].setNote('concurrent note mutation'); assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /notes|fingerprint/); }
  { const f = fixture(); completeRecord(f); f.decks[NEXT].slides[103].setSkipped(false); assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /スキップ状態/); }
  { const f = fixture(); completeRecord(f); f.decks[NEXT].slides[39].getPageElements()[0].setGeometry({ left: 9, top: 2, width: 3, height: 4 }); assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /geometry|fingerprint/); }
  { const f = fixture(); completeRecord(f); f.decks[NEXT].slides[20].getPageElements().push(shape('concurrent-extra', 'extra', { left: 8, top: 8, width: 8, height: 8 })); assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /fingerprint/); }
  { const f = fixture(); completeRecord(f); f.decks[NEXT].slides[0].getPageElements().pop(); assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /fingerprint/); }
  { const f = fixture(); completeRecord(f); const body = JSON.parse(record(f).content); body.newDeckFingerprint.sha256 = '0'.repeat(64); record(f).content = JSON.stringify(body); assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /fingerprint/); }
});

test('audits fail closed for substituted IDs, stale notes, uncertain records, and unrecorded decks', () => {
  { const f = fixture(); f.mapFile.getId = () => 'other-map'; assert.throws(() => f.G.隔離通し_J台本更新後の再組版前を検査する(), /対応表ID/); }
  { const f = fixture(); f.decks[PRIOR].slides[6].setNote('not-the-snapshot'); assert.throws(() => f.G.隔離通し_J台本更新後の再組版前を検査する(), /notes/); }
  { const f = fixture(); completeRecord(f); const body = JSON.parse(record(f).content); body.status = 'uncertain'; record(f).content = JSON.stringify(body); assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /completed実行記録/); }
  { const f = fixture(); completeRecord(f); const body = JSON.parse(record(f).content); body.newDeckId = 'unrecorded'; record(f).content = JSON.stringify(body); assert.throws(() => f.G.隔離通し_J台本更新後の再組版を読取検品する(), /hash|deck|対応表|completed/); }
});

test('source has no deck argument, no timestamp equality, explicit image limitation, and read-only audits', () => {
  const source = require('fs').readFileSync(require('path').join(__dirname, '..', '59_j_regeneration_audit.js'), 'utf8');
  const pre = source.slice(source.indexOf('function 隔離通し_J台本更新後の再組版前を検査する'), source.indexOf('function isolatedJRegenerationRecord_'));
  const writer = source.slice(source.indexOf('function 隔離通し_J台本更新後のスライドを一度だけ組み直す'), source.indexOf('function 隔離通し_J台本更新後の再組版を読取検品する'));
  const post = source.slice(source.indexOf('function 隔離通し_J台本更新後の再組版を読取検品する'));
  assert.doesNotMatch(pre, /setContent\s*\(|createFile\s*\(|スライドを組む_\s*\(/); assert.doesNotMatch(post, /setContent\s*\(|createFile\s*\(|スライドを組む_\s*\(/); assert.match(writer, /withIsolationTestConfig_\s*\(\s*target\.cfg\s*,/);
  assert.doesNotMatch(writer, /processNewFlow_|知らせを積む_|止まったを知らせる_/); assert.doesNotMatch(source, /getLastUpdated\(\)\.toISOString\(\)\s*!==\s*String\(map\.presentationRevision/); assert.match(source, /image-pixels-and-dynamic-content-url-not-readable-via-SlidesApp/); assert.match(source, /newDeckFingerprint/);
  assert.doesNotMatch(source, /function\s+隔離通し_J台本更新後の再組版前を検査する\s*\([^)]*[^\s)]/); assert.doesNotMatch(source, /function\s+隔離通し_J台本更新後の再組版を読取検品する\s*\([^)]*[^\s)]/);
});

console.log('pass ' + passed);
