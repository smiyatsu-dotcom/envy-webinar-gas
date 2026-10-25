// Fixed-fixture contract for the one J No.7 regeneration.  The preflight
// intentionally accepts the prior deck because it must still contain the
// pre-correction script.  The post audit accepts only the new deck ID written
// by the guarded no-argument wrapper below; callers can never provide IDs.
const ISOLATED_J_REGEN_AUDIT_ = {
  caseName: '隔離通し診断テスト社',
  folderId: '1Ko7cbJ2nW2C4uQbZfDnZjW6mfz_qmHZ0',
  sourceFileId: '1g08BmnhKnzdtepHhUNFkAr2mhjww2cnC',
  kouseiSheetId: '1YvJszV8Mz1tN1DVZ9xOIY5D7XGW754QQFVU_oyi34AQ',
  mapFileId: '1DukH9lGhK0T5W-PmkBBWUCVqUyJOwUtc',
  mapFileName: '【スライド対応】隔離通し診断テスト社.json',
  geometryFileName: '【配置指定】隔離通し診断テスト社.json',
  sourceFileName: '【台本JSON】隔離通し診断テスト社.txt',
  importRecordFileName: '【不確定】J台本取込記録.json',
  writerRecordFileName: '【不確定】No7挙手文言修正記録.json',
  writerSnapshotFileName: '【保全】No7挙手文言修正前台本.txt',
  regenerationRecordFileName: '【不確定】J_No7再組版記録.json',
  sourceSha256Before: '7eb361136315b71a1505d50e94cbeef12d28feb1846631ec67370072e5331645',
  sourceSha256After: 'aa136a10184d9c85799a6b80b24e8ec304309f88d8508cc33d61b3a0cda92b7e',
  other105Sha256: '3b0e762d22d218c432ea432c0e3bfd82161502e049611e5abc6722311d8ed3f3',
  oldNo7: '例えば、皆さん経営していてこんなお悩みはありませんか？\nもし、当てはまるものがあれば、挙手🖐️マークでリアクションしていただけると嬉しいです！',
  newNo7: '例えば、皆さん経営していてこんなお悩みはありませんか？\n当てはまる方は、挙手マークを押してください。',
  sheetCell: '問題教育!E2'
};

function isolatedJRegenOneFile_(folder, name, label) {
  const it = folder.getFilesByName(name);
  if (!it.hasNext()) throw new Error(label + ' が固定隔離案件にありません');
  const file = it.next();
  if (it.hasNext()) throw new Error(label + ' が複数あります');
  return file;
}
function isolatedJRegenHash_(text) { return 隔離通し_J台本_sha256_(text); }
function isolatedJRegenRows_(text, label) {
  let rows;
  try { rows = JSON.parse(String(text)); } catch (e) { throw new Error(label + 'をJSONとして読めません'); }
  if (!Array.isArray(rows) || rows.length !== 106) throw new Error(label + 'は106行の配列ではありません');
  const seen = {};
  rows.forEach(function(row) {
    const no = Number(row && row.no);
    if (!row || typeof row !== 'object' || Array.isArray(row) || Object.keys(row).sort().join(',') !== 'daihon,no' ||
        !Number.isInteger(no) || no < 1 || no > 106 || seen[no] || typeof row.daihon !== 'string' || !row.daihon.trim()) {
      throw new Error(label + 'のNo.または本文が不正です');
    }
    seen[no] = true;
  });
  return rows;
}
function isolatedJRegenFolder_(cfg) {
  const c = ISOLATED_J_REGEN_AUDIT_, output = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  if (cfg.ISOLATION_TEST !== true || output.getId() !== 隔離通し_生成_生成物rootID || output.getName() !== '【出力】生成物') throw new Error('J再組版監査の生成物rootが固定fixtureと一致しません');
  const folder = caseFolderIfExists_(cfg, c.caseName);
  if (!folder || folder.getId() !== c.folderId || folder.getName() !== c.caseName) throw new Error('J再組版監査の案件フォルダが固定fixtureと一致しません');
  隔離通し_生成の親を確かめる_(folder, output.getId(), 'J再組版監査の案件フォルダ');
  return folder;
}
function isolatedJRegenLedger_(cfg) {
  const rows = mgmtSheet_(cfg).getDataRange().getValues(), nonEmpty = rows.slice(1).filter(function(row) { return row.some(function(cell) { return String(cell == null ? '' : cell).trim() !== ''; }); });
  const hits = [], sameHearing = [];
  rows.slice(1).forEach(function(row, index) {
    const exact = String(row[0] || '').trim() === 隔離通し_生成_ヒアリングID && String(row[1] || '').trim() === ISOLATED_J_REGEN_AUDIT_.caseName + '_ヒアリングシート' && String(row[流れ_列 - 1] || '').trim() === 流れ_新;
    if (exact) hits.push({ row: row, rowNumber: index + 2 });
    if (String(row[0] || '').trim() === 隔離通し_生成_ヒアリングID) sameHearing.push(index + 2);
  });
  if (nonEmpty.length !== 1 || hits.length !== 1 || sameHearing.length !== 1 || registry_(cfg)[隔離通し_生成_ヒアリングID] !== hits[0].rowNumber) throw new Error('J再組版監査の固定台帳行またはregistryが一意ではありません');
  const id = idFromUrl_(String(hits[0].row[3] || ''));
  if (id !== ISOLATED_J_REGEN_AUDIT_.kouseiSheetId) throw new Error('J再組版監査の構成Sheet IDが固定対象と一致しません');
  const ss = SpreadsheetApp.openById(id), settings = ss.getSheetByName(TABS.SETTINGS), problem = ss.getSheetByName(TABS.PROBLEM);
  if (!settings || !problem || String(settings.getRange('B1').getValue()).trim() !== ISOLATED_J_REGEN_AUDIT_.caseName + '_ヒアリングシート' || String(settings.getRange('B4').getValue()).trim() !== 隔離通し_生成_ヒアリングID || settings.getRange('B2').getValue() !== false || settings.getRange('B22').getValue() !== false || String(problem.getRange('E2').getValue() || '') !== ISOLATED_J_REGEN_AUDIT_.newNo7) throw new Error('J再組版監査の固定構成SheetまたはNo.7セルが一致しません');
  return { row: hits[0], kouseiId: id };
}
function isolatedJRegenSourceChain_(cfg, folder) {
  const c = ISOLATED_J_REGEN_AUDIT_, script = isolatedJRegenOneFile_(folder, c.sourceFileName, 'J再組版元台本JSON');
  if (script.getId() !== c.sourceFileId) throw new Error('J再組版元台本JSON IDが固定対象と一致しません');
  const sourceText = script.getBlob().getDataAsString('UTF-8');
  if (isolatedJRegenHash_(sourceText) !== c.sourceSha256After) throw new Error('J再組版元台本JSONがNo.7更新後SHAと一致しません');
  const rows = isolatedJRegenRows_(sourceText, 'J再組版元台本JSON'), no7 = rows.filter(function(row) { return Number(row.no) === 7; });
  if (no7.length !== 1 || no7[0].daihon !== c.newNo7 || isolatedJRegenHash_(JSON.stringify(rows.filter(function(row) { return Number(row.no) !== 7; }))) !== c.other105Sha256) throw new Error('J再組版元台本JSONのNo.7または他105行hashが一致しません');
  const importRecord = isolatedJRegenOneFile_(folder, c.importRecordFileName, 'J台本取込記録');
  const snapshot = isolatedJRegenOneFile_(folder, c.writerSnapshotFileName, 'No.7修正前保全');
  const writerRecord = isolatedJRegenOneFile_(folder, c.writerRecordFileName, 'No.7修正記録');
  let imported, written;
  try { imported = JSON.parse(importRecord.getBlob().getDataAsString('UTF-8')); written = JSON.parse(writerRecord.getBlob().getDataAsString('UTF-8')); }
  catch (e) { throw new Error('J取込またはNo.7 writer記録JSONを読めません'); }
  if (!imported || imported.status !== 'completed' || imported.automaticRetry !== false || imported.folderId !== c.folderId || imported.targetScriptSha256 !== c.sourceSha256Before ||
      isolatedJRegenHash_(snapshot.getBlob().getDataAsString('UTF-8')) !== c.sourceSha256Before ||
      !written || written.schema !== 'isolation-j-no7-hand-raise/v1' || written.status !== 'completed' || written.automaticRetry !== false || written.folderId !== c.folderId || written.sourceFileId !== c.sourceFileId || written.kouseiSheetId !== c.kouseiSheetId || written.managementRow !== isolatedJRegenLedger_(cfg).row.rowNumber || written.sourceSha256Before !== c.sourceSha256Before || written.sourceSha256After !== c.sourceSha256After || written.other105Sha256 !== c.other105Sha256 || written.backupFileId !== snapshot.getId() || written.sheetCell !== c.sheetCell) {
    throw new Error('J再組版のcompleted取込/No.7 writer chainが一致しません');
  }
  return { rows: rows, snapshotRows: isolatedJRegenRows_(snapshot.getBlob().getDataAsString('UTF-8'), 'J再組版の更新前台本'), sourceSha256: c.sourceSha256After };
}
function isolatedJRegenMapDeck_(folder, expectedRows, expectedMapSha256, expectedDeckId, phase) {
  const c = ISOLATED_J_REGEN_AUDIT_, mapFile = isolatedJRegenOneFile_(folder, c.mapFileName, 'J再組版スライド対応表');
  if (mapFile.getId() !== c.mapFileId) throw new Error('J再組版スライド対応表IDが固定対象と一致しません');
  const mapText = mapFile.getBlob().getDataAsString('UTF-8');
  if (expectedMapSha256 && isolatedJRegenHash_(mapText) !== expectedMapSha256) throw new Error('J再組版スライド対応表hashが実行記録と一致しません');
  let map;
  try { map = JSON.parse(mapText); } catch (e) { throw new Error('J再組版スライド対応表JSONを読めません'); }
  if (!map || map.schema !== GEOMETRY_MAP_SCHEMA || map.caseName !== c.caseName || map.presentationId !== expectedDeckId || !Array.isArray(map.pages) || map.pages.length !== 106 || map.pages.some(function(page, index) { return !page || page.slideNo !== index + 1 || typeof page.pageId !== 'string'; }) || new Set(map.pages.map(function(page) { return page.pageId; })).size !== 106) throw new Error('J再組版スライド対応表が固定106ページ/実行記録と一致しません');
  const deckFile = DriveApp.getFileById(expectedDeckId), parents = deckFile.getParents();
  if (deckFile.isTrashed() || deckFile.getMimeType() !== MimeType.GOOGLE_SLIDES || !parents.hasNext() || parents.next().getId() !== folder.getId() || parents.hasNext() || String(deckFile.getLastUpdated().toISOString()) !== String(map.presentationRevision || '')) throw new Error('J再組版deckの親・形式・revisionが対応表と一致しません');
  const slides = SlidesApp.openById(expectedDeckId).getSlides(), byId = {};
  slides.forEach(function(slide) { const id = String(slide.getObjectId() || ''); if (!id || byId[id]) throw new Error('J再組版deckのpage IDが一意ではありません'); byId[id] = slide; });
  if (slides.length !== 106 || map.pages.some(function(page) { return !byId[page.pageId]; })) throw new Error('J再組版deckが対応表の106ページと一致しません');
  const expected = isolatedJCurrentNormalGeneratorExpected_(expectedRows).expectedByNo;
  // The existing deck was generated before the No.7 correction, so only this
  // preflight comparison uses the retained snapshot wording at No.7.  Fixed
  // 105/106 precedence remains the normal generator's current rule.
  if (phase === 'before') expected[7] = ISOLATED_J_REGEN_AUDIT_.oldNo7;
  const mismatches = [];
  map.pages.forEach(function(page) { const note = isolatedJCurrentNote_(byId[page.pageId]); if (!note.readable || note.text !== expected[page.slideNo]) mismatches.push(page.slideNo); });
  if (mismatches.length) throw new Error('J再組版deckのnotesが期待台本と一致しません: ' + mismatches.join(','));
  const geometryFile = isolatedJRegenOneFile_(folder, c.geometryFileName, 'J再組版配置指定');
  let geometry;
  try { geometry = JSON.parse(geometryFile.getBlob().getDataAsString('UTF-8')); } catch (e) { throw new Error('J再組版配置指定JSONを読めません'); }
  if (!geometry || geometry.schema !== GEOMETRY_SCHEMA || geometry.caseName !== c.caseName || !Array.isArray(geometry.records)) throw new Error('J再組版配置指定が固定案件と一致しません');
  const active = geometry.records.filter(function(record) { return record && record.status === 'active'; });
  const retained = active.map(function(record) {
    const page = map.pages[Number(record.slideNo) - 1], selector = page && page.targetSelector, elements = page ? byId[page.pageId].getPageElements() : [];
    const matches = selector ? elements.filter(function(element) { try { return selectorMatches_(element, selector); } catch (e) { return false; } }) : [];
    return !!selector && validGeometry_(record.after) && sameGeometry_(selector.templateGeometry, record.after) && matches.length === 1;
  });
  if (active.length && retained.some(function(ok) { return !ok; })) throw new Error('J再組版deckの永続geometryが対応表どおりではありません');
  return { mapFile: mapFile, mapText: mapText, map: map, deckFile: deckFile, slideCount: slides.length, geometryCount: active.length, notesSha256: isolatedJRegenHash_(JSON.stringify(map.pages.map(function(page) { return { no: page.slideNo, note: isolatedJCurrentNote_(byId[page.pageId]).text }; }))) };
}
function isolatedJRegenerationContext_() {
  const cfg = 隔離通し_生成設定を作る_(), folder = isolatedJRegenFolder_(cfg), ledger = isolatedJRegenLedger_(cfg), chain = isolatedJRegenSourceChain_(cfg, folder);
  return { cfg: cfg, folder: folder, ledger: ledger, chain: chain };
}
function isolatedJRegenerationPreflight_() {
  const context = isolatedJRegenerationContext_(), cfg = context.cfg, folder = context.folder, ledger = context.ledger, chain = context.chain;
  const mapFile = isolatedJRegenOneFile_(folder, ISOLATED_J_REGEN_AUDIT_.mapFileName, 'J再組版更新前対応表');
  const before = isolatedJRegenMapDeck_(folder, chain.snapshotRows, null, JSON.parse(mapFile.getBlob().getDataAsString('UTF-8')).presentationId, 'before');
  return { cfg: cfg, folder: folder, ledger: ledger, chain: chain, before: before,
    beforeMapSha256: isolatedJRegenHash_(before.mapText), beforeDeckId: before.map.presentationId };
}
function 隔離通し_J台本更新後の再組版前を検査する() {
  const target = isolatedJRegenerationPreflight_();
  return { ok: true, mode: 'read-only-before-no7-regeneration', sourceSha256: target.chain.sourceSha256,
    previousDeckId: target.beforeDeckId, mapFileId: target.before.mapFile.getId(), previousMapSha256: target.beforeMapSha256,
    slideCount: target.before.slideCount, geometryCount: target.before.geometryCount,
    actions: { regenerated: false, imported: false, audioStarted: false, delivered: false, notified: false } };
}
function isolatedJRegenerationRecord_(folder) {
  const c = ISOLATED_J_REGEN_AUDIT_, it = folder.getFilesByName(c.regenerationRecordFileName);
  if (!it.hasNext()) return null;
  const file = it.next(); if (it.hasNext()) throw new Error('J再組版実行記録が複数あります');
  return file;
}
// The only writer for this contract.  It runs the existing fixed normal
// builder once, then records the returned map/deck identity for a separate
// read-only post audit.  Any exception becomes uncertain and forbids retry.
function 隔離通し_J台本更新後のスライドを一度だけ組み直す() {
  const target = isolatedJRegenerationPreflight_(), c = ISOLATED_J_REGEN_AUDIT_;
  return withIsolationTestConfig_(target.cfg, function() {
    if (isolatedJRegenerationRecord_(target.folder)) throw new Error('J再組版実行記録が既にあります。自動再実行せず手動精査してください');
    const record = target.folder.createFile(c.regenerationRecordFileName, JSON.stringify({ schema: 'isolation-j-no7-regeneration/v1', status: 'prepared', automaticRetry: false, folderId: c.folderId, sourceSha256: target.chain.sourceSha256, mapFileId: c.mapFileId, previousDeckId: target.beforeDeckId, previousMapSha256: target.beforeMapSha256, createdAt: new Date().toISOString() }), MimeType.PLAIN_TEXT);
    try {
      const built = スライドを組む_(target.cfg, target.ledger.row.row, false);
      if (built !== true) throw new Error('固定隔離の正規スライド再組版が完了しませんでした: ' + built);
      const mapFile = isolatedJRegenOneFile_(target.folder, c.mapFileName, 'J再組版後対応表');
      const map = JSON.parse(mapFile.getBlob().getDataAsString('UTF-8'));
      if (!map || typeof map.presentationId !== 'string' || !map.presentationId || map.presentationId === target.beforeDeckId) throw new Error('J再組版後の新deck IDを対応表から確定できません');
      const after = isolatedJRegenMapDeck_(target.folder, target.chain.rows, null, map.presentationId, 'after'), afterMapSha256 = isolatedJRegenHash_(after.mapText);
      record.setContent(JSON.stringify({ schema: 'isolation-j-no7-regeneration/v1', status: 'completed', automaticRetry: false, folderId: c.folderId, sourceSha256: target.chain.sourceSha256, mapFileId: c.mapFileId, previousDeckId: target.beforeDeckId, previousMapSha256: target.beforeMapSha256, newDeckId: map.presentationId, newMapSha256: afterMapSha256, newDeckRevision: map.presentationRevision, completedAt: new Date().toISOString() }));
      return { ok: true, mode: 'fixed-no7-normal-regeneration', newDeckId: map.presentationId, mapFileId: c.mapFileId, newMapSha256: afterMapSha256, actions: { regenerated: true, imported: false, audioStarted: false, delivered: false, notified: false } };
    } catch (e) {
      record.setContent(JSON.stringify({ schema: 'isolation-j-no7-regeneration/v1', status: 'uncertain', automaticRetry: false, folderId: c.folderId, sourceSha256: target.chain.sourceSha256, mapFileId: c.mapFileId, previousDeckId: target.beforeDeckId, previousMapSha256: target.beforeMapSha256, failure: String(e && e.message ? e.message : e).slice(0, 240), recordedAt: new Date().toISOString() }));
      throw new Error('J再組版の結果が不確定です。自動再実行せず実行記録を手動精査してください');
    }
  });
}
function 隔離通し_J台本更新後の再組版を読取検品する() {
  const target = isolatedJRegenerationContext_(), c = ISOLATED_J_REGEN_AUDIT_, recordFile = isolatedJRegenerationRecord_(target.folder);
  if (!recordFile) throw new Error('J再組版実行記録がありません');
  let record;
  try { record = JSON.parse(recordFile.getBlob().getDataAsString('UTF-8')); } catch (e) { throw new Error('J再組版実行記録JSONを読めません'); }
  if (!record || record.schema !== 'isolation-j-no7-regeneration/v1' || record.status !== 'completed' || record.automaticRetry !== false || record.folderId !== c.folderId || record.sourceSha256 !== c.sourceSha256After || record.mapFileId !== c.mapFileId || typeof record.previousDeckId !== 'string' || typeof record.previousMapSha256 !== 'string' || typeof record.newDeckId !== 'string' || !record.newDeckId || record.newDeckId === record.previousDeckId || typeof record.newMapSha256 !== 'string' || typeof record.newDeckRevision !== 'string') throw new Error('J再組版completed実行記録が固定契約と一致しません');
  const after = isolatedJRegenMapDeck_(target.folder, target.chain.rows, record.newMapSha256, record.newDeckId, 'after');
  if (after.map.presentationRevision !== record.newDeckRevision) throw new Error('J再組版後map revisionが実行記録と一致しません');
  return { ok: true, mode: 'read-only-after-no7-regeneration', newDeckId: record.newDeckId, mapFileId: c.mapFileId, mapSha256: record.newMapSha256, slideCount: after.slideCount, geometryCount: after.geometryCount, notesSha256: after.notesSha256, actions: { regenerated: false, imported: false, audioStarted: false, delivered: false, notified: false } };
}
