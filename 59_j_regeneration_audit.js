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
  // This isolation fixture deliberately retains No.104 as the optional
  // skipped slide.  It is part of the content contract, not a UI preference.
  fixedSkippedSlideNos: [104],
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
// generateSlides_ reads the regular structure-sheet rows, not the imported
// JSON directly.  The JSON chain above still proves the No.7 writer source;
// this separate read-only projection proves the field/type actually supplied
// to the normal builder, including the regular skip decision.
function isolatedJRegenRegularGeneratorExpected_(kouseiId) {
  const data = readKouseiSheet_(kouseiId);
  if (!data || !data.rows || typeof data.rows !== 'object') throw new Error('J再組版の固定構成Sheetを読めません');
  const defs = rowDefs_(), expectedByNo = {}, skipByNo = {};
  if (!Array.isArray(defs) || defs.length !== 106) throw new Error('J再組版の生成定義が106行ありません');
  defs.forEach(function(def) {
    const no = Number(def && def.no), row = data.rows[no];
    if (!Number.isInteger(no) || no < 1 || no > 106 || Object.prototype.hasOwnProperty.call(expectedByNo, no) || !row || typeof row !== 'object') throw new Error('J再組版の構成Sheetが生成定義1〜106と一致しません');
    const skipped = isSkipRow_(def, row);
    skipByNo[no] = skipped;
    expectedByNo[no] = skipped ? '' : isolatedJCurrentNormalizeLineEndings_(normalizeWebinarTerms_(row.daihon || ''));
  });
  return { expectedByNo: expectedByNo, skipByNo: skipByNo };
}
// Drive v3 and DriveApp can legitimately report a slightly different modified
// timestamp for the same Slides file.  A timestamp therefore remains useful
// diagnostic evidence, but is not a content or revision proof.  The contract
// below instead hashes every page and every page element that SlidesApp can
// read.  Textual element contents are included in the hash; images and other
// non-textual elements are *not* silently skipped: their identity, label and
// geometry are hashed and the returned coverage explicitly says that their
// pixels/external URLs cannot be read by SlidesApp.
function isolatedJRegenElementText_(element, type) {
  if (type === 'SHAPE') return { verification: 'text', kind: 'shape-text', value: String(element.asShape().getText().asString()) };
  if (type === 'WORD_ART') return { verification: 'text', kind: 'word-art', value: String(element.asWordArt().getRenderedText()) };
  if (type === 'TABLE') {
    const table = element.asTable(), cells = [];
    for (let r = 0; r < table.getNumRows(); r++) for (let c = 0; c < table.getNumColumns(); c++) cells.push({ row: r, column: c, value: String(table.getCell(r, c).getText().asString()) });
    return { verification: 'text', kind: 'table-cells', cells: cells };
  }
  return null;
}
function isolatedJRegenElementFingerprint_(element, coverage) {
  const type = String(element.getPageElementType()), objectId = String(element.getObjectId() || '');
  if (!objectId) throw new Error('J再組版deckの要素object IDを取得できません');
  const item = { objectId: objectId, type: type, label: elementLabel_(element), geometry: geometryOf_(element) };
  if (type === 'GROUP') {
    let children;
    try { children = element.asGroup().getChildren(); } catch (e) { throw new Error('J再組版deckのGROUP要素を読めません'); }
    if (!Array.isArray(children)) throw new Error('J再組版deckのGROUP子要素を配列として読めません');
    item.content = { verification: 'structure', kind: 'group-children', children: children.map(function(child) { return isolatedJRegenElementFingerprint_(child, coverage); }) };
    coverage.structuredElements++;
    return item;
  }
  let text;
  try { text = isolatedJRegenElementText_(element, type); }
  catch (e) { throw new Error('J再組版deckの' + type + '本文を読めません'); }
  if (text) {
    item.content = text;
    coverage.textElements++;
    return item;
  }
  // SlidesApp exposes neither image bytes nor a stable source/content URL for
  // these types.  Keep their metadata in the fingerprint and disclose this
  // residual limitation instead of pretending image content was verified.
  const reason = type === 'IMAGE' ? 'SlidesApp cannot read image pixels or a stable content URL' : 'SlidesApp exposes no stable readable content for this page-element type';
  item.content = { verification: 'metadata-only', kind: 'unreadable-element', reason: reason };
  coverage.metadataOnly[type] = (coverage.metadataOnly[type] || 0) + 1;
  return item;
}
function isolatedJRegenDeckFingerprint_(slides, map) {
  const byId = {};
  slides.forEach(function(slide) {
    const id = String(slide.getObjectId() || '');
    if (!id || byId[id]) throw new Error('J再組版deckのpage IDが一意ではありません');
    byId[id] = slide;
  });
  const coverage = { textElements: 0, structuredElements: 0, metadataOnly: {} };
  const pages = map.pages.map(function(page) {
    const slide = byId[page.pageId];
    if (!slide) throw new Error('J再組版deckが対応表のpage IDを持ちません');
    const note = isolatedJCurrentNote_(slide);
    if (!note.readable) throw new Error('J再組版deckのnotesを読めません');
    const elements = slide.getPageElements();
    if (!Array.isArray(elements)) throw new Error('J再組版deckのページ要素を配列として読めません');
    let skipped;
    try { skipped = slide.isSkipped() === true; } catch (e) { throw new Error('J再組版deckのスキップ状態を読めません'); }
    // Preserve the representation returned by SlidesApp here.  This contract
    // permits the one paragraph LF we construct below, but must not make a
    // CRLF or any other raw storage difference disappear before fingerprinting.
    return { slideNo: page.slideNo, pageId: page.pageId, skipped: skipped, note: note.rawText,
      elements: elements.map(function(element) { return isolatedJRegenElementFingerprint_(element, coverage); }) };
  });
  const content = { schema: 'isolation-j-deck-fingerprint/v1', pages: pages };
  const metadataOnly = Object.keys(coverage.metadataOnly).sort().map(function(type) {
    return { type: type, count: coverage.metadataOnly[type], limitation: type === 'IMAGE' ? 'image-pixels-and-dynamic-content-url-not-readable-via-SlidesApp' : 'element-content-not-readable-via-SlidesApp' };
  });
  const pageOrder = pages.map(function(page) { return page.pageId; });
  const skippedSlideNos = pages.filter(function(page) { return page.skipped; }).map(function(page) { return page.slideNo; });
  const notes = pages.map(function(page) { return { no: page.slideNo, note: page.note }; });
  const elements = pages.map(function(page) { return { no: page.slideNo, elements: page.elements }; });
  return { schema: content.schema, sha256: isolatedJRegenHash_(JSON.stringify(content)),
    pageOrderSha256: isolatedJRegenHash_(JSON.stringify(pageOrder)), skipStatesSha256: isolatedJRegenHash_(JSON.stringify(skippedSlideNos)),
    skippedSlideNos: skippedSlideNos, notesSha256: isolatedJRegenHash_(JSON.stringify(notes)),
    elementsSha256: isolatedJRegenHash_(JSON.stringify(elements)), pageCount: pages.length,
    elementCount: pages.reduce(function(sum, page) { return sum + page.elements.length; }, 0),
    textElements: coverage.textElements, structuredElements: coverage.structuredElements,
    metadataOnlyElements: metadataOnly.reduce(function(sum, item) { return sum + item.count; }, 0), metadataOnly: metadataOnly };
}
function isolatedJRegenSameFingerprint_(left, right) {
  const fields = ['schema', 'sha256', 'pageOrderSha256', 'skipStatesSha256', 'notesSha256', 'elementsSha256', 'pageCount', 'elementCount', 'textElements', 'structuredElements', 'metadataOnlyElements'];
  return !!left && !!right && fields.every(function(field) { return left[field] === right[field]; }) && JSON.stringify(left.skippedSlideNos) === JSON.stringify(right.skippedSlideNos) && JSON.stringify(left.metadataOnly) === JSON.stringify(right.metadataOnly);
}
function isolatedJRegenValidFingerprint_(fingerprint) {
  const hashes = ['sha256', 'pageOrderSha256', 'skipStatesSha256', 'notesSha256', 'elementsSha256'];
  const counts = ['pageCount', 'elementCount', 'textElements', 'structuredElements', 'metadataOnlyElements'];
  return !!fingerprint && fingerprint.schema === 'isolation-j-deck-fingerprint/v1' &&
    hashes.every(function(field) { return typeof fingerprint[field] === 'string' && /^[0-9a-f]{64}$/.test(fingerprint[field]); }) &&
    counts.every(function(field) { return Number.isInteger(fingerprint[field]) && fingerprint[field] >= 0; }) &&
    fingerprint.pageCount === 106 && Array.isArray(fingerprint.skippedSlideNos) && Array.isArray(fingerprint.metadataOnly) &&
    fingerprint.skippedSlideNos.every(function(no) { return Number.isInteger(no) && no >= 1 && no <= 106; }) &&
    fingerprint.metadataOnly.every(function(item) { return !!item && typeof item.type === 'string' && Number.isInteger(item.count) && item.count > 0 && typeof item.limitation === 'string'; });
}
// SlidesApp TextRange.asString() represents the terminal paragraph that
// setText() creates as exactly one final LF, including for an empty note.  It
// is a storage boundary, not a license to trim: source text may not already
// end in a paragraph separator, and all internal LF/whitespace remains exact.
function isolatedJRegenStoredNote_(semanticText) {
  const text = isolatedJCurrentNormalizeLineEndings_(semanticText);
  if (text.endsWith('\n')) throw new Error('J再組版の期待note本文が終端paragraph改行を含みます');
  return text + '\n';
}
function isolatedJRegenMapDeck_(folder, expectedByNo, expectedMapSha256, expectedDeckId, phase) {
  const c = ISOLATED_J_REGEN_AUDIT_, mapFile = isolatedJRegenOneFile_(folder, c.mapFileName, 'J再組版スライド対応表');
  if (mapFile.getId() !== c.mapFileId) throw new Error('J再組版スライド対応表IDが固定対象と一致しません');
  const mapText = mapFile.getBlob().getDataAsString('UTF-8');
  if (expectedMapSha256 && isolatedJRegenHash_(mapText) !== expectedMapSha256) throw new Error('J再組版スライド対応表hashが実行記録と一致しません');
  let map;
  try { map = JSON.parse(mapText); } catch (e) { throw new Error('J再組版スライド対応表JSONを読めません'); }
  if (!map || map.schema !== GEOMETRY_MAP_SCHEMA || map.caseName !== c.caseName || map.presentationId !== expectedDeckId || !Array.isArray(map.pages) || map.pages.length !== 106 || map.pages.some(function(page, index) { return !page || page.slideNo !== index + 1 || typeof page.pageId !== 'string'; }) || new Set(map.pages.map(function(page) { return page.pageId; })).size !== 106) throw new Error('J再組版スライド対応表が固定106ページ/実行記録と一致しません');
  const deckFile = DriveApp.getFileById(expectedDeckId), parents = deckFile.getParents();
  if (deckFile.isTrashed() || deckFile.getMimeType() !== MimeType.GOOGLE_SLIDES || !parents.hasNext() || parents.next().getId() !== folder.getId() || parents.hasNext()) throw new Error('J再組版deckの親・形式が対応表と一致しません');
  if (typeof map.presentationRevision !== 'string' || !map.presentationRevision) throw new Error('J再組版スライド対応表のpresentationRevisionがありません');
  const timestampDiagnostic = { mapPresentationRevision: map.presentationRevision, deckLastUpdated: String(deckFile.getLastUpdated().toISOString()) };
  const slides = SlidesApp.openById(expectedDeckId).getSlides(), byId = {};
  slides.forEach(function(slide) { const id = String(slide.getObjectId() || ''); if (!id || byId[id]) throw new Error('J再組版deckのpage IDが一意ではありません'); byId[id] = slide; });
  if (slides.length !== 106 || map.pages.some(function(page) { return !byId[page.pageId]; })) throw new Error('J再組版deckが対応表の106ページと一致しません');
  const fixedSkipped = c.fixedSkippedSlideNos;
  if (!Array.isArray(fixedSkipped) || new Set(fixedSkipped).size !== fixedSkipped.length || fixedSkipped.some(function(no) { return !Number.isInteger(no) || no < 1 || no > 106; })) throw new Error('J再組版の固定スキップ定義が不正です');
  const skipMismatches = map.pages.filter(function(page) {
    let actual;
    try { actual = byId[page.pageId].isSkipped() === true; } catch (e) { throw new Error('J再組版deckのスキップ状態を読めません'); }
    return actual !== (fixedSkipped.indexOf(page.slideNo) !== -1);
  }).map(function(page) { return page.slideNo; });
  if (skipMismatches.length) throw new Error('J再組版deckの固定スキップ状態が一致しません: ' + skipMismatches.join(','));
  if (!expectedByNo || typeof expectedByNo !== 'object' || Object.keys(expectedByNo).length !== 106 || Object.keys(expectedByNo).some(function(key) { const no = Number(key); return !Number.isInteger(no) || no < 1 || no > 106 || typeof expectedByNo[key] !== 'string'; })) throw new Error('J再組版の通常生成期待本文が106行ではありません');
  const expectedSemantic = {};
  Object.keys(expectedByNo).forEach(function(no) { expectedSemantic[Number(no)] = expectedByNo[no]; });
  // The existing deck was generated before the No.7 correction, so only this
  // preflight comparison uses the retained snapshot wording at No.7.  Fixed
  // 105/106 precedence remains the normal generator's current rule.
  if (phase === 'before') expectedSemantic[7] = ISOLATED_J_REGEN_AUDIT_.oldNo7;
  fixedSkipped.forEach(function(no) { expectedSemantic[no] = ''; });
  const expected = {};
  Object.keys(expectedSemantic).forEach(function(no) { expected[Number(no)] = isolatedJRegenStoredNote_(expectedSemantic[no]); });
  const mismatches = [];
  map.pages.forEach(function(page) {
    const note = isolatedJCurrentNote_(byId[page.pageId]);
    // Compare the raw SlidesApp value.  Existing broad line-ending
    // normalization is appropriate for historical diagnostics, but would
    // incorrectly hide CRLF here; this release contract accepts only the
    // exact terminal LF produced by setText().
    if (!note.readable || note.rawText !== expected[page.slideNo]) mismatches.push(page.slideNo);
  });
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
  const fingerprint = isolatedJRegenDeckFingerprint_(slides, map);
  return { mapFile: mapFile, mapText: mapText, map: map, deckFile: deckFile, slideCount: slides.length, geometryCount: active.length,
    notesSha256: fingerprint.notesSha256, fingerprint: fingerprint, timestampDiagnostic: timestampDiagnostic };
}
function isolatedJRegenerationContext_() {
  const cfg = 隔離通し_生成設定を作る_(), folder = isolatedJRegenFolder_(cfg), ledger = isolatedJRegenLedger_(cfg), chain = isolatedJRegenSourceChain_(cfg, folder), generator = isolatedJRegenRegularGeneratorExpected_(ledger.kouseiId);
  return { cfg: cfg, folder: folder, ledger: ledger, chain: chain, generator: generator };
}
function isolatedJRegenerationPreflight_() {
  const context = isolatedJRegenerationContext_(), cfg = context.cfg, folder = context.folder, ledger = context.ledger, chain = context.chain;
  const mapFile = isolatedJRegenOneFile_(folder, ISOLATED_J_REGEN_AUDIT_.mapFileName, 'J再組版更新前対応表');
  const before = isolatedJRegenMapDeck_(folder, context.generator.expectedByNo, null, JSON.parse(mapFile.getBlob().getDataAsString('UTF-8')).presentationId, 'before');
  return { cfg: cfg, folder: folder, ledger: ledger, chain: chain, generator: context.generator, before: before,
    beforeMapSha256: isolatedJRegenHash_(before.mapText), beforeDeckId: before.map.presentationId };
}
function 隔離通し_J台本更新後の再組版前を検査する() {
  const target = isolatedJRegenerationPreflight_();
  return { ok: true, mode: 'read-only-before-no7-regeneration', sourceSha256: target.chain.sourceSha256,
    previousDeckId: target.beforeDeckId, mapFileId: target.before.mapFile.getId(), previousMapSha256: target.beforeMapSha256,
    previousDeckFingerprint: target.before.fingerprint, timestampDiagnostic: target.before.timestampDiagnostic,
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
    // Re-read immediately before any writer is created.  This is not a lock
    // against Drive's whole world, but it catches a map/deck replacement that
    // happened after the initial preflight; the completed record then anchors
    // the exact post-build map and deck for a separate read-only comparison.
    const confirmed = isolatedJRegenMapDeck_(target.folder, target.generator.expectedByNo, target.beforeMapSha256, target.beforeDeckId, 'before');
    if (!isolatedJRegenSameFingerprint_(confirmed.fingerprint, target.before.fingerprint)) throw new Error('J再組版の更新前deck fingerprintがpreflight後に変化しました');
    const prepared = { schema: 'isolation-j-no7-regeneration/v2', status: 'prepared', automaticRetry: false, folderId: c.folderId,
      sourceSha256: target.chain.sourceSha256, mapFileId: c.mapFileId, previousDeckId: target.beforeDeckId,
      previousMapSha256: target.beforeMapSha256, previousDeckFingerprint: target.before.fingerprint,
      previousTimestampDiagnostic: target.before.timestampDiagnostic, createdAt: new Date().toISOString() };
    const record = target.folder.createFile(c.regenerationRecordFileName, JSON.stringify(prepared), MimeType.PLAIN_TEXT);
    try {
      const built = スライドを組む_(target.cfg, target.ledger.row.row, false);
      if (built !== true) throw new Error('固定隔離の正規スライド再組版が完了しませんでした: ' + built);
      const mapFile = isolatedJRegenOneFile_(target.folder, c.mapFileName, 'J再組版後対応表');
      const map = JSON.parse(mapFile.getBlob().getDataAsString('UTF-8'));
      if (!map || typeof map.presentationId !== 'string' || !map.presentationId || map.presentationId === target.beforeDeckId) throw new Error('J再組版後の新deck IDを対応表から確定できません');
      const after = isolatedJRegenMapDeck_(target.folder, target.generator.expectedByNo, null, map.presentationId, 'after'), afterMapSha256 = isolatedJRegenHash_(after.mapText);
      record.setContent(JSON.stringify({ schema: 'isolation-j-no7-regeneration/v2', status: 'completed', automaticRetry: false, folderId: c.folderId,
        sourceSha256: target.chain.sourceSha256, mapFileId: c.mapFileId, previousDeckId: target.beforeDeckId,
        previousMapSha256: target.beforeMapSha256, previousDeckFingerprint: target.before.fingerprint,
        previousTimestampDiagnostic: target.before.timestampDiagnostic, newDeckId: map.presentationId,
        newMapSha256: afterMapSha256, newDeckFingerprint: after.fingerprint,
        newTimestampDiagnostic: after.timestampDiagnostic, completedAt: new Date().toISOString() }));
      return { ok: true, mode: 'fixed-no7-normal-regeneration', newDeckId: map.presentationId, mapFileId: c.mapFileId, newMapSha256: afterMapSha256, actions: { regenerated: true, imported: false, audioStarted: false, delivered: false, notified: false } };
    } catch (e) {
      prepared.status = 'uncertain'; prepared.failure = String(e && e.message ? e.message : e).slice(0, 240); prepared.recordedAt = new Date().toISOString();
      record.setContent(JSON.stringify(prepared));
      throw new Error('J再組版の結果が不確定です。自動再実行せず実行記録を手動精査してください');
    }
  });
}
function 隔離通し_J台本更新後の再組版を読取検品する() {
  const target = isolatedJRegenerationContext_(), c = ISOLATED_J_REGEN_AUDIT_, recordFile = isolatedJRegenerationRecord_(target.folder);
  if (!recordFile) throw new Error('J再組版実行記録がありません');
  let record;
  try { record = JSON.parse(recordFile.getBlob().getDataAsString('UTF-8')); } catch (e) { throw new Error('J再組版実行記録JSONを読めません'); }
  if (!record || record.schema !== 'isolation-j-no7-regeneration/v2' || record.status !== 'completed' || record.automaticRetry !== false || record.folderId !== c.folderId || record.sourceSha256 !== c.sourceSha256After || record.mapFileId !== c.mapFileId || typeof record.previousDeckId !== 'string' || typeof record.previousMapSha256 !== 'string' || !isolatedJRegenValidFingerprint_(record.previousDeckFingerprint) || typeof record.newDeckId !== 'string' || !record.newDeckId || record.newDeckId === record.previousDeckId || typeof record.newMapSha256 !== 'string' || !isolatedJRegenValidFingerprint_(record.newDeckFingerprint) || !record.previousTimestampDiagnostic || !record.newTimestampDiagnostic) throw new Error('J再組版completed実行記録が固定契約と一致しません');
  const after = isolatedJRegenMapDeck_(target.folder, target.generator.expectedByNo, record.newMapSha256, record.newDeckId, 'after');
  if (!isolatedJRegenSameFingerprint_(after.fingerprint, record.newDeckFingerprint)) throw new Error('J再組版後deck fingerprintがcompleted実行記録と一致しません');
  return { ok: true, mode: 'read-only-after-no7-regeneration', newDeckId: record.newDeckId, mapFileId: c.mapFileId,
    mapSha256: record.newMapSha256, deckFingerprint: after.fingerprint, timestampDiagnostic: after.timestampDiagnostic,
    slideCount: after.slideCount, geometryCount: after.geometryCount, notesSha256: after.notesSha256,
    actions: { regenerated: false, imported: false, audioStarted: false, delivered: false, notified: false } };
}
