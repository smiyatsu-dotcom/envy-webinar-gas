/**
 * Fixed-fixture audit/reconciliation entry points.
 * They deliberately do not call processNewFlow_: no registry scan, generation,
 * approval, delivery, notification, or audio action is permitted here.
 */
const ISOLATED_CASE_AUDIT_ = {
  company: '隔離通し診断テスト社',
  project: '隔離通し診断テスト社_ヒアリングシート',
  hearingId: '1ur8hPUElZMrKuitnDo1j1wzr54EkmRY9efNMaEQFErs',
  outputFolderId: '1Ko7cbJ2nW2C4uQbZfDnZjW6mfz_qmHZ0',
  slidesId: '12jfYwhluLE4r6DQrueSl45x8BnZprbaMtgJBVCugU84'
};
function 隔離通し_固定対象を識別() {
  const root = DriveApp.getFolderById(ISOLATED_CASE_AUDIT_.outputFolderId), children = [];
  const it = root.getFolders(); while (it.hasNext()) children.push(it.next().getName());
  return { rootName: root.getName(), children: children };
}

function isolatedCaseAuditFolder_() {
  const c = ISOLATED_CASE_AUDIT_;
  const root = DriveApp.getFolderById(c.outputFolderId);
  if (root.getName() === c.company || root.getName() === c.project) return root;
  const folders = root.getFoldersByName(c.project);
  if (!folders.hasNext()) throw new Error('固定隔離案件フォルダがありません');
  const folder = folders.next();
  if (folders.hasNext() || folder.getName() !== c.project) throw new Error('固定隔離案件フォルダが一意ではありません');
  return folder;
}

function isolatedCaseElementText_(element, out) {
  const type = String(element.getPageElementType());
  if (type === 'SHAPE') {
    out.shapes++;
    try { out.text.push(element.asShape().getText().asString()); } catch (e) {}
  } else if (type === 'TABLE') {
    out.tables++;
    try {
      const t = element.asTable();
      for (let r = 0; r < t.getNumRows(); r++) for (let c = 0; c < t.getNumColumns(); c++) out.text.push(t.getCell(r, c).getText().asString());
    } catch (e) {}
  } else if (type === 'GROUP') {
    out.groups++;
    try { element.asGroup().getChildren().forEach(function(child) { isolatedCaseElementText_(child, out); }); } catch (e) {}
  }
}

/** Read-only artifact inspection; does not use the Slides API. */
function 隔離通し_資料を検品() {
  const c = ISOLATED_CASE_AUDIT_;
  const pres = SlidesApp.openById(c.slidesId);
  const slides = pres.getSlides();
  const all = { shapes: 0, tables: 0, groups: 0, text: [], notes: [] };
  slides.forEach(function(slide) {
    slide.getPageElements().forEach(function(element) { isolatedCaseElementText_(element, all); });
    try { all.notes.push(slide.getNotesPage().getSpeakerNotesShape().getText().asString()); } catch (e) { all.notes.push(''); }
  });
  const corpus = all.text.concat(all.notes).join('\n');
  const last = { shapes: 0, tables: 0, groups: 0, text: [] };
  slides[slides.length - 1].getPageElements().forEach(function(element) { isolatedCaseElementText_(element, last); });

  const folder = isolatedCaseAuditFolder_();
  const root = DriveApp.getFolderById(c.outputFolderId);
  const qa = [];
  [root, folder].forEach(function(scanFolder) {
    const files = scanFolder.getFiles();
    while (files.hasNext()) {
      const f = files.next(), name = f.getName();
      if (!/Q\s*&\s*A|Q＆A|Q&A|カンペ/i.test(name)) continue;
      const item = { parent: scanFolder.getName(), name: name, id: f.getId(), mime: f.getMimeType(), size: f.getSize(), textReadable: false, chars: 0 };
      if (f.getMimeType() === MimeType.GOOGLE_DOCS) {
        try { item.chars = DocumentApp.openById(f.getId()).getBody().getText().length; item.textReadable = true; } catch (e) { item.readError = String(e); }
      }
      qa.push(item);
    }
  });
  const ss = SpreadsheetApp.openById(c.hearingId);
  const faq = [];
  ss.getSheets().forEach(function(sheet) {
    const values = sheet.getDataRange().getDisplayValues();
    values.forEach(function(row, index) {
      const line = row.join(' ').trim();
      if (/FAQ|よくある質問|Q\s*&\s*A|Q＆A|Q&A/i.test(line)) faq.push({ sheet: sheet.getName(), row: index + 1, chars: line.length });
    });
  });
  return {
    caseName: c.company, slidesId: c.slidesId, slideCount: slides.length,
    elementCounts: { shapes: all.shapes, tables: all.tables, groups: all.groups }, notesChars: all.notes.join('\n').length,
    zoomOccurrences: (corpus.match(/Zoom/gi) || []).length,
    qaIsLastSlide: /Q\s*&\s*A|Q＆A|Q&A|よくある質問/i.test(last.text.join('\n')),
    qaCueFiles: qa, hearingFaqMarkers: faq,
    passed: { noZoom: !/Zoom/i.test(corpus), qaNotLast: !/Q\s*&\s*A|Q＆A|Q&A|よくある質問/i.test(last.text.join('\n')), qaDocAndPdf: qa.some(function(x) { return x.mime === MimeType.GOOGLE_DOCS && x.textReadable; }) && qa.some(function(x) { return x.mime === MimeType.PDF; }) }
  };
}

// This is intentionally separate from ISOLATED_CASE_AUDIT_: that older audit
// is pinned to a historical deck.  The J audit can inspect only the one
// completed, non-customer isolation output and never accepts a Drive ID.
const ISOLATED_J_CURRENT_AUDIT_ = {
  folderId: '1Ko7cbJ2nW2C4uQbZfDnZjW6mfz_qmHZ0',
  deckId: '1wLXXeU95iculuFJ_NOQ5oSZ3PSSWXLRa_L6m2GuX_vM',
  mapFileName: '【スライド対応】隔離通し診断テスト社.json',
  geometryFileName: '【配置指定】隔離通し診断テスト社.json',
  scriptFileName: '【台本JSON】隔離通し診断テスト社.txt',
  textFileName: '【台本】隔離通し診断テスト社_ヒアリングシート.txt',
  importRecordFileName: '【不確定】J台本取込記録.json'
};

function isolatedJCurrentFile_(folder, name, label) {
  const file = 隔離通し_同名ファイル1件_(folder, name, label);
  if (!file) throw new Error(label + ' が固定隔離案件にありません');
  return file;
}
// The current-J audit accepts only one representational normalization: CRLF
// (and a legacy lone CR) becomes LF.  It never trims, drops an empty note,
// or removes wording to make two artifacts look equal.
function isolatedJCurrentNormalizeLineEndings_(value) {
  return String(value == null ? '' : value).replace(/\r\n?/g, '\n');
}
function isolatedJCurrentNote_(slide) {
  try {
    const rawText = String(slide.getNotesPage().getSpeakerNotesShape().getText().asString());
    return { readable: true, rawText: rawText, text: isolatedJCurrentNormalizeLineEndings_(rawText) };
  } catch (e) {
    return { readable: false, rawText: '', text: '' };
  }
}
// Reproduce the normal generator's note input in its real order.  Import JSON
// is first resolved through the same fixed-row resolver used by the structure
// sheet, then the generator's only note-text normalization is applied.  The
// audit deliberately does not trim or otherwise weaken that expected text.
function isolatedJCurrentNormalGeneratorExpected_(scriptRows) {
  const sourceByNo = {};
  scriptRows.forEach(function(row) {
    const no = Number(row && row.no);
    if (!Number.isInteger(no) || no < 1 || no > 106 || Object.prototype.hasOwnProperty.call(sourceByNo, no)) {
      throw new Error('J現行検品の台本JSON行番号が一意な1〜106ではありません');
    }
    sourceByNo[no] = String(row.daihon == null ? '' : row.daihon);
  });
  if (Object.keys(sourceByNo).length !== 106) throw new Error('J現行検品の台本JSONが106行ありません');
  const defs = rowDefs_();
  const expectedByNo = {};
  if (!Array.isArray(defs) || defs.length !== 106) throw new Error('J現行検品の生成定義が106行ありません');
  defs.forEach(function(def) {
    const no = Number(def && def.no);
    if (!Number.isInteger(no) || no < 1 || no > 106 || Object.prototype.hasOwnProperty.call(expectedByNo, no)) {
      throw new Error('J現行検品の生成定義行番号が一意な1〜106ではありません');
    }
    // resolveDaihon_ is the same routine invoked while normal generation
    // prepares the structure-sheet row.  In particular, dm=FIX wins over
    // imported JSON before normalizeWebinarTerms_ runs in generateSlides_.
    expectedByNo[no] = isolatedJCurrentNormalizeLineEndings_(normalizeWebinarTerms_(resolveDaihon_(def, sourceByNo)));
  });
  if (Object.keys(expectedByNo).length !== 106) throw new Error('J現行検品の期待ノートが106行になりません');
  return { sourceByNo: sourceByNo, expectedByNo: expectedByNo };
}
function isolatedJCurrentShortDifferenceExcerpt_(value, firstDifferentCharacter) {
  const chars = Array.from(String(value == null ? '' : value));
  const start = Math.max(0, Math.min(Number(firstDifferentCharacter) || 0, chars.length) - 12);
  const end = Math.min(chars.length, start + 32);
  return chars.slice(start, end).join('');
}
function isolatedJCurrentNoteDifference_(rawText, normalizedText, expectedText) {
  const actual = Array.from(String(normalizedText == null ? '' : normalizedText));
  const expected = Array.from(String(expectedText == null ? '' : expectedText));
  let first = 0;
  while (first < actual.length && first < expected.length && actual[first] === expected[first]) first++;
  if (first === actual.length && first === expected.length) {
    return { matches: true, category: String(rawText) === String(expectedText) ? 'exact-match' : 'line-ending-only-match', actualLength: actual.length, expectedLength: expected.length };
  }
  let category = 'content-difference';
  if (actual.length === 0) category = 'empty-note';
  else if (first === actual.length) category = 'note-is-prefix-of-expected';
  else if (first === expected.length) category = 'note-has-extra-suffix';
  else if (actual.length > expected.length && actual.slice(actual.length - expected.length).join('') === expected.join('')) category = 'note-has-extra-prefix';
  return { matches: false, category: category, actualLength: actual.length, expectedLength: expected.length,
    firstDifferentCharacter: first + 1,
    actualExcerpt: isolatedJCurrentShortDifferenceExcerpt_(actual.join(''), first),
    expectedExcerpt: isolatedJCurrentShortDifferenceExcerpt_(expected.join(''), first) };
}
function isolatedJCurrentVisibleText_(element, out) {
  let type = '';
  try { type = String(element.getPageElementType()); }
  catch (e) { out.readErrors++; return; }
  if (type === 'SHAPE') {
    try { out.text.push(isolatedJCurrentNormalizeLineEndings_(element.asShape().getText().asString())); }
    catch (e) { out.readErrors++; }
    return;
  }
  if (type === 'TABLE') {
    try {
      const table = element.asTable();
      for (let row = 0; row < table.getNumRows(); row++) {
        for (let column = 0; column < table.getNumColumns(); column++) {
          out.text.push(isolatedJCurrentNormalizeLineEndings_(table.getCell(row, column).getText().asString()));
        }
      }
    } catch (e) { out.readErrors++; }
    return;
  }
  if (type === 'WORD_ART') {
    try { out.text.push(isolatedJCurrentNormalizeLineEndings_(element.asWordArt().getRenderedText())); }
    catch (e) { out.readErrors++; }
    return;
  }
  if (type === 'GROUP') {
    try { element.asGroup().getChildren().forEach(function(child) { isolatedJCurrentVisibleText_(child, out); }); }
    catch (e) { out.readErrors++; }
  }
}
function isolatedJCurrentZoomCount_(text) {
  return (String(text || '').match(/zoom/gi) || []).length;
}
function isolatedJCurrentFixedNos_() {
  const expected = [2, 7, 33, 38, 49, 97, 101, 105, 106];
  if (typeof FIXED_DAIHON !== 'object' || !FIXED_DAIHON) throw new Error('J現行検品の固定台本定義がありません');
  const actual = Object.keys(FIXED_DAIHON).map(function(key) { return Number(key); }).sort(function(a, b) { return a - b; });
  if (actual.length !== expected.length || actual.some(function(no, index) { return no !== expected[index] || typeof FIXED_DAIHON[no] !== 'string'; })) {
    throw new Error('J現行検品の固定9行定義が基準と一致しません');
  }
  return expected;
}
function isolatedJCurrentAuditTarget_() {
  const c = ISOLATED_J_CURRENT_AUDIT_;
  const cfg = 隔離通し_生成設定を作る_();
  const output = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  if (cfg.ISOLATION_TEST !== true || output.getId() !== 隔離通し_生成_生成物rootID || output.getName() !== '【出力】生成物') {
    throw new Error('J現行検品の生成物rootが固定隔離fixtureと一致しません');
  }
  const folder = caseFolderIfExists_(cfg, 隔離通し_生成_案件名);
  if (!folder || folder.getId() !== c.folderId || folder.getName() !== 隔離通し_生成_案件名) {
    throw new Error('J現行検品の案件フォルダが固定隔離fixtureと一致しません');
  }
  隔離通し_生成の親を確かめる_(folder, output.getId(), 'J現行検品の案件フォルダ');
  const mapFile = isolatedJCurrentFile_(folder, c.mapFileName, 'J現行スライド対応表');
  const geometryFile = isolatedJCurrentFile_(folder, c.geometryFileName, 'J現行配置指定');
  const scriptFile = isolatedJCurrentFile_(folder, c.scriptFileName, 'J現行台本JSON');
  // Retain the identity check for the completed audio-text artifact, but do
  // not use its lossy, narration-oriented projection as acceptance evidence.
  isolatedJCurrentFile_(folder, c.textFileName, 'J現行音声台本TXT');
  const recordFile = isolatedJCurrentFile_(folder, c.importRecordFileName, 'J台本取込記録');
  const scriptText = scriptFile.getBlob().getDataAsString('UTF-8');
  const scriptRows = 隔離通し_J台本_本文を検証する_(scriptText);
  let importRecord;
  try { importRecord = JSON.parse(recordFile.getBlob().getDataAsString('UTF-8')); }
  catch (e) { throw new Error('J台本取込記録JSONを読めません'); }
  if (!importRecord || importRecord.status !== 'completed' || importRecord.folderId !== c.folderId ||
      importRecord.targetScriptSha256 !== 隔離通し_J台本_SHA256 || importRecord.automaticRetry !== false) {
    throw new Error('J台本取込completed記録が固定条件と一致しません');
  }
  let mapping, geometry;
  try { mapping = JSON.parse(mapFile.getBlob().getDataAsString('UTF-8')); }
  catch (e) { throw new Error('J現行スライド対応表JSONを読めません'); }
  try { geometry = JSON.parse(geometryFile.getBlob().getDataAsString('UTF-8')); }
  catch (e) { throw new Error('J現行配置指定JSONを読めません'); }
  if (!mapping || mapping.schema !== GEOMETRY_MAP_SCHEMA || mapping.caseName !== 隔離通し_生成_案件名 ||
      mapping.presentationId !== c.deckId || !Array.isArray(mapping.pages) || mapping.pages.length !== 106 ||
      mapping.pages.some(function(page, index) { return !page || page.slideNo !== index + 1 || typeof page.pageId !== 'string'; }) ||
      new Set(mapping.pages.map(function(page) { return page.pageId; })).size !== 106) {
    throw new Error('J現行スライド対応表が固定106枚deckを示しません');
  }
  if (!geometry || geometry.schema !== GEOMETRY_SCHEMA || geometry.caseName !== 隔離通し_生成_案件名 || !Array.isArray(geometry.records)) {
    throw new Error('J現行配置指定が固定案件と一致しません');
  }
  const deckFile = DriveApp.getFileById(c.deckId);
  const parents = deckFile.getParents();
  if (deckFile.isTrashed() || deckFile.getMimeType() !== MimeType.GOOGLE_SLIDES || !parents.hasNext() || parents.next().getId() !== folder.getId() || parents.hasNext()) {
    throw new Error('J現行deckが固定隔離案件フォルダに一意にありません');
  }
  return { cfg: cfg, folder: folder, mapping: mapping, geometry: geometry, deckFile: deckFile,
    scriptHash: 隔離通し_J台本_sha256_(scriptText), scriptRows: scriptRows };
}

/**
 * Read-only reconciliation for the one completed J output.  It neither saves
 * the presentation nor calls import, regeneration, map repair, delivery,
 * notification, or audio code.
 */
function 隔離通し_J現行資料を読取検品() {
  const target = isolatedJCurrentAuditTarget_();
  return withIsolationTestConfig_(target.cfg, function() {
    const slides = SlidesApp.openById(ISOLATED_J_CURRENT_AUDIT_.deckId).getSlides();
    const byId = {};
    let duplicateSlideIds = 0;
    slides.forEach(function(slide) {
      const id = String(slide.getObjectId() || '');
      if (!id || byId[id]) duplicateSlideIds++;
      byId[id] = slide;
    });
    if (slides.length !== 106 || duplicateSlideIds || target.mapping.pages.some(function(page) { return !byId[page.pageId]; })) {
      throw new Error('J現行deckが対応表の固定106ページと一致しません');
    }
    const normalExpected = isolatedJCurrentNormalGeneratorExpected_(target.scriptRows);
    const rowsByNo = normalExpected.sourceByNo;
    const expectedRowsByNo = normalExpected.expectedByNo;
    const noteRows = [];
    const noteMismatchNos = [], emptyNoteNos = [], unreadableNoteNos = [];
    const noteDifferences = [], differenceCounts = { 'exact-match': 0, 'line-ending-only-match': 0,
      'empty-note': 0, 'note-is-prefix-of-expected': 0, 'note-has-extra-prefix': 0,
      'note-has-extra-suffix': 0, 'content-difference': 0, 'unreadable-note': 0 };
    target.mapping.pages.forEach(function(page) {
      const note = isolatedJCurrentNote_(byId[page.pageId]);
      const expected = expectedRowsByNo[page.slideNo];
      noteRows.push({ no: page.slideNo, text: note.text });
      if (!note.readable) {
        unreadableNoteNos.push(page.slideNo);
        differenceCounts['unreadable-note']++;
        noteDifferences.push({ no: page.slideNo, category: 'unreadable-note', actualLength: null,
          expectedLength: Array.from(String(expected)).length });
        noteMismatchNos.push(page.slideNo);
        return;
      }
      if (note.text === '') emptyNoteNos.push(page.slideNo);
      const difference = isolatedJCurrentNoteDifference_(note.rawText, note.text, expected);
      differenceCounts[difference.category]++;
      if (!difference.matches) {
        noteMismatchNos.push(page.slideNo);
        noteDifferences.push(Object.assign({ no: page.slideNo }, difference));
      }
    });
    const fixedNos = isolatedJCurrentFixedNos_();
    const fixedBaselineHash = 隔離通し_J台本_sha256_(JSON.stringify(fixedNos.map(function(no) {
      return { no: no, daihon: isolatedJCurrentNormalizeLineEndings_(FIXED_DAIHON[no]) };
    })));
    const fixedScriptMismatchNos = fixedNos.filter(function(no) { return isolatedJCurrentNormalizeLineEndings_(rowsByNo[no]) !== isolatedJCurrentNormalizeLineEndings_(FIXED_DAIHON[no]); });
    const fixedNoteMismatchNos = fixedNos.filter(function(no) {
      return noteRows[no - 1].text !== isolatedJCurrentNormalizeLineEndings_(FIXED_DAIHON[no]);
    });
    const visible = { text: [], readErrors: 0 };
    slides.forEach(function(slide) {
      slide.getPageElements().forEach(function(element) { isolatedJCurrentVisibleText_(element, visible); });
    });
    const notesCorpus = noteRows.map(function(row) { return row.text; }).join('\n');
    const visibleCorpus = visible.text.join('\n');
    const noteHash = 隔離通し_J台本_sha256_(JSON.stringify(noteRows));
    const visibleTextHash = 隔離通し_J台本_sha256_(visibleCorpus);
    const geometry = target.geometry.records.filter(function(record) { return record && record.status === 'active'; }).map(function(record) {
      const page = target.mapping.pages[Number(record.slideNo) - 1];
      const selector = page && page.targetSelector;
      const elements = page && byId[page.pageId] ? byId[page.pageId].getPageElements() : [];
      const matches = selector ? elements.filter(function(element) {
        try { return selectorMatches_(element, selector); } catch (e) { return false; }
      }) : [];
      const selectorAfter = !!selector && validGeometry_(record.after) && sameGeometry_(selector.templateGeometry, record.after);
      return { slideNo: record.slideNo, mapSelector: !!selector, selectorAfter: selectorAfter,
        currentDeckMatches: matches.length, retained: selectorAfter && matches.length === 1 };
    });
    const geometryRetained = geometry.length > 0 && geometry.every(function(item) { return item.retained; });
    const scriptNotesMatch = noteMismatchNos.length === 0 && emptyNoteNos.length === 0 && unreadableNoteNos.length === 0;
    const fixedMatch = fixedScriptMismatchNos.length === 0 && fixedNoteMismatchNos.length === 0;
    const noZoom = isolatedJCurrentZoomCount_(notesCorpus) === 0 && isolatedJCurrentZoomCount_(visibleCorpus) === 0 && visible.readErrors === 0;
    // No source wording, folder ID, deck ID, or slide text is emitted.  This
    // result is safe to retain as an external execution record.
    return { ok: scriptNotesMatch && fixedMatch && noZoom && geometryRetained,
      mode: 'read-only-current-j-strict-audit', fixedIsolationTarget: true, slideCount: slides.length,
      scriptSha256: target.scriptHash, normalizedNotesSha256: noteHash, visibleSlideTextSha256: visibleTextHash,
      lineEndingNormalization: 'CRLF-or-CR-to-LF-only',
      normalGeneratorExpectation: { order: ['resolveDaihon_-fixed-before-imported-JSON', 'normalizeWebinarTerms_', 'CRLF-or-CR-to-LF-comparison-only'],
        expectedRows: 106, expectedNotesSha256: 隔離通し_J台本_sha256_(JSON.stringify(Object.keys(expectedRowsByNo).map(function(no) { return { no: Number(no), daihon: expectedRowsByNo[no] }; }))),
        fixedSourceDivergenceNos: fixedScriptMismatchNos },
      scriptNotes: { expectedRows: 106, comparedRows: noteRows.length, mismatchCount: noteMismatchNos.length,
        mismatchNos: noteMismatchNos, emptyNoteCount: emptyNoteNos.length, emptyNoteNos: emptyNoteNos,
        unreadableNoteCount: unreadableNoteNos.length, unreadableNoteNos: unreadableNoteNos },
      noteDifferenceSummary: differenceCounts,
      // Excerpts appear only for genuine non-equivalences and are capped at 32
      // characters each.  The target is a fixed non-customer fixture.
      noteDifferences: noteDifferences,
      fixedScript: { expectedCount: fixedNos.length, fixedNos: fixedNos, baselineSha256: fixedBaselineHash, scriptMismatchNos: fixedScriptMismatchNos,
        noteMismatchNos: fixedNoteMismatchNos },
      zoomOccurrences: { notes: isolatedJCurrentZoomCount_(notesCorpus), visibleSlides: isolatedJCurrentZoomCount_(visibleCorpus), visibleTextReadErrors: visible.readErrors },
      geometry: geometry, actions: { imported: false, regenerated: false, mapRepaired: false, delivered: false, notified: false, audioStarted: false } };
  });
}
function runIsolationCurrentJReadAudit() { return 隔離通し_J現行資料を読取検品(); }

/**
 * Apply only the calculation/write tail of 新の流れを1件進める_ to the fixed case.
 * It never accepts approval/delivery instructions and refuses any non-false B2.
 */
function 隔離通し_正規進みを再集計() {
  const c = ISOLATED_CASE_AUDIT_, cfg = getConfig_();
  const reg = registry_(cfg), rowNo = reg[c.hearingId];
  if (!rowNo) throw new Error('固定ヒアリングIDが管理台帳にありません');
  const row = mgmtSheet_(cfg).getRange(rowNo, 1, 1, 流れ_列).getValues()[0];
  if (!新の流れか_(row) || String(row[1]) !== c.project) throw new Error('固定隔離案件以外を再集計しようとしました');
  const folder = isolatedCaseAuditFolder_();
  const ss = SpreadsheetApp.openById(idFromUrl_(String(row[3] || ''))), settings = ss.getSheetByName(TABS.SETTINGS);
  if (!settings || settings.getRange('B2').getValue() !== false) throw new Error('承認falseの固定条件を満たしません');
  let observed = 観測する_(cfg, row, folder);
  observed.見張り = null;
  if (observed.承認ファイル || observed.納品ファイル) throw new Error('承認または納品指示があるため隔離再集計を中止しました');
  const record = 進み_記録を読む_(c.hearingId);
  let progress = 進みを計算_(observed, record, new Date());
  if (progress.全部できた && (!record.生成 || record.生成.revision !== progress.生成.revision)) {
    record.生成 = { revision: progress.生成.revision, 日時: 日時の文字_(new Date()) };
    progress = 進みを計算_(observed, record, new Date());
  }
  // These are the same final persistence primitives used by processNewFlow_.
  進み_記録を書く_(c.hearingId, record);
  進みを書く_(folder, progress);
  return { caseName: c.company, hearingId: c.hearingId, progressFile: 進み_ファイル名,
    revision: progress.生成.revision, revisionIsSha256: /^[0-9a-f]{64}$/.test(progress.生成.revision),
    approval: progress.承認, delivery: progress.納品, allComplete: progress.全部できた,
    actions: { generated: false, approvalChanged: false, delivered: false, notified: false, audioStarted: false } };
}

/**
 * Fixed-wrapper counterpart for the real isolation ledger.  It calls the
 * existing Q&A writer and the existing progress calculator only after the
 * wrapper has proved the one-row isolation configuration.  No main pass,
 * generation, approval, delivery, or notification is reachable here.
 */
function 隔離通し_カンペと進みを正規更新() {
  const cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function() {
    const rows = mgmtSheet_(cfg).getDataRange().getValues();
    const hits = rows.slice(1).filter(function(row) {
      return String(row[0] || '') === 隔離通し_生成_ヒアリングID &&
        String(row[1] || '') === 隔離通し_生成_案件名 + '_ヒアリングシート' &&
        新の流れか_(row);
    });
    if (hits.length !== 1 || rows.slice(1).filter(function(row) { return 新の流れか_(row); }).length !== 1) throw new Error('隔離台帳の固定新フロー1件を確認できません');
    const row = hits[0], project = String(row[1]);
    const folder = caseFolderIfExists_(cfg, project);
    if (!folder) throw new Error('隔離案件フォルダがありません');
    const settings = SpreadsheetApp.openById(idFromUrl_(String(row[3] || ''))).getSheetByName(TABS.SETTINGS);
    if (!settings || settings.getRange('B2').getValue() !== false) throw new Error('承認falseの固定条件を満たしません');
    let observed = 観測する_(cfg, row, folder); observed.見張り = null;
    if (observed.承認ファイル || observed.納品ファイル) throw new Error('承認または納品指示があるため中止しました');
    const cue = syncQaCue_(cfg, project);
    const expectedCue = qaCueText_(qaCueRequireCompleteFaq_(SpreadsheetApp.openById(隔離通し_生成_ヒアリングID).getSheetByName(HEARING_TABS.OFFER).getDataRange().getValues()));
    const actualCue = DocumentApp.openById(cue.getId()).getBody().getText();
    const pdfs = filesByName_(folder, qaCuePdfFileName_(隔離通し_生成_案件名));
    if (actualCue !== expectedCue || !pdfs.hasNext() || pdfs.next().getSize() < 1 || pdfs.hasNext()) throw new Error('Q&AカンペDoc/PDFの正規照合に失敗しました');
    const record = 進み_記録を読む_(隔離通し_生成_ヒアリングID);
    let progress = 進みを計算_(observed, record, new Date());
    if (progress.全部できた && (!record.生成 || record.生成.revision !== progress.生成.revision)) {
      record.生成 = { revision: progress.生成.revision, 日時: 日時の文字_(new Date()) };
      progress = 進みを計算_(observed, record, new Date());
    }
    進み_記録を書く_(隔離通し_生成_ヒアリングID, record);
    進みを書く_(folder, progress);
    return { ok: true, caseName: 隔離通し_生成_案件名, managementSheetId: cfg.MGMT_SHEET_ID,
      outputRootId: cfg.OUTPUT_FOLDER_ID, cueDocId: cue.getId(), cuePdf: true,
      revision: progress.生成.revision, revisionIsSha256: /^[0-9a-f]{64}$/.test(progress.生成.revision),
      approval: progress.承認, delivery: progress.納品,
      actions: { generated: false, approvalChanged: false, delivered: false, notified: false, audioStarted: false } };
  });
}

function runIsolationCueAndProgress() { return 隔離通し_カンペと進みを正規更新(); }

function runIsolationProgressOnly() {
  const cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function() {
    const rows = mgmtSheet_(cfg).getDataRange().getValues();
    const row = rows.slice(1).filter(function(r) { return String(r[0] || '') === 隔離通し_生成_ヒアリングID && String(r[1] || '') === 隔離通し_生成_案件名 + '_ヒアリングシート' && 新の流れか_(r); })[0];
    if (!row || rows.slice(1).filter(function(r) { return 新の流れか_(r); }).length !== 1) throw new Error('隔離台帳の固定新フロー1件を確認できません');
    const folder = caseFolderIfExists_(cfg, String(row[1]));
    const settings = SpreadsheetApp.openById(idFromUrl_(String(row[3] || ''))).getSheetByName(TABS.SETTINGS);
    if (!folder || !settings || settings.getRange('B2').getValue() !== false) throw new Error('隔離フォルダまたは承認falseの固定条件を確認できません');
    const observed = 観測する_(cfg, row, folder); observed.見張り = null;
    if (observed.承認ファイル || observed.納品ファイル) throw new Error('承認または納品指示があるため中止しました');
    const record = 進み_記録を読む_(隔離通し_生成_ヒアリングID);
    let progress = 進みを計算_(observed, record, new Date());
    if (progress.全部できた && (!record.生成 || record.生成.revision !== progress.生成.revision)) { record.生成 = { revision: progress.生成.revision, 日時: 日時の文字_(new Date()) }; progress = 進みを計算_(observed, record, new Date()); }
    進み_記録を書く_(隔離通し_生成_ヒアリングID, record); 進みを書く_(folder, progress);
    return { ok: true, caseName: 隔離通し_生成_案件名, managementSheetId: cfg.MGMT_SHEET_ID, outputRootId: cfg.OUTPUT_FOLDER_ID, revision: progress.生成.revision, revisionIsSha256: /^[0-9a-f]{64}$/.test(progress.生成.revision), approval: progress.承認, delivery: progress.納品, allComplete: progress.全部できた, actions: { generated:false, approvalChanged:false, delivered:false, notified:false, audioStarted:false } };
  });
}

// Manual fixed-isolation write test.  This is not an audit and is never called
// by main()/processNewFlow_: it may write the named fixture's FAQ cells and
// isolated cue Doc/PDF only after its fixed-ID and backup checks.  It does not
// deliver, notify, change approval, or start audio.
function runIsolationFixtureFaqNormalCase() {
  const cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function() {
    const root = DriveApp.getFolderById(隔離通し_生成_rootID);
    if (root.getName() !== 隔離通し_生成_root名) throw new Error('固定隔離rootが一致しません');
    const sheet = SpreadsheetApp.openById(隔離通し_生成_ヒアリングID);
    const offer = sheet.getSheetByName(HEARING_TABS.OFFER);
    if (!offer) throw new Error('固定ヒアリングのオファータブがありません');
    const values = offer.getDataRange().getValues(), norm = function(v) { return String(v == null ? '' : v).replace(/[:：\s]/g, ''); };
    let start = -1;
    for (let r = 0; r < values.length; r++) if (values[r].some(function(v) { return /よくある質問[（(]\s*Q\s*[&＆]\s*A\s*[）)]/i.test(String(v)); })) { start = r + 1; break; }
    if (start < 0) throw new Error('固定ヒアリングにFAQ欄がありません');
    const pairs = [], blank = function(a) { return a.every(function(v) { return String(v == null ? '' : v).trim() === ''; }); };
    let question = null;
    for (let r = start; r < values.length && pairs.length < 3; r++) {
      const non = values[r].map(String).filter(function(v) { return v.trim(); });
      if (non.length && /^([⒈-⒛]|[0-9０-９]+[\.．])\s*/.test(non[0])) break;
      const qi = values[r].findIndex(function(v) { return norm(v) === 'よくある質問'; });
      const ai = values[r].findIndex(function(v) { return norm(v) === '記入欄'; });
      if (qi >= 0) question = { row:r, col:qi + 1, values:values[r].slice(qi + 1) };
      if (ai >= 0 && question) { pairs.push({ q:question, a:{ row:r, col:ai + 1, values:values[r].slice(ai + 1) } }); question = null; }
    }
    if (pairs.length !== 3 || blank(pairs[0].q.values) || blank(pairs[0].a.values) || !blank(pairs[1].q.values) || !blank(pairs[1].a.values) || !blank(pairs[2].q.values) || !blank(pairs[2].a.values)) throw new Error('FAQ1のみ記入・FAQ2/3空欄という固定試験前提を確認できません');
    const backupName = '【控え】隔離通しFAQ正常系_追加前.json', old = filesByName_(root, backupName);
    if (!old.hasNext()) root.createFile(backupName, JSON.stringify({ schema:1, hearingId:隔離通し_生成_ヒアリングID, sheet:HEARING_TABS.OFFER, cells:[pairs[1].q,pairs[1].a,pairs[2].q,pairs[2].a].map(function(p) { return { row:p.row+1, col:p.col+1, values:p.values }; }) }), MimeType.PLAIN_TEXT);
    else if (old.hasNext()) throw new Error('FAQ追加前バックアップが複数あります');
    const fill = [
      ['この隔離通しテストの架空設定で、導入前に準備するものはありますか？'],
      ['架空テスト用の入力として、現在お使いの商談記録や提案資料があれば確認材料として使えます。'],
      ['この隔離通しテストの架空設定で、少人数のチームでも始められますか？'],
      ['架空テスト用の入力として、週次の振り返りと一つの改善テーマから始められます。']
    ];
    [pairs[1].q,pairs[1].a,pairs[2].q,pairs[2].a].forEach(function(p, i) { offer.getRange(p.row + 1, p.col + 1).setValue(fill[i][0]); });
    SpreadsheetApp.flush();
    const project = 隔離通し_生成_案件名 + '_ヒアリングシート', folder = caseFolderIfExists_(cfg, project);
    if (!folder) throw new Error('隔離案件フォルダがありません');
    const cue = syncQaCue_(cfg, project), expected = qaCueText_(qaCueRequireCompleteFaq_(offer.getDataRange().getValues()));
    const actual = DocumentApp.openById(cue.getId()).getBody().getText();
    const pdf = filesByName_(folder, qaCuePdfFileName_(隔離通し_生成_案件名));
    if (actual !== expected || !pdf.hasNext() || pdf.next().getSize() < 1 || pdf.hasNext()) throw new Error('隔離Q&AカンペDoc/PDFの再読照合に失敗しました');
    return { ok:true, hearingId:隔離通し_生成_ヒアリングID, backupName:backupName, cueDocId:cue.getId(), cuePdf:true, faqOrder:['FAQ1','FAQ2','FAQ3'], deliveryCalled:false, approvalChanged:false, audioStarted:false };
  });
}

/** Read-only evidence for the already-created isolation cue artifacts. */
function readIsolationCueEvidence() {
  const cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function() {
    const project = 隔離通し_生成_案件名 + '_ヒアリングシート';
    const folder = caseFolderIfExists_(cfg, project);
    if (!folder) throw new Error('隔離案件フォルダがありません');
    const docIt = filesByName_(folder, qaCueFileName_(隔離通し_生成_案件名));
    const pdfIt = filesByName_(folder, qaCuePdfFileName_(隔離通し_生成_案件名));
    if (!docIt.hasNext() || !pdfIt.hasNext()) throw new Error('隔離Q&AカンペDoc/PDFがありません');
    const docFile = docIt.next(), pdfFile = pdfIt.next();
    if (docIt.hasNext() || pdfIt.hasNext()) throw new Error('隔離Q&AカンペDoc/PDFが複数あります');
    const offer = SpreadsheetApp.openById(隔離通し_生成_ヒアリングID).getSheetByName(HEARING_TABS.OFFER);
    const pairs = qaCueRequireCompleteFaq_(offer.getDataRange().getValues()), expected = qaCueText_(pairs);
    const actual = DocumentApp.openById(docFile.getId()).getBody().getText();
    return { ok:true, caseName:隔離通し_生成_案件名, folderId:folder.getId(),
      doc:{ id:docFile.getId(), url:docFile.getUrl(), mime:docFile.getMimeType(), chars:actual.length, exactFaqText:actual === expected },
      pdf:{ id:pdfFile.getId(), url:pdfFile.getUrl(), mime:pdfFile.getMimeType(), size:pdfFile.getSize() },
      faqOrder:pairs.map(function(_, i) { return 'FAQ ' + (i + 1); }),
      deliveryReference:{ requiredDocName:qaCueFileName_(隔離通し_生成_案件名), requiredPdfName:qaCuePdfFileName_(隔離通し_生成_案件名), deliveryExecuted:false }
    };
  });
}

function readIsolationFaqExtraction() {
  const cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function() {
    const offer = SpreadsheetApp.openById(隔離通し_生成_ヒアリングID).getSheetByName(HEARING_TABS.OFFER);
    const values = offer.getDataRange().getValues();
    const strict = qaCueFaqStatusFromValues_(values), loose = qaPairsFromValues_(values);
    return { ok:true, strictState:strict.state, strictReason:strict.reason, strictPairs:strict.pairs.map(function(p) { return { qChars:String(p.q||'').length, aChars:String(p.a||'').length }; }), loosePairs:loose.map(function(p) { return { qChars:String(p.q||'').length, aChars:String(p.a||'').length }; }), rowCount:values.length };
  });
}

function readIsolationFaqCells() {
  const cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function() {
    const offer = SpreadsheetApp.openById(隔離通し_生成_ヒアリングID).getSheetByName(HEARING_TABS.OFFER), values = offer.getDataRange().getValues();
    const norm = function(v) { return String(v == null ? '' : v).replace(/[:：\s]/g, ''); }, out = []; let start = -1, q = null;
    for (let r=0;r<values.length;r++) if (values[r].some(function(v){return /よくある質問[（(]\s*Q\s*[&＆]\s*A\s*[）)]/i.test(String(v));})) { start=r+1; break; }
    if(start<0) throw new Error('FAQ見出しがありません');
    for (let r=start;r<values.length&&out.length<3;r++) {
      const non=values[r].map(String).filter(function(x){return x.trim();}); if(non.length&&/^([⒈-⒛]|[0-9０-９]+[\.．])\s*/.test(non[0])) break;
      const qi=values[r].findIndex(function(v){return norm(v)==='よくある質問';}), ai=values[r].findIndex(function(v){return norm(v)==='記入欄';});
      if(qi>=0) q={row:r+1,col:qi+2,value:String(values[r][qi+1]||'')};
      if(ai>=0&&q){out.push({faq:out.length+1,q:q,a:{row:r+1,col:ai+2,value:String(values[r][ai+1]||'')}});q=null;}
    }
    return {ok:true, headingRow:start, cells:out, sectionRowsJson:JSON.stringify(values.slice(start, Math.min(values.length, start+12)).map(function(row,i){return {row:start+i+1, values:row.map(function(v){return String(v||'');})};}))};
  });
}

// Manual fixed-isolation write test for the structured FAQ layout.  Like the
// companion fixture test above, its writes are limited to the fixed fixture
// after backup; normal flows, delivery, notification, and audio are excluded.
function runIsolationStructuredFaqNormalCase() {
  const cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function() {
    const root=DriveApp.getFolderById(隔離通し_生成_rootID), ss=SpreadsheetApp.openById(隔離通し_生成_ヒアリングID), offer=ss.getSheetByName(HEARING_TABS.OFFER);
    if(root.getName()!==隔離通し_生成_root名||!offer) throw new Error('固定隔離rootまたはhearingが一致しません');
    const range=offer.getRange(17,1,4,4), before=range.getValues();
    const q1=offer.getRange(15,3,2,2).getValues();
    const blank=before.every(function(row){return row.every(function(v){return String(v==null?'':v).trim()==='';});});
    if(!blank||String(q1[0][0]).trim()!=='よくある質問'||!String(q1[0][1]).trim()||String(q1[1][0]).trim()!=='記入欄'||!String(q1[1][1]).trim()) throw new Error('FAQ1または空FAQ2/3領域の実行前提を満たしません');
    const backupName='【控え】隔離通しFAQ2-3構造追加前.json', old=filesByName_(root,backupName);
    if(old.hasNext()) throw new Error('隔離FAQ構造追加前バックアップが既にあります');
    root.createFile(backupName,JSON.stringify({schema:1,hearingId:隔離通し_生成_ヒアリングID,range:'A17:D20',before:before,changes:[{cell:'C17',value:'よくある質問'},{cell:'D17',value:'この隔離通しテストの架空設定で、導入前に準備するものはありますか？'},{cell:'C18',value:'記入欄'},{cell:'D18',value:'架空テスト用の入力として、現在お使いの商談記録や提案資料があれば確認材料として使えます。'},{cell:'C19',value:'よくある質問'},{cell:'D19',value:'この隔離通しテストの架空設定で、少人数のチームでも始められますか？'},{cell:'C20',value:'記入欄'},{cell:'D20',value:'架空テスト用の入力として、週次の振り返りと一つの改善テーマから始められます。'}]}),MimeType.PLAIN_TEXT);
    offer.getRange(17,3,4,2).setValues([['よくある質問','この隔離通しテストの架空設定で、導入前に準備するものはありますか？'],['記入欄','架空テスト用の入力として、現在お使いの商談記録や提案資料があれば確認材料として使えます。'],['よくある質問','この隔離通しテストの架空設定で、少人数のチームでも始められますか？'],['記入欄','架空テスト用の入力として、週次の振り返りと一つの改善テーマから始められます。']]);
    SpreadsheetApp.flush(); const values=offer.getDataRange().getValues(), strict=qaCueFaqStatusFromValues_(values), loose=qaPairsFromValues_(values);
    if(strict.state!=='完成'||strict.pairs.length!==3||loose.length!==3) throw new Error('既存FAQ抽出器の3組readbackに失敗しました');
    const project=隔離通し_生成_案件名+'_ヒアリングシート', folder=caseFolderIfExists_(cfg,project); if(!folder) throw new Error('隔離案件フォルダがありません');
    const cue=syncQaCue_(cfg,project), expected=qaCueText_(strict.pairs), actual=DocumentApp.openById(cue.getId()).getBody().getText(), pdfs=filesByName_(folder,qaCuePdfFileName_(隔離通し_生成_案件名));
    if(actual!==expected||!pdfs.hasNext()) throw new Error('カンペDoc/PDF実物照合に失敗しました'); const pdf=pdfs.next(); if(pdfs.hasNext()||pdf.getSize()<1) throw new Error('カンペPDF実物照合に失敗しました');
    return {ok:true,backupName:backupName,changedCells:['C17','D17','C18','D18','C19','D19','C20','D20'],faqPairs:3,doc:{id:cue.getId(),url:cue.getUrl(),chars:actual.length,exact:true},pdf:{id:pdf.getId(),url:pdf.getUrl(),size:pdf.getSize()},deliveryExecuted:false,notificationSent:false,approvalChanged:false};
  });
}
