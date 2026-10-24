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
function isolatedJCurrentText_(slide) {
  try { return String(slide.getNotesPage().getSpeakerNotesShape().getText().asString() || '').replace(/\r\n/g, '\n').replace(/\n+$/g, ''); }
  catch (e) { return ''; }
}
function isolatedJCurrentAudioText_(notes) {
  const parts = notes.map(function(text) {
    return String(text || '').replace(/（挙手ありがとうございます。?）/g, '').trim();
  }).filter(function(text) { return !!text; });
  return parts.join('\n\n') + '\n';
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
  const textFile = isolatedJCurrentFile_(folder, c.textFileName, 'J現行音声台本TXT');
  const recordFile = isolatedJCurrentFile_(folder, c.importRecordFileName, 'J台本取込記録');
  const scriptText = scriptFile.getBlob().getDataAsString('UTF-8');
  隔離通し_J台本_本文を検証する_(scriptText);
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
      mapping.pages.some(function(page, index) { return !page || page.slideNo !== index + 1 || typeof page.pageId !== 'string'; })) {
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
    scriptHash: 隔離通し_J台本_sha256_(scriptText), text: textFile.getBlob().getDataAsString('UTF-8') };
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
    slides.forEach(function(slide) { byId[slide.getObjectId()] = slide; });
    if (slides.length !== 106 || target.mapping.pages.some(function(page) { return !byId[page.pageId]; })) {
      throw new Error('J現行deckが対応表の固定106ページと一致しません');
    }
    const notes = target.mapping.pages.map(function(page) { return isolatedJCurrentText_(byId[page.pageId]); });
    const textMatchesNotes = isolatedJCurrentAudioText_(notes) === target.text;
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
    return { ok: textMatchesNotes && geometryRetained, mode: 'read-only-current-j-artifact-audit',
      caseName: 隔離通し_生成_案件名, folderId: target.folder.getId(), deckId: ISOLATED_J_CURRENT_AUDIT_.deckId,
      slideCount: slides.length, scriptSha256: target.scriptHash, notesTxtMatch: textMatchesNotes,
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
