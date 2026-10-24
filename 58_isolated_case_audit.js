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
