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
