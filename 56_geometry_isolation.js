// Fixed, no-argument integration test. Properties are administrator-set once;
// chat, runner arguments, and case files cannot choose IDs or folders.
// This setup is intentionally separate from the runner: it only creates the
// synthetic fixture under the fixed isolation root.  It never reads a
// customer sheet, writes the regular input/output roots, or sets properties.
function スライド_隔離fixtureを作る() {
  const rootId = '1GUmxo4J9cj0WYdh4IPj2lBi3RSxiMw73';
  const caseName = '【隔離試験】架空配置永続化案件';
  const root = DriveApp.getFolderById(rootId);
  if (root.getName() !== '【隔離試験】配置永続化') throw new Error('隔離専用の出力rootではありません');
  const existing = filesByName_(root, '【隔離試験】構成シート_架空配置永続化');
  if (existing.hasNext()) {
    const file = existing.next();
    const ss = SpreadsheetApp.openById(file.getId());
    if (ss.getSheetByName(TABS.SETTINGS) && ss.getSheetByName(TABS.INTRO) &&
        ss.getSheetByName(TABS.PROBLEM) && ss.getSheetByName(TABS.STORY) && ss.getSheetByName(TABS.OFFER)) {
      return { created: false, sheetId: ss.getId(), sheetName: ss.getName(), caseName: caseName, outputRootId: rootId };
    }
    throw new Error('同名の隔離fixtureが生成仕様を満たしません。手動で確認してください');
  }
  const ss = SpreadsheetApp.create('【隔離試験】構成シート_架空配置永続化');
  DriveApp.getFileById(ss.getId()).moveTo(root);
  const settings = ss.getSheets()[0];
  settings.setName(TABS.SETTINGS);
  settings.getRange('A1:A5').setValues([['案件名'], ['承認'], ['ステータス'], ['ヒアリングファイルID'], ['素材フォルダID']]);
  settings.getRange('B1').setValue(caseName);
  settings.getRange('B2').setValue(true);
  settings.getRange('B3').setValue('隔離fixture（架空案件・顧客入力なし）');
  settings.getRange('B4').setValue('ISOLATION_SYNTHETIC_HEARING_ONLY');
  settings.getRange('B5').setValue('');
  writeContentTabs_(ss, {}, {}, {}, {});
  SpreadsheetApp.flush();
  return { created: true, sheetId: ss.getId(), sheetName: ss.getName(), caseName: caseName, outputRootId: rootId };
}

// 2026-10-23: 失われた旧fixtureの復元ではなく、新しい固定隔離用の再実行入口。
// 引数・任意ID・任意関数は受け付けず、通常 main/顧客root/通知を通さない。
var 隔離通し_生成_案件名 = '隔離通し診断テスト社';
var 隔離通し_生成_rootID = '1oQSEwyopPcvX20xuxNDHy5Kh4TfvtCBO';
var 隔離通し_生成_親rootID = '1GUmxo4J9cj0WYdh4IPj2lBi3RSxiMw73';
var 隔離通し_生成_root名 = '【隔離通し】診断から資料音声_20261023';
var 隔離通し_生成_台帳ID = '16UgS3scd3X2D2Qqhzd1WfJuGkuNcOt5WbekQZLyKDpg';
var 隔離通し_生成_開始箱ID = '1Mxch8d4SEbddMfx-C2slbmQWWa1Uz16u';
var 隔離通し_生成_生成物rootID = '18-xhLrskU8M9DvW2Rurqg5ST_NrOiSM2';
var 隔離通し_生成_ヒアリングID = '1ur8hPUElZMrKuitnDo1j1wzr54EkmRY9efNMaEQFErs';

function 隔離通し_生成の親を確かめる_(fileOrFolder, parentId, label) {
  var parents = fileOrFolder.getParents();
  if (!parents.hasNext() || parents.next().getId() !== parentId || parents.hasNext()) {
    throw new Error(label + ' が固定隔離root直下ではありません');
  }
}

function 隔離通し_生成設定を作る_() {
  var root = DriveApp.getFolderById(隔離通し_生成_rootID);
  if (root.getName() !== 隔離通し_生成_root名) throw new Error('隔離通しroot名が一致しません');
  隔離通し_生成の親を確かめる_(root, 隔離通し_生成_親rootID, '隔離通しroot');
  var startBox = DriveApp.getFolderById(隔離通し_生成_開始箱ID);
  隔離通し_生成の親を確かめる_(startBox, root.getId(), '隔離通し開始箱');
  if (startBox.getName() !== '【出力】') throw new Error('隔離通し開始箱名が一致しません');
  var output = DriveApp.getFolderById(隔離通し_生成_生成物rootID);
  隔離通し_生成の親を確かめる_(output, root.getId(), '隔離通し生成物root');
  if (output.getName() !== '【出力】生成物') throw new Error('隔離通し生成物root名が一致しません');
  var ledgerFile = DriveApp.getFileById(隔離通し_生成_台帳ID);
  隔離通し_生成の親を確かめる_(ledgerFile, root.getId(), '隔離通し台帳');

  // 実行前に専用台帳がこの一件だけを新フローとして持つことを実物照合する。
  // これにより processNewFlow_ の通常走査が顧客案件へ広がらない。
  var rows = SpreadsheetApp.openById(隔離通し_生成_台帳ID).getSheets()[0].getDataRange().getValues();
  var matches = rows.slice(1).filter(function (row) {
    return String(row[0] || '').trim() === 隔離通し_生成_ヒアリングID &&
      String(row[1] || '').trim() === 隔離通し_生成_案件名 + '_ヒアリングシート' &&
      String(row[流れ_列 - 1] || '').trim() === 流れ_新;
  });
  if (matches.length !== 1 || rows.slice(1).filter(function (row) { return String(row[流れ_列 - 1] || '').trim() === 流れ_新; }).length !== 1) {
    throw new Error('隔離通し台帳が固定の新フロー1件ではありません');
  }

  // 通常設定からテンプレート/素材の読取先だけを借り、書込先と台帳は必ず専用fixtureへ差し替える。
  var normal = getConfig_();
  if (normal.OUTPUT_FOLDER_ID === output.getId() || normal.MGMT_SHEET_ID === 隔離通し_生成_台帳ID) {
    throw new Error('通常設定と隔離fixtureが分離されていません');
  }
  return Object.assign({}, normal, {
    ISOLATION_TEST: true,
    OUTPUT_FOLDER_ID: output.getId(),
    MGMT_SHEET_ID: 隔離通し_生成_台帳ID,
    ISOLATION_START_BOX_ID: startBox.getId()
  });
}

function 隔離通し_同名ファイル1件_(folder, name, label) {
  var it = folder.getFilesByName(name);
  if (!it.hasNext()) return null;
  var file = it.next();
  if (it.hasNext()) throw new Error(label + ' が複数あります');
  return file;
}

function 隔離通し_正規開始を引き渡す_(cfg) {
  var startBox = DriveApp.getFolderById(cfg.ISOLATION_START_BOX_ID);
  var startName = '【開始】' + 隔離通し_生成_案件名 + '.json';
  var start = 隔離通し_同名ファイル1件_(startBox, startName, '隔離通し開始JSON');
  if (!start) return { handed: false, reason: '開始箱に未処理の正規開始JSONがありません' };
  var payload = 開始を読む_(start.getBlob().getDataAsString('UTF-8'));
  if (payload.案件名 !== 隔離通し_生成_案件名 || payload.ヒアリングシートID !== 隔離通し_生成_ヒアリングID) {
    throw new Error('隔離通し開始JSONの案件またはヒアリングIDが固定fixtureと一致しません');
  }

  var generatedRoot = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  var old = generatedRoot.getFoldersByName(隔離通し_生成_案件名);
  if (old.hasNext()) {
    var prior = old.next();
    if (old.hasNext()) throw new Error('隔離通し生成物の案件フォルダが複数あります');
    var root = DriveApp.getFolderById(隔離通し_生成_rootID);
    var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
    var archive = root.createFolder('【再試験控え】' + 隔離通し_生成_案件名 + '_' + stamp);
    // 削除・上書きではなく、旧済開始・状態・成果物を丸ごと控えへ保全する。
    prior.moveTo(archive);
  }
  var destination = generatedRoot.createFolder(隔離通し_生成_案件名);
  // JSON本文は変えず、正規開始処理が読む案件フォルダへだけ移す。
  start.moveTo(destination);
  return { handed: true, startFileId: start.getId(), outputFolderId: destination.getId(), outputUrl: destination.getUrl() };
}

/** 固定fixtureの正規開始JSONを既存登録処理へ安全に引き渡す。 */
function 隔離通し生成を開始する() {
  var cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function () {
    var handoff = 隔離通し_正規開始を引き渡す_(cfg);
    if (!handoff.handed) return Object.assign({ ok: true, phase: 'start' }, handoff);
    processStartOrders_();
    return Object.assign({ ok: true, phase: 'start', caseName: 隔離通し_生成_案件名, outputRootId: cfg.OUTPUT_FOLDER_ID, managementSheetId: cfg.MGMT_SHEET_ID }, handoff);
  });
}

/** 固定fixtureの1件だけを通常の新フロー処理で進める。 */
function 隔離通し生成を進める() {
  var cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function () {
    processNewFlow_();
    return { ok: true, phase: 'advance', caseName: 隔離通し_生成_案件名, outputRootId: cfg.OUTPUT_FOLDER_ID, managementSheetId: cfg.MGMT_SHEET_ID };
  });
}

function isolationUntouchedFingerprint_(slide, omit) {
  return slide.getPageElements().filter(function(e){return !omit(e);}).map(function(e){
    return JSON.stringify({label:elementLabel_(e),geometry:geometryOf_(e)});
  }).sort();
}

function スライド_隔離配置永続化を試す() {
  const p=PropertiesService.getScriptProperties();
  const sheetId=p.getProperty('GEOMETRY_ISOLATION_KOUSEI_ID');
  const rootId=p.getProperty('GEOMETRY_ISOLATION_OUTPUT_ROOT_ID');
  const caseName=p.getProperty('GEOMETRY_ISOLATION_CASE_NAME');
  if(!sheetId||!rootId||!caseName) throw new Error('隔離配置試験の固定fixture設定がありません');
  const prod=getConfig_();
  if(rootId===prod.OUTPUT_FOLDER_ID) throw new Error('顧客の【出力】生成物は隔離試験に使えません');
  const root=DriveApp.getFolderById(rootId);
  if(root.getName()!=='【隔離試験】配置永続化') throw new Error('隔離専用の出力rootではありません');
  const ss=SpreadsheetApp.openById(sheetId);
  if(String(ss.getName()).indexOf('【隔離試験】')!==0) throw new Error('隔離fixtureの構成シートではありません');
  const delivery=DriveApp.getFolderById(prod.OUTPUT_FOLDER_ID).getParents().next().getFoldersByName('【納品】');
  if(delivery.hasNext()&&delivery.next().getId()===rootId) throw new Error('納品フォルダを隔離試験に使えません');
  const folder=caseFolder_(Object.assign({},prod,{OUTPUT_FOLDER_ID:rootId}),caseName);
  // A prior isolated run may have left its own test record.  Neutralize only
  // that fixed file before recording the first deck's untouched template
  // geometry; customer records and all other case files remain untouched.
  const geometryName='【配置指定】'+caseCompany_(caseName)+'.json';
  const priorGeometry=filesByName_(folder,geometryName);
  if(priorGeometry.hasNext()){const f=priorGeometry.next();if(priorGeometry.hasNext())throw new Error('隔離配置指定が複数あります');f.setContent(JSON.stringify({schema:GEOMETRY_SCHEMA,caseName:caseCompany_(caseName),records:[]}));}
  const first=generateSlides_(sheetId,{安全:true,isolationOutputFolderId:rootId});
  const mapFile=filesByName_(folder,'【スライド対応】'+caseCompany_(caseName)+'.json');
  if(!mapFile.hasNext()) throw new Error('1回目の対応表がありません');
  const map=JSON.parse(mapFile.next().getBlob().getDataAsString('UTF-8'));
  if(map.presentationId!==first.pres.getId()||!Array.isArray(map.pages)||!map.pages.length) throw new Error('1回目の対応表がdeckと一致しません');
  // Author one deliberately small, fixed-only layout change from the live first
  // deck.  The selector carries template geometry, so the second deck must
  // prove the same template object before the override can apply.
  const page1=map.pages.filter(function(x){return x.slideNo===1;})[0];
  const slide1=page1&&first.pres.getSlides().filter(function(s){return s.getObjectId()===page1.pageId;})[0];
  if(!slide1) throw new Error('No.1の生成済み対応を読めません');
  let target=null, selector=null;
  slide1.getPageElements().some(function(e){
    try {
      const g=geometryOf_(e), s=Object.assign(elementLabel_(e),{templateGeometry:g});
      if(!validGeometry_(g)) return false;
      if(slide1.getPageElements().filter(function(x){return selectorMatches_(x,s);}).length!==1) return false;
      target=e; selector=s; return true;
    } catch(ignore) { return false; }
  });
  if(!target) throw new Error('No.1に一意な配置検証対象がありません');
  const after=Object.assign({},selector.templateGeometry,{left:selector.templateGeometry.left+1});
  const untouchedBefore=map.pages.map(function(page){
    const slide=first.pres.getSlides().filter(function(s){return s.getObjectId()===page.pageId;})[0];
    if(!slide) throw new Error('1回目の対象外要素を読めません: No.'+page.slideNo);
    return {slideNo:page.slideNo,fingerprint:isolationUntouchedFingerprint_(slide,function(e){return page.slideNo===1&&selectorMatches_(e,selector);})};
  });
  const geometryText=JSON.stringify({schema:GEOMETRY_SCHEMA,caseName:caseCompany_(caseName),records:[{status:'active',slideNo:1,selector:selector,after:after}]});
  const oldGeometry=filesByName_(folder,geometryName);
  if(oldGeometry.hasNext()){const f=oldGeometry.next();if(oldGeometry.hasNext())throw new Error('隔離配置指定が複数あります');f.setContent(geometryText);}else{folder.createFile(geometryName,geometryText,MimeType.PLAIN_TEXT);}
  const second=generateSlides_(sheetId,{安全:true,isolationOutputFolderId:rootId});
  const latest=JSON.parse(filesByName_(folder,'【スライド対応】'+caseCompany_(caseName)+'.json').next().getBlob().getDataAsString('UTF-8'));
  if(latest.presentationId!==second.pres.getId()||latest.presentationId===map.presentationId) throw new Error('2回目の対応表を読み戻せません');
  if(latest.pages.length!==map.pages.length||latest.pages.map(function(p){return p.slideNo;}).join(',')!==map.pages.map(function(p){return p.slideNo;}).join(',')) throw new Error('対象外ページ数または順序が変わりました');
  const page2=latest.pages.filter(function(x){return x.slideNo===1;})[0];
  const slide2=page2&&second.pres.getSlides().filter(function(s){return s.getObjectId()===page2.pageId;})[0];
  // selectorMatches_ is deliberately used only before mutation. Once applied,
  // templateGeometry is no longer expected to match. Read the one post-apply
  // element by its stable labels plus its exact requested final geometry.
  const labeled=slide2?slide2.getPageElements().filter(function(e){var l=elementLabel_(e);return l.type===selector.type&&l.title===selector.title&&l.description===selector.description;}):[];
  // The SlidesApp getters and setters both expose point values.  Do not hide a
  // transform/round-trip difference behind an arbitrary epsilon: capture the
  // exact readback and require numeric equality for every requested field.
  const readbacks=labeled.map(function(e){return geometryOf_(e);});
  const applied=labeled.filter(function(e){var g=geometryOf_(e);return ['left','top','width','height'].every(function(k){return Number(g[k])===Number(after[k]);});});
  if(applied.length!==1) throw new Error('2回目の現物geometryが承認済み配置と一致しません: expected='+JSON.stringify(after)+', actual='+JSON.stringify(readbacks)+', labels='+labeled.length+', expectedGeometry='+applied.length);
  untouchedBefore.forEach(function(before){
    const page=latest.pages.filter(function(p){return p.slideNo===before.slideNo;})[0];
    const slide=page&&second.pres.getSlides().filter(function(s){return s.getObjectId()===page.pageId;})[0];
    if(!slide) throw new Error('2回目の対象外要素を読めません: No.'+before.slideNo);
    const actual=isolationUntouchedFingerprint_(slide,function(e){
      if(before.slideNo!==1) return false;
      const l=elementLabel_(e),g=geometryOf_(e);
      return l.type===selector.type&&l.title===selector.title&&l.description===selector.description&&['left','top','width','height'].every(function(k){return Number(g[k])===Number(after[k]);});
    });
    if(JSON.stringify(actual)!==JSON.stringify(before.fingerprint)) throw new Error('対象外要素のgeometryまたは構成が変わりました: No.'+before.slideNo);
  });
  return {ok:true,caseName:caseName,outputRootId:rootId,firstDeckId:map.presentationId,secondDeckId:latest.presentationId,pageCount:latest.pages.length,expectedGeometry:after,actualGeometry:geometryOf_(applied[0]),geometryVerified:true};
}
