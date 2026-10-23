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
