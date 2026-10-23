// Durable approved layout input. Object IDs are intentionally not retained across regeneration.
const GEOMETRY_FILE_PREFIX = '【配置指定】';
const GEOMETRY_SCHEMA = 'consultation-slide-geometry/v2';
const GEOMETRY_MAP_PREFIX = '【スライド対応】';
function geometryOf_(e) { return { left:e.getLeft(), top:e.getTop(), width:e.getWidth(), height:e.getHeight() }; }
function sameGeometry_(a,b) { return !!a && !!b && ['left','top','width','height'].every(function(k){return Number(a[k]) === Number(b[k]);}); }
function validGeometry_(g) { return !!g && ['left','top','width','height'].every(function(k){return typeof g[k] === 'number' && isFinite(g[k]) && g[k] >= 0;}) && g.width > 0 && g.height > 0; }
function elementLabel_(e) { return {type:String(e.getPageElementType()),title:String(e.getTitle()||''),description:String(e.getDescription()||'')}; }
function activeGeometryFile_(folder, company) { const it=filesByName_(folder,GEOMETRY_FILE_PREFIX+company+'.json'); if(!it.hasNext())return null;const f=it.next();if(it.hasNext())throw new Error('有効な配置指定が複数あります');return f; }
// This map is authored during generation while template slide references are still
// available. It is deliberately separate from the durable overrides: it is proof
// for a particular generated deck, not an input that may alter it.
function writeGeometrySlideMap_(folder, company, presentationId, slideFor, qaSlide, pres) {
  const live={};pres.getSlides().forEach(function(s){live[s.getObjectId()]=true;}); const pages=[];
  for(let no=1;no<=106;no++){const s=slideFor(no);if(s&&live[s.getObjectId()])pages.push({slideNo:no,pageId:s.getObjectId()});}
  if(qaSlide&&live[qaSlide.getObjectId()])pages.push({slideNo:QA_NO,pageId:qaSlide.getObjectId(),generated:'qa'});
  const text=JSON.stringify({schema:'consultation-slide-map/v1',caseName:company,presentationId:presentationId,pages:pages}),name=GEOMETRY_MAP_PREFIX+company+'.json',it=filesByName_(folder,name);
  if(!it.hasNext()){folder.createFile(name,text,MimeType.PLAIN_TEXT);return;}const f=it.next();if(it.hasNext())throw new Error('配置対応表が複数あります');f.setContent(text);
}
function selectorMatches_(e,s) { const l=elementLabel_(e);return l.type===s.type && l.title===s.title && l.description===s.description && sameGeometry_(geometryOf_(e),s.templateGeometry); }
function applyPersistentGeometry_(folder, company, slideForNo, warnings) {
  let f;try{f=activeGeometryFile_(folder,company);}catch(e){warnings.push('配置指定を適用しません: '+e);return;}if(!f)return;
  let c;try{c=JSON.parse(f.getBlob().getDataAsString('UTF-8'));}catch(e){warnings.push('配置指定JSONを読めません');return;}
  if(!c||c.schema!==GEOMETRY_SCHEMA||c.caseName!==company||!Array.isArray(c.records)){warnings.push('配置指定の案件またはschemaが一致しません');return;}
  c.records.forEach(function(r){if(!r||r.status!=='active'||!Number.isFinite(r.slideNo)||!r.selector||!validGeometry_(r.after)){warnings.push('配置指定の形式が不正なため1件適用しません');return;}const s=slideForNo(r.slideNo);if(!s){warnings.push('No.'+r.slideNo+' の対応を証明できないため配置を適用しません');return;}const m=s.getPageElements().filter(function(e){try{return selectorMatches_(e,r.selector);}catch(x){return false;}});if(m.length!==1){warnings.push('No.'+r.slideNo+' の要素対応が一意でないため配置を適用しません');return;}try{m[0].setLeft(r.after.left).setTop(r.after.top).setWidth(r.after.width).setHeight(r.after.height);}catch(e){warnings.push('No.'+r.slideNo+' の配置適用に失敗: '+e);}});
}
