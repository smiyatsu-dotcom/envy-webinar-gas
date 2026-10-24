// Durable approved layout input. Object IDs are intentionally not retained across regeneration.
const GEOMETRY_FILE_PREFIX = '【配置指定】';
const GEOMETRY_SCHEMA = 'consultation-slide-geometry/v2';
const GEOMETRY_MAP_PREFIX = '【スライド対応】';
const GEOMETRY_MAP_SCHEMA = 'consultation-slide-map/v2';
function geometryOf_(e) { return { left:e.getLeft(), top:e.getTop(), width:e.getWidth(), height:e.getHeight() }; }
function sameGeometry_(a,b) { return !!a && !!b && ['left','top','width','height'].every(function(k){return Number(a[k]) === Number(b[k]);}); }
function validGeometry_(g) { return !!g && ['left','top','width','height'].every(function(k){return typeof g[k] === 'number' && isFinite(g[k]) && g[k] >= 0;}) && g.width > 0 && g.height > 0; }
function elementLabel_(e) { return {type:String(e.getPageElementType()),title:String(e.getTitle()||''),description:String(e.getDescription()||'')}; }
function activeGeometryFile_(folder, company) { const it=filesByName_(folder,GEOMETRY_FILE_PREFIX+company+'.json'); if(!it.hasNext())return null;const f=it.next();if(it.hasNext())throw new Error('有効な配置指定が複数あります');return f; }
// The map is output by the generator, not supplied by Discord.  It contains
// only a generator-authored target role: the primary replaced IMAGE or an
// editable heading SHAPE recognized on that generated page. A role is not inferred from
// a user request. Object IDs are used only while authoring this map and are
// deliberately omitted from it: normal read resolves the selector again.
function geometryMapSelector_(element, role) {
  const selector = Object.assign(elementLabel_(element), { templateGeometry:geometryOf_(element), role:role });
  if (selector.type === 'IMAGE' && role === 'primary-image') return selector;
  if (selector.type !== 'SHAPE' || role !== 'heading') return null;
  let text = '';
  try { text = String(element.asShape().getText().asString() || ''); } catch (e) { }
  // Preserve the exact live value; normal resolution compares it byte-for-byte.
  // Whitespace-only and overlong headings are intentionally ineligible.
  if (!text.trim() || text.length > 2000) return null;
  selector.text = text;
  return selector;
}
// A durable layout record is the authoritative selector when a regenerated
// page already has an approved correction.  Object IDs are still never
// persisted: this resolves the one live element only while the generator is
// producing the map.  In particular, it prevents No.40's generic edited
// heading and its formatting-only heading scan from creating two candidates.
function geometryActiveRecordForSlide_(folder, company, slideNo) {
  const f = activeGeometryFile_(folder, company);
  if (!f) return null;
  let spec;
  try { spec = JSON.parse(f.getBlob().getDataAsString('UTF-8')); }
  catch (e) { throw new Error('配置指定JSONを読めません'); }
  if (!spec || spec.schema !== GEOMETRY_SCHEMA || spec.caseName !== company || !Array.isArray(spec.records)) {
    throw new Error('配置指定の案件またはschemaが一致しません');
  }
  const matches = spec.records.filter(function(record) {
    return record && record.status === 'active' && Number(record.slideNo) === Number(slideNo) &&
      record.selector && validGeometry_(record.selector.templateGeometry) && validGeometry_(record.after);
  });
  if (matches.length > 1) throw new Error('No.' + slideNo + ' の有効な配置指定が複数あります');
  return matches[0] || null;
}
function geometryTargetFromActiveRecord_(folder, company, slideNo, slide) {
  const record = geometryActiveRecordForSlide_(folder, company, slideNo);
  if (!record) return { record: null, element: null };
  const elements = slide && typeof slide.getPageElements === 'function' ? slide.getPageElements() : [];
  const matches = elements.filter(function(element) {
    try {
      return String(element.getPageElementType()) === 'SHAPE' &&
        !!elementText_(element).trim() && selectorMatches_(element, record.selector);
    } catch (e) { return false; }
  });
  return { record: record, element: matches.length === 1 ? matches[0] : null };
}
function writeGeometrySlideMap_(folder, company, presentationId, pageIds, pres, targetElements) {
  const live={};pres.getSlides().forEach(function(s){live[s.getObjectId()]=s;}); const pages=[];
  Object.keys(pageIds || {}).map(Number).sort(function(a,b){return a-b;}).forEach(function(no){
    const pageId=pageIds[no],slide=live[pageId];if(!slide)return;
    const page={slideNo:no,pageId:pageId};const selected=(targetElements&&targetElements[no]||[]).filter(function(item,i,a){return item&&typeof item.id==='string'&&typeof item.type==='string'&&typeof item.role==='string'&&a.findIndex(function(other){return other&&other.id===item.id;})===i;});
    // More than one generator-authorized target leaves this slide intentionally
    // ineligible: a normal request names a slide, never a client object ID.
    if(selected.length===1){const target=selected[0],candidates=slide.getPageElements().filter(function(element){return element.getObjectId()===target.id&&String(element.getPageElementType())===target.type;});if(candidates.length===1&&validGeometry_(geometryOf_(candidates[0]))){const selector=geometryMapSelector_(candidates[0],target.role);if(selector)page.targetSelector=selector;}}
    pages.push(page);
  });
  const revision=DriveApp.getFileById(presentationId).getLastUpdated().toISOString();
  const text=JSON.stringify({schema:GEOMETRY_MAP_SCHEMA,caseName:company,presentationId:presentationId,presentationRevision:revision,pages:pages}),name=GEOMETRY_MAP_PREFIX+company+'.json',it=filesByName_(folder,name);
  if(!it.hasNext()){folder.createFile(name,text,MimeType.PLAIN_TEXT);return;}const f=it.next();if(it.hasNext())throw new Error('配置対応表が複数あります');f.setContent(text);
}
function elementText_(e) { try { return String(e.asShape().getText().asString() || ''); } catch (x) { return ''; } }
function selectorMatches_(e,s) { const l=elementLabel_(e);return l.type===s.type && l.title===s.title && l.description===s.description && sameGeometry_(geometryOf_(e),s.templateGeometry) && (!Object.prototype.hasOwnProperty.call(s,'text') || (l.type==='SHAPE' && elementText_(e)===s.text)); }
function applyPersistentGeometry_(folder, company, slideForNo, warnings) {
  let f;try{f=activeGeometryFile_(folder,company);}catch(e){warnings.push('配置指定を適用しません: '+e);return;}if(!f)return;
  let c;try{c=JSON.parse(f.getBlob().getDataAsString('UTF-8'));}catch(e){warnings.push('配置指定JSONを読めません');return;}
  if(!c||c.schema!==GEOMETRY_SCHEMA||c.caseName!==company||!Array.isArray(c.records)){warnings.push('配置指定の案件またはschemaが一致しません');return;}
  c.records.forEach(function(r){if(!r||r.status!=='active'||!Number.isFinite(r.slideNo)||!r.selector||!validGeometry_(r.after)){warnings.push('配置指定の形式が不正なため1件適用しません');return;}const s=slideForNo(r.slideNo);if(!s){warnings.push('No.'+r.slideNo+' の対応を証明できないため配置を適用しません');return;}const m=s.getPageElements().filter(function(e){try{return selectorMatches_(e,r.selector);}catch(x){return false;}});if(m.length!==1){warnings.push('No.'+r.slideNo+' の要素対応が一意でないため配置を適用しません');return;}try{m[0].setLeft(r.after.left).setTop(r.after.top).setWidth(r.after.width).setHeight(r.after.height);}catch(e){warnings.push('No.'+r.slideNo+' の配置適用に失敗: '+e);}});
}
