const fs=require('fs'),path=require('path'),vm=require('vm'),assert=require('assert'),root=path.join(__dirname,'..'),ctx={console,Logger:{log(){}},PropertiesService:{},DriveApp:{},SpreadsheetApp:{},SlidesApp:{},DocumentApp:{},MimeType:{},LockService:{},ScriptApp:{},Utilities:{DigestAlgorithm:{SHA_256:'SHA_256'},Charset:{UTF_8:'UTF_8'},computeDigest:(_,value)=>Array.from(Buffer.from(String(value),'utf8'))}};vm.createContext(ctx);fs.readdirSync(root).filter(f=>f.endsWith('.js')).sort().forEach(f=>vm.runInContext(fs.readFileSync(path.join(root,f),'utf8').replace(/^const /gm,'var '),ctx));
const it=a=>{let i=0;return{hasNext:()=>i<a.length,next:()=>a[i++]}};function el(left=1){const o={getLeft:()=>left,getTop:()=>2,getWidth:()=>3,getHeight:()=>4,getPageElementType:()=> 'IMAGE',getTitle:()=>'',getDescription:()=>'',setLeft:v=>(left=v,o),setTop:()=>o,setWidth:()=>o,setHeight:()=>o};return o;}function apply(records,elements){ctx.filesByName_=()=>it([{getBlob:()=>({getDataAsString:()=>JSON.stringify({schema:'consultation-slide-geometry/v2',caseName:'A社',records})})}]);const w=[];ctx.applyPersistentGeometry_({},'A社',n=>n===12?{getPageElements:()=>elements}:null,w);return w;}
const s={type:'IMAGE',title:'',description:'',templateGeometry:{left:1,top:2,width:3,height:4}},r={status:'active',slideNo:12,selector:s,after:{left:9,top:2,width:3,height:4}};let e=el();assert.deepStrictEqual(apply([r],[e]),[]);assert.strictEqual(e.getLeft(),9);assert.deepStrictEqual(apply([r],[el(),el()]).length,1);assert.deepStrictEqual(apply([{...r,after:{left:9}}],[el()]).length,1);console.log('ok persistent geometry: repeated/ambiguous/invalid safely contained');
// Mapping is authored from live generator references, not a post-generation
// array position. It may name the single primary IMAGE or a generator-marked
// editable heading SHAPE; arbitrary shapes and image-baked text stay ineligible.
let saved=[];ctx.filesByName_=()=>it([]);ctx.DriveApp={getFileById:()=>({getLastUpdated:()=>new Date('2026-10-24T00:00:00.000Z')})};const folder={createFile:(n,t)=>saved.push({n,t})};
const image=id=>({getObjectId:()=>id,getPageElementType:()=> 'IMAGE',getTitle:()=>id==='hero'?'hero':'',getDescription:()=>'',getLeft:()=>1,getTop:()=>2,getWidth:()=>3,getHeight:()=>4});
const heading=(id,text)=>({getObjectId:()=>id,getPageElementType:()=> 'SHAPE',getTitle:()=>'',getDescription:()=>'',getLeft:()=>1,getTop:()=>2,getWidth:()=>3,getHeight:()=>4,asShape:()=>({getText:()=>({asString:()=>text})})});
const slide=(id,elements=[])=>({getObjectId:()=>id,getPageElements:()=>elements});
const p1=slide('p1',[image('hero')]),p12=slide('p12',[heading('heading12',' 見出し\n')]),tooLong=slide('too-long',[heading('heading-long','a'.repeat(2001))]),burned=slide('burned',[image('pixels')]),qa=slide('qa');
ctx.writeGeometrySlideMap_(folder,'A社','deck-1',{1:'p1',12:'p12',13:'too-long',41:'burned',104.5:'qa'},{getSlides:()=>[p1,p12,tooLong,burned,qa]},{1:[{id:'hero',type:'IMAGE',role:'primary-image'}],12:[{id:'heading12',type:'SHAPE',role:'heading'}],13:[{id:'heading-long',type:'SHAPE',role:'heading'}],41:[{id:'pixels',type:'SHAPE',role:'heading'}],104.5:[{id:'missing',type:'IMAGE',role:'primary-image'}]});
const map=JSON.parse(saved[0].t);assert.strictEqual(map.schema,'consultation-slide-map/v2');assert.strictEqual(map.presentationRevision,'2026-10-24T00:00:00.000Z');assert.deepStrictEqual(map.pages.map(x=>x.slideNo),[1,12,13,41,104.5]);assert.deepStrictEqual(map.pages[0].targetSelector,{type:'IMAGE',title:'hero',description:'',templateGeometry:{left:1,top:2,width:3,height:4},role:'primary-image'});assert.deepStrictEqual(map.pages[1].targetSelector,{type:'SHAPE',title:'',description:'',templateGeometry:{left:1,top:2,width:3,height:4},role:'heading',text:' 見出し\n'});assert.strictEqual(map.pages[2].targetSelector,undefined);assert.strictEqual(Object.prototype.hasOwnProperty.call(map.pages[0],'objectId'),false);assert.strictEqual(map.pages[3].targetSelector,undefined);
// The map survives JSON save/reopen and its exact selectors resolve again in
// the normal reader. Whitespace is significant; overlong headings are omitted.
const reopened={id:'deck-1',slides:[{id:'p1',elements:[{id:'hero',type:'IMAGE',title:'hero',description:'',geometry:{left:1,top:2,width:3,height:4},text:''}]},{id:'p12',elements:[{id:'heading12',type:'SHAPE',title:'',description:'',geometry:{left:1,top:2,width:3,height:4},text:' 見出し\n'}]}]};
const normalRuntime={resolve:()=>map,revision:()=>map.presentationRevision,open:()=>reopened,presentationId:d=>d.id,slides:d=>d.slides,slideId:s=>s.id,elements:s=>s.elements,element:e=>e,pageSize:()=>({width:720,height:450})};
assert.strictEqual(ctx.normalGeometryRead_({caseName:'A社',slideNo:1},normalRuntime).target.objectId,'hero');assert.strictEqual(ctx.normalGeometryRead_({caseName:'A社',slideNo:12},normalRuntime).target.objectId,'heading12');
let replacement='';ctx.filesByName_=()=>it([{setContent:t=>replacement=t}]);const next=slide('new1',[image('a'),image('b')]);ctx.writeGeometrySlideMap_(folder,'A社','deck-2',{1:'new1'},{getSlides:()=>[next]},{1:[{id:'a',type:'IMAGE',role:'primary-image'},{id:'b',type:'IMAGE',role:'primary-image'}]});const updated=JSON.parse(replacement);assert.strictEqual(updated.presentationId,'deck-2');assert.strictEqual(updated.pages[0].targetSelector,undefined);console.log('ok generated v2 map: live IMAGE or explicit editable heading SHAPE only');
const editable=(id,text,top)=>({getObjectId:()=>id,getTop:()=>top,getWidth:()=>500,getHeight:()=>60,getText:()=>({asString:()=>text})});const genericPage={getShapes:()=>[editable('heading','新しい見出し',30),editable('body','本文',240)]};assert.strictEqual(ctx.generatorEditedHeading_(genericPage,{heading:'旧見出し',body:'本文'}).getObjectId(),'heading');const ambiguousPage={getShapes:()=>[editable('one','見出し1',20),editable('two','見出し2',70)]};assert.strictEqual(ctx.generatorEditedHeading_(ambiguousPage,{}),null);console.log('ok generated v2 map: live IMAGE or explicit editable heading SHAPE only');
// No.40 must follow the existing durable record, not whichever heading the
// common text replacement or the formatting-only top scan happened to see.
const placedHeading=(id,text,g)=>({getObjectId:()=>id,getPageElementType:()=> 'SHAPE',getTitle:()=>'',getDescription:()=>'',getLeft:()=>g.left,getTop:()=>g.top,getWidth:()=>g.width,getHeight:()=>g.height,asShape:()=>({getText:()=>({asString:()=>text})})});
const durableBefore={left:40.44094488188976,top:37.21555118110236,width:639.1181102362204,height:47.40944881889764};
const commonCandidate=placedHeading('common-edited','共通候補',{left:40,top:10,width:300,height:50});
const durableCandidate=placedHeading('durable-heading','永続配置の見出し',durableBefore);
ctx.filesByName_=()=>it([{getBlob:()=>({getDataAsString:()=>JSON.stringify({schema:'consultation-slide-geometry/v2',caseName:'A社',records:[{status:'active',slideNo:40,selector:{type:'SHAPE',title:'',description:'',templateGeometry:durableBefore,text:'永続配置の見出し'},after:{...durableBefore,top:40}}]})})}]);
const no40=ctx.geometryTargetFromActiveRecord_({},'A社',40,slide('p40',[commonCandidate,durableCandidate]));
assert.strictEqual(no40.record.slideNo,40);assert.strictEqual(no40.element.getObjectId(),'durable-heading');assert.notStrictEqual(no40.element.getObjectId(),commonCandidate.getObjectId());
ctx.filesByName_=()=>it([{getBlob:()=>({getDataAsString:()=>JSON.stringify({schema:'consultation-slide-geometry/v2',caseName:'A社',records:[{status:'active',slideNo:40,selector:{type:'SHAPE',title:'',description:'',templateGeometry:durableBefore},after:{...durableBefore,top:40}},{status:'active',slideNo:40,selector:{type:'SHAPE',title:'',description:'',templateGeometry:durableBefore},after:{...durableBefore,top:40}}]})})}]);
assert.throws(()=>ctx.geometryTargetFromActiveRecord_({},'A社',40,slide('p40',[durableCandidate])),/複数/);
console.log('ok No.40 geometry: durable selector wins over a distinct common candidate and duplicate records fail closed');
// The public normal endpoint may exercise exactly one fixed isolation case,
// but no spelling variation or other case can select its separate ledger.
const normalCfg={kind:'normal'}, isolationCfg={ISOLATION_TEST:true,MGMT_SHEET_ID:ctx.隔離通し_生成_台帳ID,OUTPUT_FOLDER_ID:ctx.隔離通し_生成_生成物rootID};
ctx.getConfig_=()=>normalCfg;ctx.隔離通し_生成設定を作る_=()=>isolationCfg;
assert.strictEqual(ctx.normalGeometryConfigForCase_(ctx.ISOLATION_GEOMETRY_CASE),isolationCfg);
assert.strictEqual(ctx.normalGeometryConfigForCase_(ctx.ISOLATION_GEOMETRY_CASE+' '),normalCfg);
assert.strictEqual(ctx.normalGeometryConfigForCase_('A社'),normalCfg);
assert.throws(()=>{ctx.隔離通し_生成設定を作る_=()=>({ISOLATION_TEST:true,MGMT_SHEET_ID:'other',OUTPUT_FOLDER_ID:ctx.隔離通し_生成_生成物rootID});ctx.normalGeometryConfigForCase_(ctx.ISOLATION_GEOMETRY_CASE);},/固定隔離案件/);
console.log('ok normal geometry: only the exact fixed isolation case selects its verified ledger');
// No.40 resolves the exact durable heading when a correction exists. The
// Arial formatter still scans its topmost text, but it is no longer a second
// geometry candidate that can suppress targetSelector.
const slidesSource=fs.readFileSync(path.join(root,'50_slides.js'),'utf8');
assert.match(slidesSource,/if \(def\.no === 40\)[\s\S]*?geometryTargetFromActiveRecord_/);
assert.doesNotMatch(slidesSource,/addGeometryTarget\(def\.no, top, 'heading'\)/);
assert.match(slidesSource,/if \(no40UsesDurableRecord\)[\s\S]*?geometryTargets\[def\.no\] = \[\]/);
console.log('ok generated v2 map: No.40 records the durable-selected heading only');

function normalApplyFixture(mode){
  let geometry={left:1,top:2,width:3,height:4};
  const before={...geometry},calls=[],raw={};
  const selector={type:'IMAGE',title:'hero',description:'',templateGeometry:{...before},role:'primary-image'};
  const resolved={caseName:'A社',hearingId:'hearing',presentationId:'deck',presentationRevision:'rev-1',pages:[{slideNo:1,pageId:'p1',targetSelector:selector}]};
  const deck={id:'deck',slides:[{id:'p1',elements:[raw]}]};
  const runtime={
    resolve:()=>resolved,revision:()=> 'rev-1',open:()=>deck,presentationId:d=>d.id,slides:d=>d.slides,slideId:s=>s.id,elements:s=>s.elements,pageSize:()=>({width:720,height:450}),
    element:()=>({id:'target',type:'IMAGE',title:'hero',description:'',text:'',geometry:{...geometry},contentHash:'content',raw}),
    setGeometry:(_,next)=>{calls.push({...next});if((mode==='readback-mismatch'||mode==='restore-failure')&&calls.length===1)geometry={...next,top:next.top+.5};else if(mode==='restore-failure'&&calls.length===2)geometry={...next,left:next.left+.5};else geometry={...next};}
  };
  const read=ctx.normalGeometryRead_({caseName:'A社',slideNo:1},runtime),element=read.snapshot.pages[0].elements[0];
  const plan={schema:ctx.NORMAL_GEOMETRY_SCHEMA_,status:'approved',approvedBy:'owner',revision:1,caseName:'A社',presentationId:'deck',pageId:'p1',objectId:'target',presentationRevision:read.snapshot.revision,pageOrder:read.snapshot.pageOrder,pageHashes:ctx.normalGeometryPageHashes_(read.snapshot.pages),before:{...before},after:{...before,top:9},contentHash:element.contentHash,persistentSelector:{slideNo:1,selector:element.persistentSelector.selector},untouched:ctx.normalGeometryUntouched_(read.snapshot.pages,'p1','target')};
  plan.planHash=ctx.normalGeometryPlanHash_(plan);
  return {runtime,plan,calls,current:()=>({...geometry}),before};
}
let normal=normalApplyFixture('success');
assert.strictEqual(ctx.normalGeometryApply_(normal.plan,normal.runtime).ok,true);
assert.deepStrictEqual(normal.current(),normal.plan.after);assert.strictEqual(normal.calls.length,1);
normal=normalApplyFixture('readback-mismatch');
assert.throws(()=>ctx.normalGeometryApply_(normal.plan,normal.runtime),/元の配置へ復元し、読戻しを確認しました/);
assert.deepStrictEqual(normal.current(),normal.before);assert.strictEqual(normal.calls.length,2);
normal=normalApplyFixture('restore-failure');
assert.throws(()=>ctx.normalGeometryApply_(normal.plan,normal.runtime),/復元または読戻し確認にも失敗/);
assert.notDeepStrictEqual(normal.current(),normal.before);assert.strictEqual(normal.calls.length,2);
console.log('ok normal geometry: successful apply, readback rollback, and rollback failure are distinguished');

const isolationAuditSource=fs.readFileSync(path.join(root,'58_isolated_case_audit.js'),'utf8');
const fixtureFaqStart=isolationAuditSource.indexOf('function runIsolationFixtureFaqNormalCase()');
const structuredFaqStart=isolationAuditSource.indexOf('function runIsolationStructuredFaqNormalCase()');
assert.ok(fixtureFaqStart>0&&structuredFaqStart>fixtureFaqStart);
assert.strictEqual((isolationAuditSource.match(/runIsolationFixtureFaqNormalCase\(/g)||[]).length,1);
assert.strictEqual((isolationAuditSource.match(/runIsolationStructuredFaqNormalCase\(/g)||[]).length,1);
const fixtureFaqBody=isolationAuditSource.slice(fixtureFaqStart,structuredFaqStart),structuredFaqBody=isolationAuditSource.slice(structuredFaqStart);
for(const body of [fixtureFaqBody,structuredFaqBody]){
  assert.match(body,/withIsolationTestConfig_/);assert.match(body,/隔離通し_生成_(?:rootID|ヒアリングID)/);
  assert.doesNotMatch(body,/(?:notify|MailApp|UrlFetchApp|processNewFlow_|main\s*\()/);
}
assert.match(isolationAuditSource,/Manual fixed-isolation write test/);
console.log('ok FAQ fixture writers: fixed-ID manual-only entrypoints with no normal-flow or notification call');
