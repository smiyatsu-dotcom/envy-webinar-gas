// Q&Aカンペの納品は既存納品関数を通し、同名を増やさないことを検証する。
const assert=require('assert'); const {makeEnv,loadGas}=require('./fake_gas.cjs');
const e=makeEnv(),G=loadGas(e); const out=e.makeFolder('出力',null,'OUT'),del=e.makeFolder('【納品】',null,'1l3qvbW2u-4RLBvNov4ZUj-opEGapdZnJ'),src=out.addFolder('案件'),dst=del.addFolder('案件');
src.addFile('【Q&Aカンペ】案件','doc','application/vnd.google-apps.document');src.addFile('【Q&Aカンペ】案件.pdf','pdf','application/pdf');
G.deliveryFolderOf_=()=>dst; G.caseFolderIfExists_=()=>src; G.syncQaCue_=()=>{}; G.webinarTitleOf_=()=>''; const st={getRange:()=>({setValue(){}})};
G.納品する_({ASSETS_FOLDER_ID:'none'},'案件',st); G.納品する_({ASSETS_FOLDER_ID:'none'},'案件',st);
assert.strictEqual(dst.getFilesByName('【Q&Aカンペ】案件').hasNext(),true);assert.strictEqual(dst.getFilesByName('【Q&Aカンペ】案件.pdf').hasNext(),true);assert.strictEqual(e.log.filter(x=>x[0]==='copy'&&x[1].includes('Q&Aカンペ')).length,2);console.log('ok Q&A納品重複防止');
