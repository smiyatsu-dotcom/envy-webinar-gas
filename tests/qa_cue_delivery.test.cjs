// Q&Aカンペの納品は、同名を増やさず、二度目も実際に成功して最新版を
// delivery record と一致させる。Google へは接続せず fake Drive だけを使う。
const assert = require('assert');
const { makeEnv, loadGas } = require('./fake_gas.cjs');

const DOC = '【Q&Aカンペ】案件';
const PDF = '【Q&Aカンペ】案件.pdf';

function setup() {
  const env = makeEnv(), G = loadGas(env);
  const out = env.makeFolder('出力', null, 'OUT');
  const delivery = env.makeFolder('【納品】', null, '1l3qvbW2u-4RLBvNov4ZUj-opEGapdZnJ');
  const src = out.addFolder('案件'), dest = delivery.addFolder('案件');
  src.addFile(DOC, 'doc', 'application/vnd.google-apps.document');
  src.addFile(PDF, 'pdf', 'application/pdf');
  G.deliveryFolderOf_ = () => dest;
  G.caseFolderIfExists_ = () => src;
  G.webinarTitleOf_ = () => '';
  let message = '';
  const settings = { getRange: () => ({ setValue: value => { message = String(value); } }) };
  return { env, G, dest, settings, message: () => message };
}
function deliveryFiles(env, dest, name) {
  return Object.values(env.files).filter(file => file.parent === dest && file.name === name);
}
function activeDeliveryFile(env, dest, name) {
  const files = deliveryFiles(env, dest, name).filter(file => !file.trashed);
  assert.strictEqual(files.length, 1, name + ' の現行納品物が1件であること');
  return files[0];
}

{
  const { env, G, dest, settings } = setup();
  G.syncQaCue_ = () => {};

  const first = G.納品する_({ ASSETS_FOLDER_ID: 'none' }, '案件', settings);
  assert.strictEqual(first.入れた, true, '1回目はQ&A Doc/PDFを納品する');
  const firstDoc = activeDeliveryFile(env, dest, DOC);
  const firstPdf = activeDeliveryFile(env, dest, PDF);

  const second = G.納品する_({ ASSETS_FOLDER_ID: 'none' }, '案件', settings);
  assert.strictEqual(second.入れた, true, '2回目も既存自動納品物を入れ直して成功する');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(second.残した)), []);
  assert.strictEqual(firstDoc.trashed, true, '前回のDocだけをゴミ箱へ移す');
  assert.strictEqual(firstPdf.trashed, true, '前回のPDFだけをゴミ箱へ移す');
  assert.strictEqual(env.log.filter(entry => entry[0] === 'trash' && entry[1].includes('Q&Aカンペ')).length, 2);
  assert.strictEqual(env.log.filter(entry => entry[0] === 'copy' && entry[1].includes('Q&Aカンペ')).length, 4);

  const currentDoc = activeDeliveryFile(env, dest, DOC);
  const currentPdf = activeDeliveryFile(env, dest, PDF);
  assert.notStrictEqual(currentDoc.getId(), firstDoc.getId());
  assert.notStrictEqual(currentPdf.getId(), firstPdf.getId());
  const docRecord = JSON.parse(env.props['DELIVERED_' + dest.getId() + '_Q&Aカンペ']);
  const pdfRecord = JSON.parse(env.props['DELIVERED_' + dest.getId() + '_Q&AカンペPDF']);
  assert.strictEqual(docRecord.id, currentDoc.getId(), 'Docの現行リンク先を納品記録が指す');
  assert.strictEqual(pdfRecord.id, currentPdf.getId(), 'PDFの現行リンク先を納品記録が指す');
  assert.strictEqual(currentDoc.getUrl(), 'https://drive/' + docRecord.id);
  assert.strictEqual(currentPdf.getUrl(), 'https://drive/' + pdfRecord.id);
}

{
  const { env, G, dest, settings, message } = setup();
  G.syncQaCue_ = () => { throw new Error('Q&Aカンペを生成・納品できません: FAQ 1の質問が未記入です'); };
  const result = G.納品する_({ ASSETS_FOLDER_ID: 'none' }, '案件', settings);
  assert.strictEqual(result.入れた, false, '未完成FAQでは納品しない');
  assert.match(message(), /何も入れていません/);
  assert.strictEqual(env.log.filter(entry => entry[0] === 'copy').length, 0, '未完成FAQでは既存Q&Aを含めコピーしない');
  assert.strictEqual(deliveryFiles(env, dest, DOC).length, 0);
  assert.strictEqual(deliveryFiles(env, dest, PDF).length, 0);
}

console.log('ok Q&A納品: 二回成功・現行リンク・旧版trash・未完成FAQ停止');
