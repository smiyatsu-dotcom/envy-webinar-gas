// Q&A（No.104.5）の単体テスト。GASを使わず node だけで動かす。
//   node tests/qa.test.cjs
// ★拡張子を .cjs にしてあるのは、clasp push の対象（.js/.gs/.json）に入れないため。
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const root = path.join(__dirname, '..');

// ---- 何を呼ばれても壊れない「なんでも」オブジェクト（GASの API の代わり） ----
function any() {
  const f = function () { return any(); };
  return new Proxy(f, {
    get(t, k) {
      if (k === 'hasNext') return () => false;
      if (k === Symbol.toPrimitive) return () => '';
      if (k === 'length') return 0;
      if (k === 'forEach' || k === 'map' || k === 'filter') return () => [];
      return any();
    },
    apply() { return any(); }
  });
}

function load(extra) {
  const ctx = Object.assign({
    console, Logger: { log() {} },
    PropertiesService: any(), DriveApp: any(), SpreadsheetApp: any(), SlidesApp: any(),
    DocumentApp: any(), MimeType: any(), LockService: any(), ScriptApp: any()
  }, extra || {});
  vm.createContext(ctx);
  fs.readdirSync(root).filter(f => /\.js$/.test(f)).sort().forEach(f => {
    // const は vm の外から見えないので var に寄せる（中身は変えない）
    const src = fs.readFileSync(path.join(root, f), 'utf8').replace(/^const /gm, 'var ');
    vm.runInContext(src, ctx, { filename: f });
  });
  return ctx;
}

let pass = 0;
function test(name, fn) {
  try { fn(); pass++; console.log('ok  ' + name); }
  catch (e) { console.log('NG  ' + name + '\n    ' + (e && e.stack || e)); process.exitCode = 1; }
}

const G = load();

// ============================================================
// 1. 構成（行定義）: 既存の番号と並びが1つも変わっていない
// ============================================================
test('行定義: 1〜106はそのまま・104.5が104と105の間に1行だけ', () => {
  const nos = G.rowDefs_().map(r => r.no);
  const ints = nos.filter(n => n !== 104.5);
  assert.deepStrictEqual(Array.from(ints), Array.from({ length: 106 }, (_, i) => i + 1));
  assert.strictEqual(nos.length, 107);
  assert.strictEqual(nos.indexOf(104.5), nos.indexOf(104) + 1);
  assert.strictEqual(nos.indexOf(105), nos.indexOf(104.5) + 1);
  const def = G.rowDefs_().find(r => r.no === 104.5);
  assert.strictEqual(def.tab, 'オファー');
  assert.strictEqual(def.km, 'GEN'); assert.strictEqual(def.dm, 'GEN');
});

test('AIの生成対象: 台本・構成テキストとも 104.5 が入る（104の後ろ）', () => {
  const gen = G.daihonGenRows_();
  assert.ok(gen.partB.indexOf(104.5) === gen.partB.indexOf(104) + 1);
  const k = G.kouseiGenRows_();
  assert.ok(k.indexOf(104.5) === k.indexOf(104) + 1);
  assert.ok(G.kouseiSpec_()[104.5]);
});

test('SKIP_ROWS: 104.5 は数値でも引ける', () => {
  assert.ok(G.SKIP_ROWS[104.5]);
  assert.ok(G.isSkipRow_({ no: 104.5 }, { kousei: '' }));
  assert.ok(G.isSkipRow_({ no: 104.5 }, { kousei: '（斜線）' }));
  assert.ok(!G.isSkipRow_({ no: 104.5 }, { kousei: 'Q. 何？\nA. こう' }));
});

// ============================================================
// 2. 取り込み: 104.5 が無くても止まらない（空で補完）
// ============================================================
test('取り込み: AIが104.5を出さなくても欠落エラーにならない', () => {
  const nos = G.kouseiGenRows_();
  const rows = nos.filter(n => n !== 104.5 && n !== 99 && n !== 104).map(n => ({ no: n, kousei_text: 'x 図解（' + 'a' + '）' }));
  const w = [];
  const m = G.validateRows_(rows, nos, 'kousei_text', w);
  assert.strictEqual(m[104.5], '');
});

test('取り込み: 104.5 以外の欠落は今までどおり止まる', () => {
  const nos = G.kouseiGenRows_();
  const rows = nos.filter(n => n !== 50).map(n => ({ no: n, kousei_text: 'x' }));
  assert.throws(() => G.validateRows_(rows, nos, 'kousei_text', []), /No\.50/);
});

test('取り込み: "no": "104.5"（文字列）でも拾える', () => {
  const m = G.validateRows_([{ no: '104.5', daihon: 'よくいただくご質問…' }], [104.5], 'daihon', []);
  assert.strictEqual(m[104.5], 'よくいただくご質問…');
});

// ============================================================
// 3. ヒアリングシートの「⒌ よくある質問（Q&A）」の読み取り（原本の形 A60:D66 に合わせる）
// ============================================================
function offerValues(qa) {
  const v = [];
  for (let i = 0; i < 70; i++) v.push(['', '', '', '']);
  v[53] = ['', '4. 基本情報', '', ''];
  v[54] = ['', '', '質問1：', '登壇者情報（プロフィール文・画像）'];
  v[55] = ['', '', '記入欄：', '山田太郎。…'];
  v[59] = ['', '⒌ よくある質問（Q&A）', '', ''];
  for (let k = 0; k < 3; k++) {
    v[60 + k * 2] = ['', '', 'よくある質問：', (qa[k] && qa[k][0]) || ''];
    v[61 + k * 2] = ['', '', '記入欄：', (qa[k] && qa[k][1]) || ''];
  }
  return v;
}

test('Q&A読み取り: 3組とも空 → []', () => {
  assert.deepStrictEqual(JSON.parse(JSON.stringify(G.qaPairsFromValues_(offerValues([])))), []);
});
test('Q&A読み取り: 記入がある組だけ（2組目が空）', () => {
  const r = G.qaPairsFromValues_(offerValues([['費用は？', '月5万円です'], ['', ''], ['期間は？', '3か月です']]));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r)), [{ q: '費用は？', a: '月5万円です' }, { q: '期間は？', a: '3か月です' }]);
});
test('Q&A読み取り: 片方だけの組は使わない', () => {
  const r = G.qaPairsFromValues_(offerValues([['質問だけ', ''], ['', '答えだけ'], ['Q3', 'A3']]));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r)), [{ q: 'Q3', a: 'A3' }]);
});
test('Q&A読み取り: 4. 基本情報の「記入欄：」を拾わない／見出しが無いシートは []', () => {
  const v = offerValues([]); v[59] = ['', '', '', ''];
  assert.strictEqual(G.qaPairsFromValues_(v).length, 0);
});
test('Q&A読み取り: 行が足されてずれていても見出しから読む', () => {
  const v = offerValues([['費用は？', '月5万円です']]);
  v.splice(10, 0, ['', '', '記入欄：', '別の欄'], ['', '', '記入欄：', '別の欄']);
  assert.strictEqual(G.qaPairsFromValues_(v)[0].q, '費用は？');
});

test('プロンプト: 記入なし → 空文字を出す指示／記入あり → 組がそのまま入る', () => {
  assert.match(G.qaPromptBlock_([]), /空文字/);
  assert.match(G.qaPromptBlock_(undefined), /空文字/);
  const b = G.qaPromptBlock_([{ q: '費用は？', a: '月5万円です' }]);
  assert.match(b, /Q1\. 費用は？\nA1\. 月5万円です/);
});
test('プロンプト全体: 104.5 の台本ルールとQ&Aの組が入る', () => {
  const p = G.buildManualPrompt_({ offer: 'o', problem: 'p', story: 's', qa: [{ q: '費用は？', a: '月5万円です' }] }, { iconP: [], iconW: [], bg85: [], bg25: [], bg15: [], select: [] }, 'サンプル社_ヒアリングシート');
  assert.match(p, /No\.104\.5: よくある質問/);
  assert.match(p, /A1\. 月5万円です/);
  assert.match(p, /生成対象行: [^\n]*104,104\.5,105|生成対象行: [^\n]*104,104\.5/);
});

// ============================================================
// 4. 取り込み時の安全装置: ヒアリングにQ&Aが無いのにAIが書いてきたら捨てる
// ============================================================
test('取り込み安全装置: ヒアリング空 → 104.5 を空にする／記入あり → そのまま', () => {
  const ss = { getSheetByName: () => ({ getRange: () => ({ getValue: () => 'HID' }) }) };
  const G2 = load();
  G2.readHearingQa_ = () => [];
  const d = { 104.5: '作り話' }, k = { 104.5: 'Q. 作り話' }, w = [];
  G2.dropQaIfHearingEmpty_(ss, d, k, w);
  assert.strictEqual(d[104.5], ''); assert.strictEqual(k[104.5], '');
  G2.readHearingQa_ = () => [{ q: 'a', a: 'b' }];
  const d2 = { 104.5: 'ok' }, k2 = { 104.5: 'Q. ok' };
  G2.dropQaIfHearingEmpty_(ss, d2, k2, []);
  assert.strictEqual(d2[104.5], 'ok');
});

// ============================================================
// 5. スライド番号の計算: 1〜106 は今までの式と完全に同じ
// ============================================================
test('スライド番号: 1〜106は旧式と同じ・104.5は -1（テンプレに無い）', () => {
  for (let no = 1; no <= 106; no++) {
    const old = no <= 96 ? no - 1 : (no <= 102 ? no : no + 1);
    assert.strictEqual(G.slideIndexFor_(no), old, 'No.' + no);
  }
  assert.strictEqual(G.slideIndexFor_(104.5), -1);
});

// ============================================================
// 6. Q&Aスライドの中身
// ============================================================
test('構成テキスト→Q&Aの組（Q./A.・全角・番号付き・続き行）', () => {
  const r = G.qaPairsFromKousei_('Q. 費用は？\nA. 月5万円です\nＱ２．期間は？\nＡ２：3か月です\n延長もできます');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r)), [{ q: '費用は？', a: '月5万円です' }, { q: '期間は？', a: '3か月です\n延長もできます' }]);
  assert.strictEqual(G.qaPairsFromKousei_('Q.1\nA.1\nQ.2\nA.2\nQ.3\nA.3\nQ.4\nA.4').length, 3);
  assert.strictEqual(G.qaPairsFromKousei_('（斜線）').length, 1); // ※斜線は isSkipRow_ で先に弾く
});
test('文字の大きさ: 1組>2組>3組／長い答えは縮めて収める', () => {
  const p = (n, len) => Array.from({ length: n }, () => ({ q: 'あ'.repeat(20), a: 'い'.repeat(len) }));
  const a1 = G.qaLayout_(p(1, 30), 562, 319), a2 = G.qaLayout_(p(2, 30), 562, 319), a3 = G.qaLayout_(p(3, 30), 562, 319);
  assert.ok(a1.q > a2.q && a2.q > a3.q, JSON.stringify([a1, a2, a3]));
  const big = G.qaLayout_(p(3, 200), 562, 319);
  assert.ok(big.a < a3.a && big.a >= 12, JSON.stringify(big));
});

// ============================================================
// 7. 通し（generateSlides_）をニセのスライドで: 既存案件で割り込まない・新案件は104の直後に1枚
// ============================================================
function fakeSlide(i, log) {
  const s = {
    i, notes: null, skipped: false,
    getShapes: () => [], getImages: () => [], replaceAllText() {},
    getNotesPage: () => ({ getSpeakerNotesShape: () => ({ getText: () => ({ setText: t => { s.notes = t; } }) }) }),
    setSkipped: v => { s.skipped = v; },
    remove: () => { log.push('remove:' + i); },
    duplicate: () => { log.push('duplicate:' + i); const d = fakeSlide('dup' + i, log); d.isDup = true; log.dupSlide = d; return d; },
    insertTextBox: () => any()
  };
  return s;
}
function runGenerate(rows, ryoukin) {
  const log = [];
  const slides = Array.from({ length: 108 }, (_, i) => fakeSlide(i, log));
  const pres = { getSlides: () => slides, saveAndClose() {}, getPageWidth: () => 720, getPageHeight: () => 450 };
  let daihonTxt = null;
  const G3 = load({
    SlidesApp: new Proxy({}, { get: (t, k) => k === 'openById' ? () => pres : any() }),
    DriveApp: new Proxy({}, { get: (t, k) => k === 'createFile' ? (n, txt) => { daihonTxt = txt; return any(); } : any() })
  });
  G3.getConfig_ = () => ({ TEMPLATE_SLIDE_ID: 'T', OUTPUT_FOLDER_ID: 'O', ASSETS_FOLDER_ID: 'A' });
  G3.readKouseiSheet_ = () => ({ name: 'サンプル社', assetsFolderId: '', ryoukin: ryoukin || '単一', rows });
  G3.fillReal_ = () => {}; G3.replaceImages_ = () => {};
  let qaFilled = null;
  G3.fillQa_ = (slide, row) => { qaFilled = { slide, row }; };
  const res = G3.generateSlides_('K');
  return { log, slides, res, qaFilled, daihonTxt };
}
function baseRows(withQa, qaKousei, qaDaihon) {
  const rows = {};
  G.rowDefs_().forEach(d => {
    if (d.no === 104.5 && !withQa) return;   // 古い構成シートには104.5の行が無い
    rows[d.no] = { kousei: 'k' + d.no, daihon: 'd' + d.no, kouseiResolved: 'k' + d.no, icon: '', bg: '' };
  });
  if (withQa) rows[104.5] = { kousei: qaKousei, daihon: qaDaihon, kouseiResolved: qaKousei };
  return rows;
}

test('通し: 古い構成シート（104.5の行なし・ホスピタソン社の形）→ 複製しない・警告も出さない', () => {
  const r = runGenerate(baseRows(false));
  assert.ok(!r.log.some(x => /^duplicate/.test(x)), r.log.join(','));
  assert.strictEqual(r.qaFilled, null);
  assert.ok(!r.res.warnings.some(w => /104\.5/.test(w)), r.res.warnings.join('\n'));
  assert.strictEqual(r.slides[105].notes, 'd104');          // No.104 は今までどおり
  assert.strictEqual(r.slides[106].notes, 'd105');
  assert.ok(!/Q&A|よくある/.test(r.daihonTxt));
  assert.deepStrictEqual(r.log.filter(x => /remove/.test(x)), ['remove:103', 'remove:96']); // 料金バリアント削除も同じ
});

test('通し: 新しい構成シートでQ&Aが斜線 → 複製しない・台本txtにも入らない', () => {
  const r = runGenerate(baseRows(true, '（斜線）', '（斜線: このスライドはテンプレのままスキップ設定で出力されます）'));
  assert.ok(!r.log.some(x => /^duplicate/.test(x)));
  assert.ok(!/斜線/.test(r.daihonTxt));
});

test('通し: 新しい構成シートでQ&Aあり → No.104のテンプレ（108枚の106枚目）を複製・台本は104と105の間', () => {
  const r = runGenerate(baseRows(true, 'Q. 費用は？\nA. 月5万円です', 'ここで、よくいただくご質問にお答えします。'));
  assert.deepStrictEqual(r.log.filter(x => /^duplicate/.test(x)), ['duplicate:105']);
  assert.ok(r.qaFilled && r.qaFilled.slide.isDup);
  assert.strictEqual(r.log.dupSlide.notes, 'ここで、よくいただくご質問にお答えします。');
  assert.strictEqual(r.slides[105].notes, 'd104');           // 元のNo.104はそのまま
  assert.strictEqual(r.slides[106].notes, 'd105');           // No.105以降もずれない
  const t = r.daihonTxt;
  assert.ok(t.indexOf('d104') < t.indexOf('よくいただく') && t.indexOf('よくいただく') < t.indexOf('d105'), t.slice(-300));
});

test('通し: No.104が斜線（分割なし）でもQ&Aは出る・Q&Aは斜線スキップを引き継がない', () => {
  const rows = baseRows(true, 'Q. 費用は？\nA. 月5万円です', 'Q&Aの台本');
  rows[104] = { kousei: '（斜線）', daihon: '（斜線）', kouseiResolved: '（斜線）' };
  const r = runGenerate(rows);
  assert.strictEqual(r.slides[105].skipped, true);
  assert.strictEqual(r.log.dupSlide.skipped, false);
  assert.strictEqual(r.log.dupSlide.notes, 'Q&Aの台本');
});

test('通し: 松竹梅でも同じ（No.104の位置は料金バリアントの影響を受けない）', () => {
  const r = runGenerate(baseRows(true, 'Q. a\nA. b', 'qa'), '松竹梅');
  assert.deepStrictEqual(r.log.filter(x => /^duplicate/.test(x)), ['duplicate:105']);
  assert.deepStrictEqual(r.log.filter(x => /remove/.test(x)), ['remove:102', 'remove:95']);
});

// ============================================================
// 8. 構成シートへの書き込み・読み戻し
// ============================================================
function fakeSheet(values) {
  const sh = {
    values: values || null,
    getRange: (r, c, nr, nc) => ({
      setValues: v => { sh.values = v; return any(); },
      setFontWeight: () => any(), setWrap: () => any()
    }),
    getDataRange: () => ({ getValues: () => sh.values }),
    setColumnWidth() { return sh; }, setFrozenRows() {}
  };
  return sh;
}
test('構成シート: オファータブで104の次の行が104.5・取り込み済みで空なら（斜線）', () => {
  const tabs = {};
  const ss = { getSheetByName: n => tabs[n], insertSheet: n => (tabs[n] = fakeSheet()) };
  G.writeContentTabs_(ss, { 104: 'k104' }, { 104: 'd104' }, {}, {});
  const v = tabs['オファー'].values;
  const i = v.findIndex(r => r[0] === 104);
  assert.strictEqual(v[i + 1][0], 104.5);
  assert.strictEqual(v[i + 1][3], '（斜線）');
  assert.strictEqual(v[i + 2][0], 105);
});
test('構成シート読み戻し: 104.5 を数値で読める／古い形（行なし）は rows[104.5] が無い', () => {
  function sheetWith(offerRows) {
    const tab = n => fakeSheet(n === 'オファー' ? offerRows : [['No.']]);
    const settings = { getRange: () => ({ getValue: () => 'サンプル社' }) };
    return { getSheetByName: n => n === '設定' ? settings : tab(n) };
  }
  const head = ['No.', '基本構成', '小分類', '構成テキスト', '台本', 'アイコン', '背景'];
  const G4 = load({ SpreadsheetApp: { openById: () => sheetWith([head, [104, '分割', '', 'k', 'd', '', ''], [104.5, 'Q&A', '', 'Q. a\nA. b', 'qa', '', ''], [105, 'BYAF', '', '', '', '', '']]) } });
  const d = G4.readKouseiSheet_('x');
  assert.strictEqual(d.rows[104.5].daihon, 'qa');
  assert.strictEqual(d.rows[105].kousei, '');
  const G5 = load({ SpreadsheetApp: { openById: () => sheetWith([head, [104, '分割', '', 'k', 'd', '', ''], [105, 'BYAF', '', '', '', '', '']]) } });
  assert.strictEqual(G5.readKouseiSheet_('x').rows[104.5], undefined);
});

console.log('\n' + pass + '件 通過' + (process.exitCode ? '（失敗あり）' : ''));
