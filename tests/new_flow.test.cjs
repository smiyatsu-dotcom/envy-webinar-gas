// 資料生成の新しい流れ（管理シート H列＝新）の単体テスト。GASを使わず node だけで動かす。
//   node tests/new_flow.test.cjs
// ★本物のGoogleにはつながない。実在の案件（ホスピタソン社など）の名前・IDは使わない。すべて偽のデータ。
// ★いちばん大事なのは「旧の流れが1つも変わらない」こと。変更前のコミット（BASE）のコードと、いまのコードに
//   同じ偽の管理シート・構成シート・ドライブを渡して main() を回し、書き込みの記録が1つ残らず同じかを比べる。
const assert = require('assert');
const { makeEnv, loadGas } = require('./fake_gas.cjs');

const BASE = '37c44db';   // 新しい流れを入れる前（Q&Aまで入ったもの）

let pass = 0;
function test(name, fn) {
  try { fn(); pass++; console.log('ok  ' + name); }
  catch (e) { console.log('NG  ' + name + '\n    ' + (e && e.stack || e)); process.exitCode = 1; }
}

const G = loadGas(makeEnv());

// ============================================================
// 1. 新の流れの見分け・【開始】の読み方
// ============================================================
test('新の流れか_: H列が「新」の行だけ。空・短い行・ほかの字は旧', () => {
  assert.strictEqual(G.新の流れか_(['id', 'A社', '完了', 'u', '', '', '', '新']), true);
  assert.strictEqual(G.新の流れか_(['id', 'A社', '完了', 'u', '', '', '', ' 新 ']), true);
  assert.strictEqual(G.新の流れか_(['id', 'A社', '完了', 'u', '', '', '', '']), false);
  assert.strictEqual(G.新の流れか_(['id', 'A社', '完了', 'u', '', '', '']), false);   // H列が無い（旧の管理シート）
  assert.strictEqual(G.新の流れか_(['id', 'A社', '完了', 'u', '', '', '', '旧']), false);
  assert.strictEqual(G.新の流れか_(null), false);
});

test('開始を読む_: つなぎの形を読み、足りなければ日本語で止める', () => {
  const ok = G.開始を読む_(JSON.stringify({ 案件名: '動画撮影テスト社', ヒアリングシートID: '1AbCdEfGhIjKlMnOpQrStUvWxYz012345', 台帳の行: 12, 置いた日時: '2026-09-26T10:00:00+09:00' }));
  assert.strictEqual(ok.案件名, '動画撮影テスト社');
  assert.strictEqual(ok.ヒアリングシートID, '1AbCdEfGhIjKlMnOpQrStUvWxYz012345');
  assert.strictEqual(ok.台帳の行, 12);
  // URLで入っていてもIDを取り出す
  assert.strictEqual(G.開始を読む_(JSON.stringify({ 案件名: 'X社', ヒアリングシートID: 'https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz012345/edit' })).ヒアリングシートID, '1AbCdEfGhIjKlMnOpQrStUvWxYz012345');
  assert.throws(() => G.開始を読む_('{'), /JSONとして読めません/);
  assert.throws(() => G.開始を読む_('{"ヒアリングシートID":"1AbCdEfGhIjKlMnOpQrStUvWxYz012345"}'), /案件名/);
  assert.throws(() => G.開始を読む_('{"案件名":"X社","ヒアリングシートID":"short"}'), /ヒアリングシートID/);
});

test('新の流れの案件名_: 素材フォルダ名（◯◯社_ヒアリングシート）にそろえる・二重に付けない', () => {
  assert.strictEqual(G.新の流れの案件名_('動画撮影テスト社'), '動画撮影テスト社_ヒアリングシート');
  assert.strictEqual(G.新の流れの案件名_('X社_ヒアリングシート'), 'X社_ヒアリングシート');
  assert.strictEqual(G.caseCompany_(G.新の流れの案件名_('動画撮影テスト社')), '動画撮影テスト社');
});

test('開始ファイルか_: 【開始】◯◯社.json だけ（【済】開始_・txt は拾わない。分解された濁点も拾う）', () => {
  assert.ok(G.開始ファイルか_('【開始】動画撮影テスト社.json'));
  assert.ok(!G.開始ファイルか_('【済】開始_動画撮影テスト社.json'));
  assert.ok(!G.開始ファイルか_('【開始】動画撮影テスト社.txt'));
  assert.ok(G.開始ファイルか_('【開始】グループ社.json'.normalize('NFD')));
});

test('テスト案件か_: 定型連絡.gs 共通_テスト案件か_ と同じ（名前に「テスト」「リハーサル」・決まった名前）', () => {
  assert.ok(G.テスト案件か_('動画撮影テスト社_ヒアリングシート'));
  assert.ok(G.テスト案件か_('リハーサル'));
  assert.ok(G.テスト案件か_('4章_畠'));
  assert.ok(!G.テスト案件か_('過去A社_ヒアリングシート'));
  assert.ok(!G.テスト案件か_(''));
});

// ============================================================
// 2. ファイルの仕分け（つなぎの名前）
// ============================================================
function 偽ファイル(名, 中身, mime) {
  return { 名, mime: mime || 'text/plain', url: 'https://drive/' + 名, 日時: '2026-09-26T10:00:00+09:00', 読む: () => (中身 === undefined ? null : 中身) };
}

test('ファイルを仕分ける_: つなぎのファイルを見分ける（濁点が分解されていても）', () => {
  const 観 = G.ファイルを仕分ける_('動画撮影テスト社', [
    偽ファイル('【タイトル案】動画撮影テスト社.txt'),
    偽ファイル('【LP】動画撮影テスト社', undefined, 'application/vnd.google-apps.document'),
    偽ファイル('【保留】サムネ指示_動画撮影テスト社.txt'),
    偽ファイル('【指示】資料_動画撮影テスト社.txt'),
    偽ファイル('【作業中】動画撮影テスト社_図解.json', { 仕事: '図解', 始: '2026-09-26T10:00:00+09:00', 見込み分: 25, 心拍: '2026-09-26T10:05:00+09:00', 枚: '3/15' }),
    偽ファイル('【作業中】動画撮影テスト社_壊れ.json', null),   // 書きかけは飛ばす
    偽ファイル('【指示】直し_動画撮影テスト社_a1.json'.normalize('NFD'), { ID: 'a1', 対象: 'スライド', 依頼: '12枚目' }),
    偽ファイル('【済】直し_動画撮影テスト社_a0.json', { ID: 'a0', 結果: '直した' }),
    偽ファイル('【承認】動画撮影テスト社.json', { 人: '福田真里（PM）' }),
    偽ファイル('【済】承認_動画撮影テスト社.json', { 人: '前の承認' }),
    偽ファイル('【指示】LP_動画撮影テスト社.txt'),
    偽ファイル('【状態】進み.json', { 流れ: '新' })
  ]);
  assert.ok(観.ファイル.タイトル案 && 観.ファイル.LP);
  assert.strictEqual(観.ファイル.保留サムネ, true);
  assert.strictEqual(観.ファイル.指示資料, true);
  assert.strictEqual(観.ファイル.指示LP, true);
  assert.strictEqual(観.作業中.length, 1);
  assert.strictEqual(観.直し指示[0].ID, 'a1');
  assert.strictEqual(観.直し済[0].ID, 'a0');
  assert.strictEqual(観.承認ファイル.中身.人, '福田真里（PM）');
  assert.strictEqual(観.納品ファイル, null);
});

// ============================================================
// 3. 進みの計算
// ============================================================
const いま = new Date('2026-09-26T10:30:00+09:00');
function 観の元(上書き) {
  const F = { タイトル案: null, LP: null, LP待ち: false, 指示LP: false, 指示資料: false, 指示サムネ: false, 保留サムネ: false, 要確認サムネ: false, 台本txt: null, 音声: null, サムネイル: null, 図解済み: null };
  const o = { ファイル: Object.assign(F, (上書き && 上書き.ファイル) || {}), 作業中: [], 直し指示: [], 直し済: [], 案件名: '動画撮影テスト社', 状態: 'AI出力待ち', メッセージ: '', スライドURL: '' };
  Object.keys(上書き || {}).forEach(k => { if (k !== 'ファイル') o[k] = 上書き[k]; });
  return o;
}
const 済の印 = { url: 'https://x', 日時: '2026-09-26T10:10:00+09:00' };
function 全部できた観() {
  return 観の元({ 状態: '完了', スライドURL: 'https://slide', ファイル: { タイトル案: 済の印, LP: 済の印, サムネイル: 済の印, 台本txt: 済の印, 音声: 済の印, 図解済み: { 枚数: 15, 日時: '2026-09-26T10:20:00+09:00', 止まった: [] } } });
}
const 名と状態 = (進) => 進.系統.LP.concat(進.系統.資料).map(x => x.名 + ':' + x.状態).join(' ');

test('進み: 始めた直後（LP・資料の指示あり）は「一気に作っています」、押せるものなし', () => {
  const 進 = G.進みを計算_(観の元({ ファイル: { 指示LP: true, 指示資料: true } }), G.進み_記録の既定_(), いま);
  assert.strictEqual(進.流れ, '新');
  assert.strictEqual(名と状態(進), 'タイトル:作成中 サムネイル:待ち LP:作成中 構成・台本:作成中 図解・挿絵:待ち スライド:待ち 音声:待ち');
  assert.ok(/一気に作っています（7つ中 0つできた）/.test(進.いまここ.文));
  // あと分: 資料の系統（30+25+5+15=75）のほうが長い
  assert.strictEqual(進.いまここ.あと分, 75);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(進.押せる)), { 承認: false, 納品: false });
  assert.strictEqual(進.更新, '2026-09-26T10:30:00+09:00');
});

test('進み: 【作業中】の見込み分から、あと何分を出す', () => {
  const 観 = 観の元({ 状態: 'レビュー待ち', 作業中: [{ 仕事: '図解', 始: '2026-09-26T10:10:00+09:00', 見込み分: 25, 心拍: '2026-09-26T10:28:00+09:00', 枚: '12/15' }] });
  const 進 = G.進みを計算_(観, G.進み_記録の既定_(), いま);
  const 図 = 進.系統.資料.find(x => x.名 === '図解・挿絵');
  assert.strictEqual(図.状態, '作成中');
  assert.strictEqual(図.あと分, 5);   // 25 − 20分たった
  assert.strictEqual(図.枚, '12/15');
  assert.strictEqual(進.系統.資料[0].状態, '済');   // 取り込み済み（レビュー待ち）
});

test('進み: 【作業中】の心拍が20分より古いと「止まった」、いまここも止まった', () => {
  const 観 = 観の元({ 状態: 'レビュー待ち', 作業中: [{ 仕事: '図解', 始: '2026-09-26T09:00:00+09:00', 見込み分: 25, 心拍: '2026-09-26T10:00:00+09:00' }] });
  const 進 = G.進みを計算_(観, G.進み_記録の既定_(), いま);
  const 図 = 進.系統.資料.find(x => x.名 === '図解・挿絵');
  assert.strictEqual(図.状態, '止まった');
  assert.ok(/30分途絶えています/.test(図.文));
  assert.ok(/止まっています/.test(進.いまここ.文));
  assert.strictEqual(進.押せる.承認, false);
});

test('進み: 取り込みに失敗・サムネの参考画像なし・スライドのエラーは「止まった」', () => {
  let 進 = G.進みを計算_(観の元({ 状態: 'AI出力待ち', メッセージ: 'txt自動取り込み失敗: 台本JSONがありません' }), G.進み_記録の既定_(), いま);
  assert.strictEqual(進.系統.資料[0].状態, '止まった');
  進 = G.進みを計算_(観の元({ ファイル: { 要確認サムネ: true } }), G.進み_記録の既定_(), いま);
  assert.strictEqual(進.系統.LP[1].状態, '止まった');
  const 記 = G.進み_記録の既定_(); 記.取り込み済み = true;
  進 = G.進みを計算_(観の元({ 状態: 'エラー', メッセージ: 'Exception: xx' }), 記, いま);
  assert.strictEqual(進.系統.資料.find(x => x.名 === 'スライド').状態, '止まった');
});

test('進み: 全部できたら「確認して承認」が押せる。承認のあとは納品だけ押せる', () => {
  const 観 = 全部できた観();
  const 記 = G.進み_記録の既定_();
  let 進 = G.進みを計算_(観, 記, いま);
  assert.strictEqual(進.全部できた, true);
  assert.strictEqual(進.いまここ.文, '全部できました');
  assert.strictEqual(進.押せる.承認, true);
  assert.strictEqual(進.押せる.納品, false);
  assert.strictEqual(進.系統.資料.find(x => x.名 === 'スライド').url, 'https://slide');
  記.承認 = { 済: true, 人: '福田真里（PM）', 日時: '2026-09-26T10:40:00+09:00', revision: 進.生成.revision };
  進 = G.進みを計算_(観, 記, いま);
  assert.strictEqual(進.押せる.承認, false);
  assert.strictEqual(進.押せる.納品, true);
  記.納品 = { 済: true, 日時: '…', revision: 進.生成.revision };
  assert.strictEqual(G.進みを計算_(観, 記, いま).いまここ.文, '納品しました');
});

test('進み: 図解の「止まった」一覧は、済の行の注意として出す（スライドは組める）', () => {
  const 観 = 全部できた観();
  観.ファイル.図解済み.止まった = ['工程3'];
  const 図 = G.進みを計算_(観, G.進み_記録の既定_(), いま).系統.資料.find(x => x.名 === '図解・挿絵');
  assert.strictEqual(図.状態, '済');
  assert.ok(/工程3/.test(図.文));
  assert.strictEqual(図.枚, '15枚');
  // ★作れなかった図解が残っているうちは承認を押せない。いまここに理由が出る
  const 進 = G.進みを計算_(観, G.進み_記録の既定_(), いま);
  assert.strictEqual(進.押せる.承認, false); assert.strictEqual(進.絵待ち, true);
  assert.ok(/工程3/.test(進.いまここ.文) && /ChatGPT/.test(進.いまここ.次), 進.いまここ.文);
  // 制作担当が手で上書きしたら（図解済みより新しい）承認できる
  観.ファイル.画像の日時 = { '工程3': Date.parse('2026-09-26T11:00:00+09:00') };
  assert.strictEqual(G.進みを計算_(観, G.進み_記録の既定_(), いま).押せる.承認, true);
});

test('進み: 図解・挿絵を2回作れず【図解止まった】があれば「止まった」（待ちのままにしない）', () => {
  const 観 = 全部できた観();
  観.状態 = 'AI出力待ち'; 観.ファイル.図解済み = null;
  観.ファイル.図解止まった = { 日時: '2026-09-26T10:30:00+09:00', 文: '2回作り直しても作れませんでした。Discordの現場相談へ' };
  const 進 = G.進みを計算_(観, G.進み_記録の既定_(), いま);
  const 図 = 進.系統.資料.find(x => x.名 === '図解・挿絵');
  assert.strictEqual(図.状態, '止まった'); assert.ok(/現場相談/.test(図.文));
  assert.ok(/図解・挿絵/.test(進.いまここ.文));
});

test('進み: 直しを受けたら該当する行が「作り直し中」、承認は押せない', () => {
  const 記 = G.進み_記録の既定_();
  記.直し = G.直しをまとめる_([], [{ ID: 'a1', 対象: 'スライド', 依頼: '12枚目の数字', 頼んだ人: '福田真里', 元: 'PM相談室', 受付: '2026-09-26T10:15:00+09:00' }], [], '2026-09-26T10:16:00+09:00');
  const 進 = G.進みを計算_(全部できた観(), 記, いま);
  const s = 進.系統.資料.find(x => x.名 === 'スライド');
  assert.strictEqual(s.状態, '作り直し中');
  assert.strictEqual(s.url, 'https://slide');
  assert.ok(/直しています（スライド）/.test(進.いまここ.文));
  assert.strictEqual(進.押せる.承認, false);
  assert.strictEqual(進.直し[0].状態, '作り直し中');
  // 台本の直しは「構成・台本」の行
  記.直し = G.直しをまとめる_([], [{ ID: 'b1', 対象: '台本' }], [], '');
  assert.strictEqual(G.進みを計算_(全部できた観(), 記, いま).系統.資料[0].状態, '作り直し中');
});

test('直しをまとめる_・直しの差分_: 受付 → 済（ボットが【済】を消しても記録に残る）', () => {
  const 一 = G.直しをまとめる_([], [{ ID: 'a1', 対象: 'LP', 依頼: 'x' }], [], 'T1');
  let 差 = G.直しの差分_([], 一);
  assert.strictEqual(差.受けた.length, 1); assert.strictEqual(差.終わった.length, 0);
  const 二 = G.直しをまとめる_(一, [], [{ ID: 'a1', 結果: '直した', 文: 'ok', 後: 'u2' }], 'T2');
  assert.strictEqual(二[0].状態, '済'); assert.strictEqual(二[0].後, 'u2');
  差 = G.直しの差分_(一, 二);
  assert.strictEqual(差.受けた.length, 0); assert.strictEqual(差.終わった.length, 1);
  const 三 = G.直しをまとめる_(二, [], [], 'T3');   // 【済】が消えた
  assert.strictEqual(三[0].状態, '済');
  assert.strictEqual(G.直しの差分_(二, 三).終わった.length, 0);
  const 四 = G.直しをまとめる_(三, [], [{ ID: 'a2', 結果: '直せない', 文: '元の画像が無い' }], 'T4');
  assert.strictEqual(四[1].状態, '止まった');
  差 = G.直しの差分_(三, 四);
  assert.strictEqual(差.終わった.length, 1);
});

test('直しをまとめる_: 図解・タイトルでも、台本を合わせた（直すセル）ときは構成シートへ書くまで作り直し中', () => {
  const 一 = G.直しをまとめる_([], [{ ID: 'z1', 対象: '図解', 依頼: '図解2' }], [], 'T1');
  const 二 = G.直しをまとめる_(一, [], [{ ID: 'z1', 結果: '直した', 文: 'ok', 直すセル: true }], 'T2', {});
  assert.strictEqual(二[0].状態, '作り直し中'); assert.strictEqual(二[0].セル待ち, true);
  const 三 = G.直しをまとめる_(二, [], [{ ID: 'z1', 結果: '直した', 文: 'ok', 直すセル: true }], 'T3', { z1: { 結果: '直した', 文: '' } });
  assert.strictEqual(三[0].状態, '済');
  const 別 = G.直しをまとめる_([], [{ ID: 'z2', 対象: '図解' }], [{ ID: 'z2', 結果: '直した', 文: 'ok', 直すセル: false }], 'T4', {});
  assert.strictEqual(別[0].状態, '済');   // 台本を合わせなかった図解の直しは、絵ができたら済
});

test('図の印_: 素材フォルダの画像を手で差し替えると印が変わる（ChatGPT での手直しでスライドを組み直す）', () => {
  const 観 = { ファイル: { 図解済み: { 日時: 'A', 更新: 'B' }, サムネイル: { 日時: 'C' }, 画像の更新: '2026-09-26T10:00:00+09:00' } };
  const 前 = G.図の印_(観);
  観.ファイル.画像の更新 = '2026-09-26T11:00:00+09:00';
  assert.notStrictEqual(G.図の印_(観), 前);
  assert.strictEqual(G.図の印_({ ファイル: { 図解済み: null } }), '');   // 図解がまだなら組まない
});

test('直しをまとめる_: タイトルの直しは台本のセルを書き終えても、【済】直し_（LP・サムネイルまで済）を見るまで作り直し中', () => {
  const 一 = G.直しをまとめる_([], [{ ID: 't1', 対象: 'タイトル', 依頼: 'x' }], [], 'T1', { t1: { 結果: '直した', 文: '' } });
  assert.strictEqual(一[0].状態, '作り直し中');
  const 二 = G.直しをまとめる_(一, [], [{ ID: 't1', 結果: '直した', 文: 'ok', 直すセル: true }], 'T2', { t1: { 結果: '直した', 文: '' } });
  assert.strictEqual(二[0].状態, '済');
  // 構成シートへ書けなかった（直せない）あとで【済】が「直した」で来ても、直せないのまま
  const 失 = G.直しをまとめる_([], [{ ID: 't2', 対象: 'タイトル' }], [], 'T1', { t2: { 結果: '直せない', 文: '違う' } });
  const 失2 = G.直しをまとめる_(失, [], [{ ID: 't2', 結果: '直した', 文: 'ok', 直すセル: true }], 'T2', { t2: { 結果: '直せない', 文: '違う' } });
  assert.strictEqual(失2[0].状態, '止まった'); assert.strictEqual(失2[0].結果, '直せない');
  const ス = G.直しをまとめる_([], [{ ID: 's1', 対象: 'スライド' }], [], 'T1', { s1: { 結果: '直した', 文: '' } });
  assert.strictEqual(ス[0].状態, '済');   // スライドの直しはセルを書けば済（今までどおり）
});

test('日時の文字_: 日本時間の +09:00 の形', () => {
  assert.strictEqual(G.日時の文字_(new Date('2026-09-26T01:15:00Z')), '2026-09-26T10:15:00+09:00');
});

test('成果物revision_: 同じ観測は同じ64桁、出力の更新で変わる', () => {
  const 観 = 全部できた観();
  const a = G.成果物revision_(観);
  const b = G.成果物revision_(JSON.parse(JSON.stringify(観)));
  assert.match(a, /^[a-f0-9]{64}$/); assert.strictEqual(a, b);
  観.ファイル.音声.日時 = '2026-09-26T11:00:00+09:00';
  assert.notStrictEqual(G.成果物revision_(観), a);
});

test('承認済みか_: 過去のrevision無しを推測補完せず、同じ生成版だけ通す', () => {
  const st = { getRange: () => ({ getValue: () => true }) };
  assert.strictEqual(G.新の流れの承認済みか_('id', st, { 生成: { revision: 'a'.repeat(64) }, 承認: { 済: true } }), false);
  assert.strictEqual(G.新の流れの承認済みか_('id', st, { 生成: { revision: 'a'.repeat(64) }, 承認: { 済: true, revision: 'b'.repeat(64) } }), false);
  assert.strictEqual(G.新の流れの承認済みか_('id', st, { 生成: { revision: 'a'.repeat(64) }, 承認: { 済: true, revision: 'a'.repeat(64) } }), true);
});

// ============================================================
// 4. 旧の流れが変わらない（変更前のコードと main() の書き込みを突き合わせる）
// ============================================================
const OLD = {
  h1: 'H1oldxxxxxxxxxxxxxxxxxxxxxxxxxx', h2: 'H2oldxxxxxxxxxxxxxxxxxxxxxxxxxx', h3: 'H3oldxxxxxxxxxxxxxxxxxxxxxxxxxx',
  h4: 'H4oldxxxxxxxxxxxxxxxxxxxxxxxxxx', h5: 'H5oldxxxxxxxxxxxxxxxxxxxxxxxxxx'
};
const K = (n) => 'K' + n + 'kouseixxxxxxxxxxxxxxxxxxxxxxxx';
const 開始のID = 'H9newxxxxxxxxxxxxxxxxxxxxxxxxxx';

/**
 * 偽の世界を作る。旧の案件は5件（名前はすべて架空の「過去◯社」）。
 * withNew=true なら、旧の案件のあとに新の流れの案件（テスト社）と【開始】ファイルを足す。
 */
function 世界(withNew, 付け足し) {
  const env = makeEnv();
  const P = env.props;
  const root = env.makeFolder('【固定】ウェビナー自動生成');
  const IN = root.addFolder('【入力】ヒアリングシート');
  const OUT = root.addFolder('【出力】生成物');
  const AS = root.addFolder('【素材】投影資料');
  const DEL = root.addFolder('【納品】', '1l3qvbW2u-4RLBvNov4ZUj-opEGapdZnJ');
  Object.assign(P, { INPUT_FOLDER_ID: IN.id, OUTPUT_FOLDER_ID: OUT.id, TEMPLATE_SLIDE_ID: 'T', ASSETS_FOLDER_ID: AS.id, MGMT_SHEET_ID: 'MGMTxxxxxxxxxxxxxxxxxxxxxxxxxxxx' });

  const mg = env.makeSpreadsheet('MGMTxxxxxxxxxxxxxxxxxxxxxxxxxxxx', ['管理']).sheets['管理'];
  mg.getRange(1, 1, 1, 7).setValues([['ヒアリングファイルID', '案件名', 'ステータス', '構成シートURL', 'スライドURL', 'メッセージ', '更新日時']]);
  const 構成 = (n, name) => {
    const ss = env.makeSpreadsheet(K(n), n === 9 ? ['設定', 'AI出力貼り付け', 'オファー'] : ['設定', 'AI出力貼り付け']);
    if (n === 9) ss.sheets['オファー'].getRange(1, 1, 3, 5).setValues([['No.', '基本構成', '小分類', '構成テキスト', '台本'], [40, 'a', 'b', '前の構成', '前の台本'], [41, 'a', 'b', 'x', 'y']]);
    ss.sheets['設定'].getRange('B1').setValue(name);
    return ss;
  };
  const 行 = (vals) => mg.appendRow(vals);
  const url = (n) => 'https://docs.google.com/spreadsheets/d/' + K(n) + '/edit';
  // 納品成功系は新仕様どおり、ヒアリングにFAQ 3組が確定している形を作る。
  // 欠損時の拒否は qa_cue_gate.test.cjs で別途検証し、ここで旧4点例外を復活させない。
  const 完成FAQ = (id) => {
    const hs = env.makeSpreadsheet(id, ['【オファー記入用】']);
    hs.sheets['【オファー記入用】'].getRange(1, 1, 7, 4).setValues([
      ['⒌ よくある質問（Q&A）', '', '', ''],
      ['', '', 'よくある質問：', '質問1'], ['', '', '記入欄：', '回答1'],
      ['', '', 'よくある質問：', '質問2'], ['', '', '記入欄：', '回答2'],
      ['', '', 'よくある質問：', '質問3'], ['', '', '記入欄：', '回答3']
    ]);
    return hs;
  };

  // 旧1: AI出力待ち。案件フォルダに台本・構成のtxt → 自動取り込み
  構成(1, '過去A社_ヒアリングシート');
  const A = OUT.addFolder('過去A社');
  A.addFile('【台本JSON】過去A社.txt', '[{"no":1}]'); A.addFile('【構成テキストJSON】過去A社.txt', '[{"no":1}]');
  if (付け足し && 付け足し.資料中) A.addFile('【指示】資料_過去A社.txt', '指示');
  行([OLD.h1, '過去A社_ヒアリングシート', 'AI出力待ち', url(1), '', '', 'X']);
  // 旧2: レビュー待ち・B2（承認）ON → スライド生成
  構成(2, '過去B社_ヒアリングシート').sheets['設定'].getRange('B2').setValue(true);
  OUT.addFolder('過去B社');
  行([OLD.h2, '過去B社_ヒアリングシート', 'レビュー待ち', url(2), '', '', 'X']);
  // 旧3: 完了・B16（作り直し）ON・B22（納品）ON。音声あり
  const s3 = 構成(3, '過去C社_ヒアリングシート').sheets['設定'];
  完成FAQ(OLD.h3);
  s3.getRange('A16').setValue('直したら必ずここ｜セルや素材の直しをスライドへ反映する'); s3.getRange('B16').setValue(true);
  s3.getRange('A21').setValue('― 検品が終わったら納品する ―'); s3.getRange('A22').setValue('4点'); s3.getRange('B22').setValue(true);
  const C = OUT.addFolder('過去C社');
  C.addFile('【音声】過去C社.mp3', 'mp3', 'audio/mpeg'); C.addFile('【スライド】過去C社_ヒアリングシート', 's', 'application/vnd.google-apps.presentation');
  DEL.addFolder('過去C社');
  行([OLD.h3, '過去C社_ヒアリングシート', '完了', url(3), 'https://slide/3', '', 'X']);
  // 旧4: 完了・B19（タイトル確定）ON・B14（不具合の調査）ON
  const s4 = 構成(4, '過去D社_ヒアリングシート').sheets['設定'];
  s4.getRange('B19').setValue(true); s4.getRange('B13').setValue('42枚目がはみ出す'); s4.getRange('B14').setValue(true);
  const D = OUT.addFolder('過去D社');
  D.addFile('【保留】サムネ指示_過去D社.txt', 'x'); D.addFile('【タイトル案】過去D社.txt', 't');
  行([OLD.h4, '過去D社_ヒアリングシート', '完了', url(4), 'https://slide/4', '', 'X']);
  // 旧5: エラー（何もしない）
  構成(5, '過去E社_ヒアリングシート');
  行([OLD.h5, '過去E社_ヒアリングシート', 'エラー', url(5), '', 'x', 'X']);

  if (withNew) {
    // 新1: レビュー待ち・B2 を誰かが先に押した・B22 も押した（承認前なので納品しない）・図解済みあり
    const n1 = 構成(9, '動画撮影テスト社_ヒアリングシート').sheets['設定'];
    完成FAQ('H8newxxxxxxxxxxxxxxxxxxxxxxxxxxx');
    n1.getRange('B2').setValue(true); n1.getRange('A21').setValue('―'); n1.getRange('A22').setValue('4点'); n1.getRange('B22').setValue(true);
    const N = OUT.addFolder('動画撮影テスト社');
    N.addFile('【保留】サムネ指示_動画撮影テスト社.txt', 'x'); N.addFile('【タイトル案】動画撮影テスト社.txt', 't');
    AS.addFolder('動画撮影テスト社_ヒアリングシート').addFile('【図解済み】.json', JSON.stringify({ 枚数: 15, 日時: '2026-09-26T10:40:00+09:00', 止まった: [] }), 'application/json');
    DEL.addFolder('動画撮影テスト社');
    mg.appendRow(['H8newxxxxxxxxxxxxxxxxxxxxxxxxxxx', '動画撮影テスト社_ヒアリングシート', 'レビュー待ち', url(9), '', '', 'X', '新']);
    // 【開始】: 別のテスト社
    OUT.addFolder('開始テスト社').addFile('【開始】開始テスト社.json', JSON.stringify({ 案件名: '開始テスト社', ヒアリングシートID: 開始のID, 台帳の行: 20, 置いた日時: '2026-09-26T10:00:00+09:00' }), 'application/json');
  }
  return env;
}

/** 重い処理（Slides・Docs・AI出力の解析）は、どちらのコードでも同じ偽物に差し替える */
function 差し替え(ctx, env) {
  ctx.generateSlides_ = (kid) => { env.log.push(['generateSlides', kid]); if (env.組めない) throw new Error('組めない（テスト）'); return { pres: { getUrl: () => 'https://slide/new/' + kid }, warnings: ['w'] }; };
  ctx.importAiOutput_ = (ss, texts) => { env.log.push(['import', ss.getId(), texts ? 'txt' : 'paste']); return ['警告x']; };
  ctx.DocumentApp = { openById: () => ({ getBody: () => ({ getText: () => 'サブタイトル\nメインタイトル' }) }), create: () => { throw new Error('使わない'); } };
  // DocumentApp自体はこの巨大なフロー試験の対象外なので、生成物の置き方だけを偽装する。
  // ただしFAQの3組必須判定は本物の関数を通す。欠損拒否は qa_cue_gate.test.cjs が実物のsyncQaCue_で検証する。
  ctx.syncQaCue_ = (_cfg, pname) => {
    const rows = env.spreadsheets.MGMTxxxxxxxxxxxxxxxxxxxxxxxxxxxx.sheets['管理'].getDataRange().getValues();
    const row = rows.slice(1).find(r => String(r[1]) === String(pname));
    const hs = row && env.spreadsheets[row[0]], offer = hs && hs.getSheetByName('【オファー記入用】');
    if (!offer) throw new Error('ヒアリングシートにオファー記入用タブがありません');
    ctx.qaCueRequireCompleteFaq_(offer.getDataRange().getValues());
    const folder = ctx.caseFolderIfExists_(_cfg, pname), company = ctx.caseCompany_(pname);
    const docName = ctx.qaCueFileName_(company), pdfName = ctx.qaCuePdfFileName_(company);
    if (!folder.getFilesByName(docName).hasNext()) folder.addFile(docName, 'FAQ 3組', 'application/vnd.google-apps.document');
    if (!folder.getFilesByName(pdfName).hasNext()) folder.addFile(pdfName, 'FAQ 3組', 'application/pdf');
  };
  ctx.createPromptDoc_ = (pname, hid) => { env.log.push(['createPromptDoc', pname, hid]); return { getUrl: () => 'https://doc/' + hid }; };
  ctx.createKouseiScaffold_ = (pname, hid) => {
    env.log.push(['createKouseiScaffold', pname, hid]);
    const ss = env.makeSpreadsheet('Knewkouseixxxxxxxxxxxxxxxxxxxxxxx', ['設定', 'AI出力貼り付け']);
    ss.sheets['設定'].getRange('B1').setValue(pname);
    return ss;
  };
  ctx.createUrlSheet_ = () => {};
}

function 回す(rev, withNew) {
  const env = 世界(withNew);
  const ctx = loadGas(env, rev);
  差し替え(ctx, env);
  env.log.length = 0;   // 世界を作ったときの記録は比べない
  ctx.main();
  return env;
}
/** vm の中で作られた配列は外と原型が違うので、JSONにして比べる */
const 平ら = (x) => JSON.parse(JSON.stringify(x));

/** 記録のうち、旧の案件にかかわるものだけ（新の流れの案件の分を除く） */
const 旧の印 = ['過去', 'K1kousei', 'K2kousei', 'K3kousei', 'K4kousei', 'K5kousei'].concat(Object.values(OLD));
function 旧の分(log) {
  return log.filter(e => { const s = JSON.stringify(e); return 旧の印.some(k => s.indexOf(k) >= 0) && s.indexOf('テスト社') < 0; });
}
/** 監視専用inventoryは旧案件の業務書込みではないので、旧版との業務副作用比較から除く。 */
function 業務の分(log) { return log.filter(e => JSON.stringify(e).indexOf('【状態】資料実物_webinar-gas.json') < 0); }
/** 今回承認された5点納品の表示差だけを旧基準へ正規化する。業務処理の他差分は比較で検出する。 */
function 五点納品を旧基準へ_(log) {
  return 業務の分(log).filter(e => !(e[0] === 'copy' && /【Q&Aカンペ】/.test(JSON.stringify(e)))).map(e => {
    const x = 平ら(e), s = JSON.stringify(x);
    if (s.indexOf('5点を【納品】へコピー') >= 0) x[3][1][0] = '4点を【納品】へコピー（チェックすると実行）';
    if (s.indexOf('Q&Aカンペ') >= 0 && x[0] === 'set' && x[2] === 'B23') x[3] = String(x[3]).replace(/・Q&Aカンペ(?:PDF)?(?:\([^)]*\))?/g, '');
    return x;
  });
}

test('旧の流れ: 変更前（' + BASE + '）と今で、main() の書き込みが1つ残らず同じ', () => {
  const 前 = 回す(BASE, false).log;
  const 今 = 回す(null, false).log;
  assert.ok(前.length > 20, '比べる記録が少なすぎる: ' + 前.length);
  assert.deepStrictEqual(五点納品を旧基準へ_(今), 五点納品を旧基準へ_(前));
});

test('旧の流れ: 新の流れの案件と【開始】が混ざっていても、旧の案件への書き込みは変更前と同じ', () => {
  const 前 = 旧の分(回す(BASE, false).log);
  const 今 = 旧の分(回す(null, true).log);
  assert.ok(前.length > 20);
  assert.deepStrictEqual(五点納品を旧基準へ_(今), 五点納品を旧基準へ_(前));
});

test('取り込み: Mac がまだ資料を作っている（【指示】資料_ が残っている）うちは、txt があっても取り込まない', () => {
  const env = 世界(false, { 資料中: true });
  const ctx = loadGas(env, null); 差し替え(ctx, env); env.log.length = 0; ctx.main();
  assert.ok(!env.log.some(e => e[0] === 'import' && e[1] === K(1)), JSON.stringify(env.log.filter(e => e[0] === 'import')));
});

test('旧の流れ: 見てほしい動きが実際に起きている（取り込み・B2でスライド・B16で音声を消す・B19・B22）', () => {
  const log = 回す(null, false).log;
  const s = JSON.stringify(log);
  assert.ok(log.some(e => e[0] === 'import' && e[1] === K(1)));
  assert.ok(log.some(e => e[0] === 'generateSlides' && e[1] === K(2)));   // 旧は B2 でスライド
  assert.ok(log.some(e => e[0] === 'trash' && /過去C社\/【音声】過去C社\.mp3/.test(e[1])));   // 旧の B16 は音声を消す
  assert.ok(log.some(e => e[0] === 'rename' && e[2] === '【指示】サムネ_過去D社.txt'));
  assert.ok(s.indexOf('各タブをレビューし、設定タブの承認チェックをONにしてください') >= 0);   // 旧の案内の文のまま
  assert.ok(log.some(e => e[0] === 'copy' && /過去C社\/【スライド】/.test(e[1])));   // B22で納品した（音声はB16で先に消えている）
  assert.ok(log.some(e => e[0] === 'set' && e[1] === K(3) + '/設定' && e[2] === 'B23' && /納品しました: スライド/.test(e[3])));
  assert.ok(!log.some(e => e[0] === 'set' && e[2] === 'H2'));   // 管理シートの H列には何も書かない
  assert.ok(!log.some(e => e[0] === 'formula'));   // 旧の構成シートに操作パネルのリンクを置かない
});

// ============================================================
// 5. 新の流れの動き（偽の世界で main() を回す）
// ============================================================
test('新の流れ: 【開始】を拾って登録（H列＝新）・名前を【済】開始_に・【入力】を通さない', () => {
  const env = 回す(null, true);
  const log = env.log;
  assert.ok(log.some(e => e[0] === 'createPromptDoc' && e[1] === '開始テスト社_ヒアリングシート' && e[2] === 開始のID));
  assert.ok(log.some(e => e[0] === 'rename' && /【開始】開始テスト社\.json$/.test(e[1]) && e[2] === '【済】開始_開始テスト社.json'));
  const mg = env.spreadsheets['MGMTxxxxxxxxxxxxxxxxxxxxxxxxxxxx'].sheets['管理'];
  const rows = mg.getDataRange().getValues();
  const r = rows.find(x => x[0] === 開始のID);
  assert.strictEqual(r[2], 'AI出力待ち');
  assert.strictEqual(r[7], '新');
  assert.strictEqual(rows[0][7], '流れ');
  // 2回目は登録し直さない
  const ctx = loadGas(env); 差し替え(ctx, env); env.log.length = 0;
  ctx.main();
  assert.ok(!env.log.some(e => e[0] === 'createPromptDoc'));
});

test('新の流れ: B2を先に押しても承認にならず外れ、B22も承認前なので納品しない。図解がそろえばスライドを組む', () => {
  const env = 回す(null, true);
  const log = env.log;
  const st = env.spreadsheets[K(9)].sheets['設定'];
  assert.ok(!log.some(e => e[0] === 'copy' && /テスト社/.test(JSON.stringify(e))), '承認前に納品した');
  assert.ok(/まだ承認されていない/.test(st.getRange('B23').getValue()));
  assert.ok(log.some(e => e[0] === 'generateSlides' && e[1] === K(9)), '図解がそろったのにスライドを組んでいない');
  assert.strictEqual(st.getRange('B2').getValue(), false);   // 全部できる前のB2は外れる
  // タイトル1位の自動採用（LPの仕事が終わっている）
  assert.ok(log.some(e => e[0] === 'rename' && e[2] === '【指示】サムネ_動画撮影テスト社.txt'));
  // 進み.json を書いた
  const 進 = JSON.parse(Object.values(env.files).find(f => f.name === '【状態】進み.json' && /動画撮影テスト社/.test(f.parent.path)).content);
  assert.strictEqual(進.流れ, '新');
  assert.strictEqual(進.案件名, '動画撮影テスト社');
  assert.strictEqual(進.系統.資料.find(x => x.名 === 'スライド').状態, '済');
  assert.strictEqual(進.承認.済, false);
});

test('新の流れ: LPの仕事中（【指示】LP_ が残っている）は、【保留】サムネ指示の名前を変えない', () => {
  const env = 世界(true);
  const N = env.folders[Object.keys(env.folders).find(k => env.folders[k].name === '動画撮影テスト社' && /出力/.test(env.folders[k].path))];
  N.addFile('【指示】LP_動画撮影テスト社.txt', 'x');
  const ctx = loadGas(env); 差し替え(ctx, env); env.log.length = 0;
  ctx.main();
  assert.ok(!env.log.some(e => e[0] === 'rename' && e[2] === '【指示】サムネ_動画撮影テスト社.txt'));
});

test('新の流れ: 全部できたあとパネルの【承認】→ 承認・B2 ON、【納品指示】→ 納品。直しが来たら承認が外れる', () => {
  const env = 回す(null, true);   // 1回目: スライドを組んだ
  const N = env.folders[Object.keys(env.folders).find(k => env.folders[k].name === '動画撮影テスト社' && /出力/.test(env.folders[k].path))];
  N.addFile('【LP】動画撮影テスト社', 'lp', 'application/vnd.google-apps.document');
  N.addFile('【音声】動画撮影テスト社.mp3', 'mp3', 'audio/mpeg');
  N.addFile('【サムネイル】動画撮影テスト社.png', 'png', 'image/png');
  N.addFile('【承認】動画撮影テスト社.json', JSON.stringify({ 人: '福田真里（PM）', メール: 'pm@example.com', 日時: '…' }));
  const run = () => { const ctx = loadGas(env); 差し替え(ctx, env); env.log.length = 0; ctx.main(); return ctx; };
  let ctx = run();
  const st = env.spreadsheets[K(9)].sheets['設定'];
  const 記 = () => JSON.parse(env.props['SHINKO_H8newxxxxxxxxxxxxxxxxxxxxxxxxxxx']);
  assert.strictEqual(記().承認.済, true, JSON.stringify(記().動き));
  assert.match(記().生成.revision, /^[a-f0-9]{64}$/);
  assert.strictEqual(記().承認.revision, 記().生成.revision);
  assert.strictEqual(st.getRange('B2').getValue(), true);
  assert.ok(env.log.some(e => e[0] === 'rename' && e[2] === '【済】承認_動画撮影テスト社.json'));
  // 納品
  N.addFile('【納品指示】動画撮影テスト社.json', JSON.stringify({ 人: '福田真里（PM）' }));
  run();
  assert.strictEqual(記().納品.済, true, JSON.stringify(記().動き));
  assert.strictEqual(記().納品.revision, 記().承認.revision);
  // ふつうに入ったときは「入れられませんでした」も★も出ない
  assert.ok(!JSON.stringify(記().動き).includes('入れられませんでした'), JSON.stringify(記().動き));
  assert.ok(!JSON.stringify(記().動き).includes('入れ直していません'));
  assert.ok(env.log.some(e => e[0] === 'copy' && /動画撮影テスト社\/【音声】/.test(e[1])));
  // 直しが来たら承認が外れる → 【済】直し（スライド・直した）で、音声を残してスライドだけ組み直す
  N.addFile('【指示】直し_動画撮影テスト社_a1.json', JSON.stringify({ ID: 'a1', 対象: 'スライド', 依頼: '12枚目', 頼んだ人: '福田真里', 元: 'PM相談室', 受付: '…' }));
  run();
  assert.strictEqual(記().承認.済, false);
  assert.strictEqual(st.getRange('B2').getValue(), false);
  N.files.find(f => f.name === '【指示】直し_動画撮影テスト社_a1.json').trashed = true;
  N.addFile('【済】直し_動画撮影テスト社_a1.json', JSON.stringify({ ID: 'a1', 結果: '直した', 前: 'u1', 後: 'u2', 文: '直しました' }));
  run();
  // 【済】だけでは組み直さない（【直すセル】を書くまで「作り直し中」）
  assert.ok(!env.log.some(e => e[0] === 'generateSlides' && e[1] === K(9)));
  assert.strictEqual(記().直し[0].状態, '作り直し中');
  N.addFile('【直すセル】動画撮影テスト社_a1.json', JSON.stringify({ ID: 'a1', 案件名: '動画撮影テスト社', 対象: 'スライド', 直す: [{ no: 40, 欄: '構成テキスト', 列: 'D', 前: '前の構成', 後: '後の構成' }], 日時: '…' }));
  run();
  const of = env.spreadsheets[K(9)].sheets['オファー'];
  assert.strictEqual(of.getRange('D2').getValue(), '後の構成');
  assert.strictEqual(of.getRange('E2').getValue(), '前の台本');
  assert.ok(env.log.some(e => e[0] === 'generateSlides' && e[1] === K(9)));
  assert.ok(!env.log.some(e => e[0] === 'trash' && /【音声】動画撮影テスト社/.test(e[1])), 'スライドの直しで音声を消した');
  assert.ok(env.log.some(e => e[0] === 'rename' && e[2] === '【済】直すセル_動画撮影テスト社_a1.json'));
  assert.strictEqual(記().直し[0].状態, '済');
  // 台本の直しは音声も作り直す（音声をゴミ箱へ）
  N.addFile('【直すセル】動画撮影テスト社_b2.json', JSON.stringify({ ID: 'b2', 対象: '台本', 直す: [{ no: 40, 欄: '台本', 列: 'E', 前: '前の台本', 後: '後の台本' }] }));
  run();
  assert.strictEqual(of.getRange('E2').getValue(), '後の台本');
  assert.ok(env.log.some(e => e[0] === 'trash' && /【音声】動画撮影テスト社/.test(e[1])));
  assert.strictEqual(記().直し.find(d => d.ID === 'b2').状態, '済');
  // 「前」が今のセルと違えば書かずに知らせ、直しは止まった
  N.addFile('【直すセル】動画撮影テスト社_c3.json', JSON.stringify({ ID: 'c3', 対象: 'スライド', 直す: [{ no: 41, 列: 'D', 前: '違う文', 後: 'z' }] }));
  run();
  assert.strictEqual(of.getRange('D3').getValue(), 'x');
  assert.ok(!env.log.some(e => e[0] === 'generateSlides'));
  assert.strictEqual(記().直し.find(d => d.ID === 'c3').状態, '止まった');
  assert.ok(Object.values(env.files).some(f => /【知らせ待ち】$/.test(f.parent.path) && /直し（c3）/.test(f.content)));
  // ★1回の投稿でまとめて頼んだ直し（2026-09-27）：届いている【直すセル】を全部書いて、組み直すのは1回だけ。
  //   同じセルを続けて直しても（e5 の後 → f6 の前）、受けた順に書くので両方入る。1つが違っても、ほかは入れる
  N.addFile('【直すセル】動画撮影テスト社_f6.json', JSON.stringify({ ID: 'f6', 対象: 'スライド', 直す: [{ no: 40, 欄: '構成テキスト', 列: 'D', 前: 'まとめ1', 後: 'まとめ2' }], 日時: '2026-09-27T10:00:02+09:00' }));
  N.addFile('【直すセル】動画撮影テスト社_e5.json', JSON.stringify({ ID: 'e5', 対象: 'スライド', 直す: [{ no: 40, 欄: '構成テキスト', 列: 'D', 前: '後の構成', 後: 'まとめ1' }], 日時: '2026-09-27T10:00:01+09:00' }));
  N.addFile('【直すセル】動画撮影テスト社_g7.json', JSON.stringify({ ID: 'g7', 対象: 'スライド', 直す: [{ no: 41, 列: 'D', 前: '違う文', 後: 'z' }], 日時: '2026-09-27T10:00:03+09:00' }));
  run();
  assert.strictEqual(of.getRange('D2').getValue(), 'まとめ2');
  assert.strictEqual(env.log.filter(e => e[0] === 'generateSlides' && e[1] === K(9)).length, 1, '組み直しは1回だけ');
  ['e5', 'f6'].forEach(id => assert.ok(env.log.some(e => e[0] === 'rename' && e[2] === '【済】直すセル_動画撮影テスト社_' + id + '.json'), id));
  assert.ok(env.log.some(e => e[0] === 'rename' && e[2] === '【止まった】直すセル_動画撮影テスト社_g7.json'));
  assert.ok(JSON.stringify(記().動き).includes('直し（2件・2か所）'), JSON.stringify(記().動き.slice(-3)));
  // セルは書けたがスライドを組めなかった → 【止まった】直すセル_（ボットが「直せませんでした」と返す）
  env.組めない = true;
  N.addFile('【直すセル】動画撮影テスト社_d4.json', JSON.stringify({ ID: 'd4', 対象: 'スライド', 直す: [{ no: 40, 欄: '構成テキスト', 列: 'D', 前: 'まとめ2', 後: 'もっと後の構成' }] }));
  run();
  assert.ok(env.log.some(e => e[0] === 'rename' && e[2] === '【止まった】直すセル_動画撮影テスト社_d4.json'), JSON.stringify(env.log.filter(e => e[0] === 'rename')));
  env.組めない = false;
});

test('直すセルを当てる_: No・列で当て、前が違えば1つも書かない。No.96の松竹梅の2行目も見る', () => {
  const タブ = { オファー: [['No.'], [40, '', '', '構40', '台40'], [96, '', '', '（斜線）', '（斜線）'], ['', '', '', '松の構成', '松の台本']] };
  let r = G.直すセルを当てる_(タブ, [{ no: 40, 列: 'D', 前: '構40', 後: '新' }, { no: 96, 列: 'E', 前: '松の台本', 後: '新松' }]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r.書く)), [{ タブ: 'オファー', 行: 2, 列: 4, 後: '新' }, { タブ: 'オファー', 行: 4, 列: 5, 後: '新松' }]);
  r = G.直すセルを当てる_(タブ, [{ no: 40, 列: 'D', 前: '構40', 後: '新' }, { no: 40, 列: 'E', 前: 'ちがう', 後: 'x' }]);
  assert.strictEqual(r.書く.length, 0); assert.ok(/手で直されたかも/.test(r.違う[0]));
  // 取り込みで図解タグを足した行：タグを外して比べ、書くときはタグを残す
  const 札つき = { オファー: [['No.'], [36, '', '', 'デメリット\n\n図解（図解2）', '台']] };
  r = G.直すセルを当てる_(札つき, [{ no: 36, 列: 'D', 前: 'デメリット', 後: 'デメリットは3つ' }]);
  assert.strictEqual(r.違う.length, 0, r.違う.join());
  assert.strictEqual(r.書く[0].後, 'デメリットは3つ\n\n図解（図解2）');
  // 2回目（組むのが次の回へ回った）：セルはもう「後＋タグ」。当てられる
  const 二回目 = { オファー: [['No.'], [36, '', '', 'デメリットは3つ\n\n図解（図解2）', '台']] };
  r = G.直すセルを当てる_(二回目, [{ no: 36, 列: 'D', 前: 'デメリット', 後: 'デメリットは3つ' }]);
  assert.strictEqual(r.違う.length, 0, r.違う.join()); assert.strictEqual(r.書く[0].後, 'デメリットは3つ\n\n図解（図解2）');
  // 前にタグがあり、後でタグが落ちた（AIが落とした）：タグを残す
  r = G.直すセルを当てる_(札つき, [{ no: 36, 列: 'D', 前: 'デメリット\n\n図解（図解2）', 後: 'デメリットは4つ' }]);
  assert.strictEqual(r.書く[0].後, 'デメリットは4つ\n\n図解（図解2）');
  r = G.直すセルを当てる_(タブ, [{ no: 77, 列: 'D', 前: '', 後: 'x' }]);
  assert.ok(/ありません/.test(r.違う[0]));
});

test('見張りの様子_・進み: Macの心拍が20分より古いと「生成の担当が止まっています」。心拍が1つも無ければ何も言わない', () => {
  assert.strictEqual(G.見張りの様子_([], いま).止まった, false);
  const 古 = G.見張りの様子_([{ 最後に回った時刻: '2026-09-26T10:00:00+09:00' }, { 最後に回った時刻: '2026-09-26T09:00:00+09:00' }], いま);
  assert.strictEqual(古.止まった, true); assert.strictEqual(古.分, 30);
  assert.strictEqual(G.見張りの様子_([{ 最後に回った時刻: '2026-09-26T10:25:00+09:00' }], いま).止まった, false);
  const 観 = 観の元({ ファイル: { 指示LP: true } }); 観.見張り = 古;
  const 進 = G.進みを計算_(観, G.進み_記録の既定_(), いま);
  assert.strictEqual(進.いまここ.文, '生成の担当が止まっています（宮津さんへ知らせ済み）');
  assert.strictEqual(進.見張り.止まった, true);
  const 観2 = 全部できた観(); 観2.見張り = 古;   // 全部できたあとは、見張りが止まっていても言わない
  assert.strictEqual(G.進みを計算_(観2, G.進み_記録の既定_(), いま).いまここ.文, '全部できました');
});

test('新の流れ: 心拍の古いファイルがあれば 進み.json に出す。テスト案件しか無いときは、宮津さんへの知らせは置かない', () => {
  const env = 世界(true);
  const 固定 = Object.values(env.folders).find(d => d.name === '【固定】ウェビナー自動生成');
  固定.addFile('【心拍】資料の見張り_mac1.json', JSON.stringify({ PC: 'mac1', 最後に回った時刻: '2020-01-01T00:00:00+09:00' }), 'application/json');
  for (let k = 0; k < 2; k++) { const ctx = loadGas(env); 差し替え(ctx, env); ctx.main(); }
  const 進 = JSON.parse(Object.values(env.files).find(f => f.name === '【状態】進み.json' && /動画撮影テスト社/.test(f.parent.path)).content);
  assert.strictEqual(進.見張り.止まった, true);
  const 通 = Object.values(env.files).filter(f => /^【知らせ】見張り_/.test(f.name));
  // ★テスト案件だけのとき（Mac が眠るたびに社内へ出ないように。2026-09-27）。本物のまだ納品していない案件があるときだけ1通置く
  assert.strictEqual(通.length, 0);
});

test('新の流れ: 図解だけ差し替えたら（【図解済み】を書き直し）、音声を残してスライドだけ組み直す', () => {
  const env = 回す(null, true);
  const af = Object.values(env.files).find(f => f.name === '【図解済み】.json');
  const N = env.folders[Object.keys(env.folders).find(k => env.folders[k].name === '動画撮影テスト社' && /出力/.test(env.folders[k].path))];
  N.addFile('【音声】動画撮影テスト社.mp3', 'mp3', 'audio/mpeg');
  af.content = JSON.stringify({ 枚数: 15, 日時: '2026-09-26T11:40:00+09:00', 止まった: [] });
  const ctx = loadGas(env); 差し替え(ctx, env); env.log.length = 0; ctx.main();
  assert.ok(env.log.some(e => e[0] === 'generateSlides' && e[1] === K(9)));
  assert.ok(!env.log.some(e => e[0] === 'trash' && /【音声】/.test(e[1])));
  // 変わっていなければ組み直さない
  env.log.length = 0; const ctx2 = loadGas(env); 差し替え(ctx2, env); ctx2.main();
  assert.ok(!env.log.some(e => e[0] === 'generateSlides' && e[1] === K(9)));
});

test('新の流れ: 止まったらテスト案件の知らせを 進み.json の「知らせ」に積む（外へは送らない）', () => {
  const env = 世界(true);
  const ctx = loadGas(env); 差し替え(ctx, env);
  ctx.generateSlides_ = () => { throw new Error('Exception: テンプレが開けない'); };
  ctx.UrlFetchApp = { fetch: () => { throw new Error('外へ送ってはいけない'); } };
  ctx.main();
  const 進 = JSON.parse(Object.values(env.files).find(f => f.name === '【状態】進み.json' && /動画撮影テスト社/.test(f.parent.path)).content);
  const 知 = 進.知らせ.find(x => /スライド/.test(x.文));
  assert.ok(知, JSON.stringify(進.知らせ));
  assert.strictEqual(知.テスト, true);
  assert.ok(/^動画撮影テスト社 … /.test(知.文));
  // 【固定】ウェビナー自動生成/【知らせ待ち】/ に1通1ファイル（宛先＝テスト）
  const 置いた = Object.values(env.files).filter(f => /【知らせ待ち】$/.test(f.parent.path) && /^【知らせ】動画撮影テスト社_\d{8}-\d{6}_\d+\.json$/.test(f.name));
  assert.ok(置いた.length >= 1, Object.values(env.files).map(f => f.parent.path + '/' + f.name).join('\n'));
  const o = JSON.parse(置いた[0].content);
  assert.strictEqual(o.宛先, 'テスト'); assert.strictEqual(o.案件名, '動画撮影テスト社');
  assert.ok(/止まりました/.test(o.文) && /\+09:00$/.test(o.置いた日時));
});

test('知らせの中身_: テスト案件は宛先「テスト」、それ以外は「社内」', () => {
  const t = new Date('2026-09-26T01:15:00Z');
  assert.strictEqual(G.知らせの中身_('過去A社', 'x', false, t).宛先, '社内');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(G.知らせの中身_('動画撮影テスト社', 'x', true, t))), { 宛先: 'テスト', 文: '動画撮影テスト社 … x', 案件名: '動画撮影テスト社', 置いた日時: '2026-09-26T10:15:00+09:00' });
  assert.strictEqual(G.知らせのファイル名_('動画撮影テスト社', t, 3), '【知らせ】動画撮影テスト社_20260926-101500_3.json');
});

// ============================================================
// 6. テスト案件の片付け（撮り直し用）
// ============================================================
test('片付け: テスト案件の名前でないと止まる（何も消さない）', () => {
  const env = 回す(null, true);
  const ctx = loadGas(env); 差し替え(ctx, env); env.log.length = 0;
  assert.throws(() => ctx.テスト案件を片付ける_('過去A社', false), /テスト案件の名前ではない/);
  assert.throws(() => ctx.テスト案件を片付ける_('', false), /空です/);
  assert.strictEqual(env.log.length, 0);
});

test('片付け: 下見は一覧だけ。本番はゴミ箱へ・管理シートの行と覚えを消す（旧の案件には触らない）', () => {
  const env = 回す(null, true);
  const ctx = loadGas(env); 差し替え(ctx, env); env.log.length = 0;
  const 下見 = ctx.テスト案件を片付ける_('動画撮影テスト社', true);
  assert.strictEqual(env.log.length, 0);
  assert.ok(下見.some(x => /管理シートの/.test(x)) && 下見.some(x => /【状態】進み\.json/.test(x)) && 下見.some(x => /【図解済み】/.test(x)));
  ctx.テスト案件を片付ける_('動画撮影テスト社', false);
  const mg = env.spreadsheets['MGMTxxxxxxxxxxxxxxxxxxxxxxxxxxxx'].sheets['管理'];
  const 名 = mg.getDataRange().getValues().map(r => r[1]);
  assert.ok(名.indexOf('動画撮影テスト社_ヒアリングシート') < 0);
  assert.ok(名.indexOf('過去A社_ヒアリングシート') >= 0 && 名.indexOf('過去E社_ヒアリングシート') >= 0);
  assert.ok(env.props['SHINKO_H8newxxxxxxxxxxxxxxxxxxxxxxxxxxx'] === undefined);
  assert.ok(env.log.some(e => e[0] === 'trash' && /動画撮影テスト社\/【状態】進み\.json/.test(e[1])));
  assert.ok(!env.log.some(e => /過去/.test(JSON.stringify(e)) && e[0] !== 'deleteRow'));
});

test('止まった知らせ: 「◯分途絶えています」の分が増えても、同じ止まりは1回だけ知らせる', () => {
  const 記 = G.進み_記録の既定_();
  const 進 = 分 => ({ 系統: { LP: [], 資料: [{ 名: '図解・挿絵', 状態: '止まった', 文: 'Macからの知らせが' + 分 + '分途絶えています（図解）' }] } });
  [25, 30, 35, 40].forEach(function (分) { G.止まったを知らせる_('本物社', 進(分), 記); });
  assert.strictEqual(記.知らせ.length, 1, JSON.stringify(記.知らせ));
  // 直って、また止まったら知らせる
  G.止まったを知らせる_('本物社', { 系統: { LP: [], 資料: [] } }, 記);
  G.止まったを知らせる_('本物社', 進(21), 記);
  assert.strictEqual(記.知らせ.length, 2);
});

test('進み: Mac の仕事が3回続けて失敗して止まったら、作っている行を「止まった」に（待ちの行はそのまま）', () => {
  const 見 = G.見張りの様子_([{ 最後に回った時刻: '2026-09-26T10:28:00+09:00', 仕事: { lp: { 止まった: true }, audio: { 止まった: false }, zukai: { 止まった: true } } }], いま);
  assert.strictEqual(JSON.stringify(見.止まった仕事), '["lp","zukai"]');
  const 観 = 全部できた観();
  観.状態 = 'AI出力待ち'; 観.スライドURL = ''; 観.ファイル.LP = null; 観.ファイル.タイトル案 = null; 観.ファイル.指示LP = true; 観.ファイル.図解済み = null;
  観.ファイル.音声 = null; 観.ファイル.台本txt = null;
  観.見張り = 見;
  const 進 = G.進みを計算_(観, G.進み_記録の既定_(), いま);
  const 行 = 名 => 進.系統.LP.concat(進.系統.資料).find(x => x.名 === 名);
  assert.strictEqual(行('LP').状態, '止まった'); assert.ok(/現場相談/.test(行('LP').文));
  assert.strictEqual(行('音声').状態, '待ち');
  assert.strictEqual(進.押せる.承認, false);
});

test('進み: 止まった仕事の割り当て：直しの「作り直し中」は lp の止まりでは止まったにしない（naoshi・fix なら止まった）', () => {
  const 観 = 全部できた観();
  const 記 = G.進み_記録の既定_();
  記.直し = G.直しをまとめる_([], [{ ID: 'L1', 対象: 'LP', 依頼: 'x', 受付: 'T' }], [], 'T1');
  観.見張り = G.見張りの様子_([{ 最後に回った時刻: '2026-09-26T10:28:00+09:00', 仕事: { lp: { 止まった: true } } }], いま);
  let 行 = G.進みを計算_(観, 記, いま).系統.LP.find(x => x.名 === 'LP');
  assert.strictEqual(行.状態, '作り直し中');
  観.見張り = G.見張りの様子_([{ 最後に回った時刻: '2026-09-26T10:28:00+09:00', 仕事: { fix: { 止まった: true } } }], いま);
  行 = G.進みを計算_(観, 記, いま).系統.LP.find(x => x.名 === 'LP');
  assert.strictEqual(行.状態, '止まった');
});

console.log('\n' + pass + '件 通過');
