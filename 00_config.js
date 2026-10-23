/**
 * 00_config.gs（APIレス版・実測規定反映）
 * スクリプトプロパティ・106行定義テーブル（＋Q&Aの No.104.5）・レイアウト定義・文字数規定
 * ※ フォントサイズ・字数はサンプル1,2の全106枚×2の実測値
 *   （詳細は「文字数・フォントサイズ規定書.md」参照。本ファイルが単一ソース）
 */

function getConfig_() {
  const p = PropertiesService.getScriptProperties();
  const req = ['INPUT_FOLDER_ID', 'OUTPUT_FOLDER_ID', 'TEMPLATE_SLIDE_ID', 'ASSETS_FOLDER_ID', 'MGMT_SHEET_ID'];
  const cfg = {};
  req.forEach(function (k) {
    cfg[k] = p.getProperty(k);
    if (!cfg[k]) throw new Error('スクリプトプロパティ未設定: ' + k + '（bootstrapAll() を実行すると自動設定されます）');
  });
  return cfg;
}

// ============================================================
// タブ名・素材サブフォルダ名
// ============================================================
const TABS = { SETTINGS: '設定', INTRO: 'イントロ', PROBLEM: '問題教育', STORY: 'ストーリー', OFFER: 'オファー' };
const HEARING_TABS = { OFFER: '【オファー記入用】', PROBLEM: '【問題教育の記入用】', STORY: '【ストーリーの記入用】' };
const ASSET_SUB = { ICON_P: 'アイコン（紫）', ICON_W: 'アイコン（白）', BG: '背景画像', SELECT: '選択画像' }; // 全案件共通

// ユーザー指定の既定フォルダID（スクリプトプロパティ ICON_PURPLE_FOLDER_ID 等で上書き可 →
// 未設定なら素材ルート直下の同名サブフォルダ → それも無ければこの既定ID）
const DEFAULT_ASSET_IDS = {
  ICON_P: '1cbcsXZmF8BUt9LSoFC9_w7D3mHrvfobm',
  ICON_W: '1A0yNQC8MWu3_iJgN_WmClrEJ413VDpzi',
  BG: '1Cz-d4w3t3Iv4lloUVcg74m65JxZ1j5Ot',
  SELECT: '1jbtXYqdLM_5wDEkZ7agQ-GO0nGY0izaS'
};
const ASSET_PROP_KEYS = { ICON_P: 'ICON_PURPLE_FOLDER_ID', ICON_W: 'ICON_WHITE_FOLDER_ID', BG: 'BG_FOLDER_ID', SELECT: 'SELECT_BG_FOLDER_ID' };

/** アイコン/背景の選定対象行（ユーザー指定） */
function designTargets_() {
  return {
    iconPurple: [42,43,44,51,53,87,88,89,90],   // 円内・工程アイコン
    iconWhite: [80],                              // 特長カード上の白アイコン×3（／区切り。82,84へ自動再掲）
    bg: range_(3,5).concat(range_(8,10), range_(12,31), range_(57,70), [93,94,100]), // 背景画像フォルダ
    select: [72]                                  // 選択画像フォルダ
  };
}

/**
 * よくある質問（Q&A）の行番号。No.104（分割払い）の直後・No.105（BYAF）の前に入る。
 * ★104.5 にしたのは、既存の No.1〜106 の番号と並びを1つも変えないため（2026-09-25・宮津さん決定）。
 *   ヒアリングシート【オファー記入用】「⒌ よくある質問（Q&A）」に記入がある案件だけ、
 *   スライドを1枚（テンプレの No.104 を複製して作る）・台本・音声が入る。記入が無ければ何も入らない。
 * ★この行が無い構成シート（この仕組みより前に作った案件）では、Q&Aは一切入らない。→ generateSlides_
 */
const QA_NO = 104.5;

/** 斜線（該当なし）でスキップ可能な行 → スキップ時にテンプレへ入れる文言 */
const SKIP_ROWS = { 99: '特典②', 104: '分割払いも可能です。\n24回払いなら\n◯万円／月' };

/**
 * 不透明度ルール（サンプル1実測・単位%＝画像の表示濃度）
 * Googleスライドは画像replaceで不透明度が100%にリセットされ、GASに設定APIが無いため、
 * 白の半透明オーバーレイ（alpha = 1 - 値/100）を画像直上に重ねて等価な見た目を再現する。
 * 行別上書き: 構成シートG列（背景）に「ファイル名@25」のように%を付与できる。
 */
const OPACITY_RULES = {
  stepInactive: 20  // 工程の非現在ステップアイコン（実測: 20%）
};

/**
 * No.→背景不透明度の対応表（サンプル1全スライド実測を85/25/15の3水準に丸めたもの）
 * 素材は「背景画像」フォルダ内のサブフォルダ「85」「25」「15」に濃度別で格納する。
 * この表を書き換えれば選定元フォルダが変わる（単一ソース）。
 * ※実測の丸め: No.12(100%)→85 / No.29,70(70%)→85 / No.61(16%)→15
 */
const BG_OPACITY_MAP = {
  '85': [3,4,5,12,13,14,23,24,25,26,27,28,29,30,64,65,66,67,68,69,70,93,94,100],
  '25': [8,9,10,20,21,22,57,58,59,60],
  '15': [15,16,17,18,19,31,61,62,63]
};
function bgOpacityFor_(no) {
  if (BG_OPACITY_MAP['85'].indexOf(no) !== -1) return '85';
  if (BG_OPACITY_MAP['25'].indexOf(no) !== -1) return '25';
  if (BG_OPACITY_MAP['15'].indexOf(no) !== -1) return '15';
  return '85';
}

// ============================================================
// レイアウトタグ（テンプレデッキのノート1行目 {{TPL:xxx}} と対応）
// フォント規定は 05_bootstrap.gs のレイアウト生成が本規定どおりに実装する
// ============================================================
const LAYOUT = {
  COVER: 'cover',            // No.1 全面サムネイル
  QUESTION: 'question',      // 問いかけ 45pt
  IDEAL: 'ideal',            // 理想/お悩み 45pt（1行25字 or 2行各14字）+ICON
  NARRATIVE: 'narrative',    // リード30pt+メイン55pt +ICON
  TEXT55: 'text55',          // 大見出し 55pt 1-2行 +ICON
  DIAGRAM: 'diagram',        // 見出し20pt+図解IMG
  DIAGRAM40: 'diagram40',    // 見出し40pt+図解IMG（特徴詳細・特典）
  CASE: 'case',              // 事例 名前30pt+本文20pt+図解IMG
  SERVICE: 'service',        // サービス名45pt+概要15pt+モックアップIMG
  MERIT_INTRO: 'merit_intro',// No.41 帯20pt+数字円×3
  MERIT: 'merit',            // No.42-44 帯(No.41再掲)+円下見出し19pt+説明20pt
  POINT_INTRO: 'point_intro',// No.50 帯20pt+数字大円×2
  POINT1: 'point1',          // No.51 帯(No.50再掲)+円内30pt×1
  POINT2: 'point2',          // No.53 帯(No.50再掲)+円内30pt×2（①はNo.51再掲）
  CHALLENGE: 'challenge',    // No.55 タイトル55pt+箇条書き30pt
  FEATURES: 'features',      // No.80/82/84 サービス名45pt+カード3枚(見出し14pt/補足11pt)
  STEPS: 'steps',            // No.86 36pt+20pt+ステップ名12pt×4+アイコン72pt×4
  KOTEI: 'kotei',            // No.87-90 上部ステップ帯(アイコン72pt・非現在20%)+図解IMG
  RINEN: 'rinen',            // No.78/103 45pt+上下固定画像帯
  IMPACT: 'impact',          // No.77/91 70pt+40pt
  PRICE: 'price',            // No.96/102 45pt+価格行40pt
  INVEST: 'invest',          // No.92 30pt+金額80pt+30pt
  STRIKE: 'strike',          // No.95 invest構成+金額行に赤二重線
  PROFILE: 'profile',        // No.6
  QA: 'qa',                  // No.104.5 よくある質問（No.104を複製し、Q太字＋Aの組を最大3つ）
  S_ZOOM: 'static_zoom', S_MEIGEN: 'static_meigen', S_BYAF: 'static_byaf', S_OREI: 'static_orei'
};

// ============================================================
// 固定台本
// ============================================================
const FIXED_DAIHON = {
  2: 'では最初に、\n皆さん、今から1か月後のことを少しだけ想像してみてください。',
  7: '例えば、皆さん経営していてこんなお悩みはありませんか？\nもし、当てはまるものがあれば、挙手🖐️マークでリアクションしていただけると嬉しいです！',
  33: 'なぜ、このような問題が起こるのでしょうか？',
  38: 'では、どうすればいいのか？',
  49: 'では、どうすればいいのか？',
  97: 'さらに、本日ZOOM終了までにお申し込みいただける方には、追加特典をご用意しております。',
  101: '経営の神様ピーター・ドラッカーも、\n今、上手くいっているから、今の仕事のやり方だけを続ける経営者は必然的に失敗する。\nという名言があるように。\n新しい仕事のやり方、営業のやり方にチャレンジする皆様の背中を押すために。',
  105: 'どちらでも構いません。\n大切なのは、ご自身の意思で決断することです。',
  106: '以上で、\n本日のご案内はすべてとなります。\n\n最後までご視聴いただき、\n本当にありがとうございました。'
};

// ============================================================
// 文字数規定（実測ベース）: no範囲 → {fmt: プロンプト用書式説明, maxLine: 1行最大字数, maxLines: 最大行数}
// バリデーションは図解タグ行を除いた各行に適用（超過は警告・停止しない）
// ============================================================
function kouseiSpec_() {
  const S = {};
  function set(nos, fmt, maxLine, maxLines) {
    nos.forEach(function (no) { S[no] = { fmt: fmt, maxLine: maxLine, maxLines: maxLines }; });
  }
  set([3,4,5,8,9,10,93,94,100], '45pt。15字以内なら1行（最大25字）、16字以上は意味の切れ目で2行に分割し各行14字以内。', 25, 2);
  set([11], '45pt×2行。「ある〇〇会社の経営者」＋「〇〇社長の話」。各行12字以内。', 12, 2);
  set(range_(12,32).concat(range_(57,72)), '1行目=リード（状況説明・17字以内・省略可）、2行目=メイン（強調・15字以内・必ず1行。2行に分けない）。', 17, 2);
  set([60,61,62,63,65,66,68,69,70], '1行目=リード（状況説明・17字以内・省略可）、2行目=メイン（強調・15字以内・必ず1行。2行に分けない）。①②③などの丸数字・連番をテキストに含めない（ヒアリング項目名の「失敗①」「検証①」「型化①」等の番号を転記しない）。', 17, 2);
  set([34,47,48], '大見出し。1〜2行・各行17字以内。', 17, 2);
  set([39], '大見出し。1行・25字以内（改行しない）。', 25, 1);
  set([78,103], '理念。45pt・2行・各行12字以内（上下の実績画像帯は固定）。', 12, 2);
  set([35,36,37], '1行目=旧ビリーフ（No.34）の文言をカギ括弧や「という誤解」等の注釈なしで1行にしたもの（3行とも同一）＋図解タグ。', 24, 1);
  set([40], '1行目=新ビリーフ（No.39）の文言を1行にしたもの＋図解タグ。', 24, 1);
  set([52,54], '図解の見出し1行・20字以内＋図解タグ。', 20, 1);
  set([87,88,89,90], '図解タグのみ（見出し不要。上部のステップ帯はNo.86から自動反映）。', 20, 1);
  set([45,73,75], '1行目=「〇〇社長の事例」（9字以内）、2行目=「業種 役職」（例: IT会社 代表取締役・14字以内）、3行目以降=要約（1〜2行・各27字以内）＋図解タグ。', 27, 4);
  set([46,74,76], '要約のみ（1〜2行・各27字以内）＋図解タグ。社長名・業種役職は直前のBefore行から自動反映される。', 27, 2);
  set([56,79], '1行目=サービス名（15字以内）、2〜3行目=概要文2行（各45字以内）。概要文は台本の説明文とほぼ同一で「〜サービス。」の一文で終えること＋図解タグ。', 45, 4);
  set([41], '「（動作）することで得られる3つのメリット」形式・19字以内・1行。', 19, 1);
  set([42,43,44], '1行目=円内見出し（体言止め・9字以内）、2行目以降=説明文（1〜2行・計36字以内・各行23字以内）。', 23, 3);
  set([50], '「正しい〇〇のポイント」形式・14字以内・1行。', 14, 1);
  set([51,53], '円内テキスト。1行なら16字以内、2行に分ける場合は各8字以内。', 16, 2);
  set([55], '1行目=「よくある課題」、2〜4行目=「・」＋課題（・を除き各18字以内）×3点。', 19, 4);
  set([80], '1行目=サービス名（15字以内）、2〜4行目=「見出し／補足」（全角スラッシュ区切り。見出し11字以内・補足25字以内）×3。', 40, 4);
  set([81,83,85,98,99], '「特徴① 〇〇」等の見出し1行・13字以内＋図解タグ。', 13, 1);
  set([86], '1行目=サービス名（12字以内）、2行目=「導入フローは4ステップ」等（19字以内）、3〜6行目=ステップ名（各8字以内）×4。', 19, 6);
  set([77,91], '1行目=「◯社を超える」（◯は導入社数の実数・70pt）、2行目=「お客様で成果再現済み」（固定・40pt）。No.77とNo.91は完全に同一の文言にする。', 12, 2);
  set([92,95], '1行目=前段（17字以内）、2行目=金額のみ（9字以内・80pt表示）、3行目=後段（17字以内）。', 17, 3);
  set([96], '【単一】1行目=サービス名、2行目以降=「プラン名 金額」。契約条件の補足がある場合は金額の直後に（◯ヶ月契約）の形で書く（改行は不要。表示時に自動で改行・縮小される）。【松竹梅】1行目=サービス名、2〜3行目=サービス概要2行（No.79の概要2行と完全に同一・各45字以内）、続けて3行=「プラン名｜プラン概要（27字以内）｜月額○万円」（全角パイプ｜区切り）、最終行=共通注記（例: ※共通：全プラン初期費用10万円）。JSONに taikei フィールド（"単一" or "松竹梅"）を必ず含める。', 45, 7);
  set([102], '【単一】1行目=バッジのOFF表記（例: 10万円OFF、初月無料。初期費用には触れない）、2行目以降=「商品名 値引き後表示」（No.96と同じ商品名＋半角スペース＋値引き後の表示。1行につき1商品。例: 動画講座 40万円、1on1サポート 3ヶ月無料。値引きのない商品は定価をそのまま書く）。サービス名・商品名・定価はNo.96から自動反映。【松竹梅】1行目=値引きオファーの短い表記1行（例: 初月無料、月額5万円OFF。初期費用には触れない）、2〜4行目=各プランの値引き後表示のみ（プラン名・OFF額・括弧書きを付けない。初月無料の場合は「初月無料」とだけ書き、月額や金額を併記しない。金額値引きの場合は値引き後の月額のみを書く。例:「初月無料」または「月額25万円」）、5行目（任意）=下部に表示する注記（例: 初期費用無料。不要なら書かない）。バッジ（1行目）には初期費用を書かない。サービス名・概要・プラン名はNo.96から自動反映。JSONに taikei フィールドを必ず含める。', 20, 6);
  set([104], '3行・各10字以内（例:「分割払いも可能です。」「24回払いなら」「〇万円／月」）。', 10, 3);
  set([QA_NO], 'よくある質問（Q&A）。ヒアリングの「⒌ よくある質問（Q&A）」で記入がある組だけを、1組につき「Q. 質問」1行＋「A. 答え」1〜2行で並べる（最大3組・シートの順番どおり）。質問・答えはシートの文言をそのまま使い、答えが長いときだけ意味を変えずに縮める。Q&Aに無い内容を足さない。記入が無い案件は空文字 ""。', 32, 9);
  set([6], '1行目=図解タグ、2行目=会社名・肩書、3行目=氏名、4行目以降=プロフィール文。', 200, 12);
  set([1], '「図解（サムネイル）」固定。', 30, 1);
  return S;
}
function range_(a, b) { const r = []; for (let i = a; i <= b; i++) r.push(i); return r; }

// ============================================================
// 106行定義テーブル
//  km: GEN/FIX/REF/EMPTY, dm: GEN/FIX, layout, img: 図解キーワード, icon: アイコン枠あり
// ============================================================
function rowDefs_() {
  const R = [];
  function d(no, tab, label, sub, km, kFix, kRef, dm, layout, img, icon) {
    R.push({ no: no, tab: tab, label: label, sub: sub || '', km: km, kFix: kFix || '', kRef: kRef || 0, dm: dm, layout: layout, img: img || null, icon: !!icon });
  }
  const T = TABS, L = LAYOUT;
  d(1, T.INTRO, '表紙（タイトル）', '', 'FIX', '図解（サムネイル）', 0, 'GEN', L.COVER, 'サムネイル');
  d(2, T.INTRO, '問いかけ', '', 'FIX', 'もし、こうなったら？', 0, 'FIX', L.QUESTION, null);
  d(3, T.INTRO, '理想①②③', '', 'GEN', '', 0, 'GEN', L.IDEAL, null);
  d(4, T.INTRO, '', '', 'GEN', '', 0, 'GEN', L.IDEAL, null);
  d(5, T.INTRO, '', '', 'GEN', '', 0, 'GEN', L.IDEAL, null);
  d(6, T.INTRO, '講師紹介', '', 'GEN', '', 0, 'GEN', L.PROFILE, '*プロフ');
  d(7, T.PROBLEM, '問いかけ', '', 'FIX', '例えば、こんなお悩みありませんか？', 0, 'FIX', L.QUESTION, null);
  d(8, T.PROBLEM, 'お悩み①②③', '', 'GEN', '', 0, 'GEN', L.IDEAL, null);
  d(9, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.IDEAL, null);
  d(10, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.IDEAL, null);
  d(11, T.PROBLEM, 'ある経営者の話', '', 'GEN', '', 0, 'GEN', L.IDEAL, null);
  d(12, T.PROBLEM, '以前の状況', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(13, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(14, T.PROBLEM, '課題のきっかけ', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(15, T.PROBLEM, '引き起こる課題（一次）', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(16, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(17, T.PROBLEM, '引き起こる課題（二次）', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(18, T.PROBLEM, '連鎖・拡大', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(19, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(20, T.PROBLEM, '経営リスクの認識', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(21, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(22, T.PROBLEM, '解決への決断', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(23, T.PROBLEM, '取り組み①の内容', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(24, T.PROBLEM, '取り組み①のエピソード', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(25, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(26, T.PROBLEM, '取り組み①の結果', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(27, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(28, T.PROBLEM, '取り組み②の内容', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(29, T.PROBLEM, '取り組み②のエピソード', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(30, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(31, T.PROBLEM, '取り組み②の結果', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(32, T.PROBLEM, '経営者の嘆き', '', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(33, T.PROBLEM, '問いかけ', '', 'FIX', 'なぜ、このような問題が起こるのか？', 0, 'FIX', L.QUESTION, null);
  d(34, T.PROBLEM, '旧ビリーフ', '', 'GEN', '', 0, 'GEN', L.TEXT55, null);
  d(35, T.PROBLEM, '旧ビリーフの定義', '', 'GEN', '', 0, 'GEN', L.DIAGRAM, '図解1');
  d(36, T.PROBLEM, 'デメリット', '', 'GEN', '', 0, 'GEN', L.DIAGRAM, '図解2');
  d(37, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.DIAGRAM, '図解3');
  d(38, T.PROBLEM, '問いかけ', '', 'FIX', 'では、どうすれば？', 0, 'FIX', L.QUESTION, null);
  d(39, T.PROBLEM, '新ビリーフ', '', 'GEN', '', 0, 'GEN', L.TEXT55, null);
  d(40, T.PROBLEM, '新ビリーフの定義', '', 'GEN', '', 0, 'GEN', L.DIAGRAM, '図解4');
  d(41, T.PROBLEM, 'メリット', '', 'GEN', '', 0, 'GEN', L.MERIT_INTRO, null);
  d(42, T.PROBLEM, 'メリット①', '', 'GEN', '', 0, 'GEN', L.MERIT, null, true);
  d(43, T.PROBLEM, 'メリット②', '', 'GEN', '', 0, 'GEN', L.MERIT, null, true);
  d(44, T.PROBLEM, 'メリット③', '', 'GEN', '', 0, 'GEN', L.MERIT, null, true);
  d(45, T.PROBLEM, '事例（Before）', '', 'GEN', '', 0, 'GEN', L.CASE, '事例①（before）');
  d(46, T.PROBLEM, '事例（After）', '', 'GEN', '', 0, 'GEN', L.CASE, '事例①（after）');
  d(47, T.PROBLEM, 'よくある間違い', '', 'GEN', '', 0, 'GEN', L.TEXT55, null);
  d(48, T.PROBLEM, '', '', 'GEN', '', 0, 'GEN', L.TEXT55, null);
  d(49, T.PROBLEM, '問いかけ', '', 'FIX', 'では、どうすればいいのか？', 0, 'FIX', L.QUESTION, null);
  d(50, T.PROBLEM, '正しい行動のポイント', '', 'GEN', '', 0, 'GEN', L.POINT_INTRO, null);
  d(51, T.PROBLEM, 'ポイント①', '', 'GEN', '', 0, 'GEN', L.POINT1, null, true);
  d(52, T.PROBLEM, 'ポイント①の解説', '', 'GEN', '', 0, 'GEN', L.DIAGRAM, '図解5');
  d(53, T.PROBLEM, 'ポイント②', '', 'GEN', '', 0, 'GEN', L.POINT2, null, true);
  d(54, T.PROBLEM, 'ポイント②の解説', '', 'GEN', '', 0, 'GEN', L.DIAGRAM, '図解6');
  d(55, T.PROBLEM, 'よくある課題', '', 'GEN', '', 0, 'GEN', L.CHALLENGE, null);
  d(56, T.PROBLEM, 'サービス表紙\nサービスを提供している理由', '', 'GEN', '', 0, 'GEN', L.SERVICE, 'モックアップ');
  d(57, T.STORY, '⒈ 等身大のスタート（原点）', 'キャリアスタート', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(58, T.STORY, '', 'きっかけ（原点）', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(59, T.STORY, '', '小さな成功体験', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(60, T.STORY, '⒉ 挫折と誤った努力（苦労）', '間違った努力と失敗①', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(61, T.STORY, '', '間違った努力と失敗②', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(62, T.STORY, '', '間違った努力と失敗③', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(63, T.STORY, '', '間違った努力と失敗④', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(64, T.STORY, '⒊ 勝ちパターンとの出会い（転機）', '◯◯との出会い（見つける）', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(65, T.STORY, '⒋ 検証から初期成果（小さな成功）', '検証過程（取り組んだこと）', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(66, T.STORY, '', '検証過程（取り組んだこと）', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(67, T.STORY, '', '初期成果', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(68, T.STORY, '⒌ 型化（再現性）', '成功体験の型化（取り組んだこと）', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(69, T.STORY, '', '成功体験の型化（取り組んだこと）', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(70, T.STORY, '', '成功体験の型化（取り組んだこと）', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(71, T.STORY, '', '再現性（自分以外も成果が出た）', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(72, T.STORY, '', 'サービス化（開発した）', 'GEN', '', 0, 'GEN', L.NARRATIVE, null);
  d(73, T.STORY, '⒍ 他社での成功と理念', '事例① Before', 'GEN', '', 0, 'GEN', L.CASE, '事例②（before）');
  d(74, T.STORY, '', '事例① After', 'GEN', '', 0, 'GEN', L.CASE, '事例②（after）');
  d(75, T.STORY, '', '事例② Before', 'GEN', '', 0, 'GEN', L.CASE, '事例③（before）');
  d(76, T.STORY, '', '事例② After', 'GEN', '', 0, 'GEN', L.CASE, '事例③（after）');
  d(77, T.STORY, '', '成果再現済み', 'GEN', '', 0, 'GEN', L.IMPACT, null);
  d(78, T.STORY, '', '理念（想い）', 'GEN', '', 0, 'GEN', L.RINEN, null);
  d(79, T.OFFER, 'サービス表紙', '', 'GEN', '', 0, 'GEN', L.SERVICE, 'モックアップ');
  d(80, T.OFFER, '特長①②③', '', 'GEN', '', 0, 'GEN', L.FEATURES, null, true);
  d(81, T.OFFER, 'サンプル①', '', 'GEN', '', 0, 'GEN', L.DIAGRAM40, '特徴①');
  d(82, T.OFFER, '特長①②③', '', 'REF', '', 80, 'GEN', L.FEATURES, null);
  d(83, T.OFFER, 'サンプル②', '', 'GEN', '', 0, 'GEN', L.DIAGRAM40, '特徴②');
  d(84, T.OFFER, '特長①②③', '', 'REF', '', 80, 'GEN', L.FEATURES, null);
  d(85, T.OFFER, 'サンプル③', '', 'GEN', '', 0, 'GEN', L.DIAGRAM40, '特徴③');
  d(86, T.OFFER, '利用イメージ（全体像）', '', 'GEN', '', 0, 'GEN', L.STEPS, null);
  d(87, T.OFFER, '工程①', '', 'FIX', '図解（工程①）', 0, 'GEN', L.KOTEI, '工程①', true);
  d(88, T.OFFER, '工程②', '', 'FIX', '図解（工程②）', 0, 'GEN', L.KOTEI, '工程②', true);
  d(89, T.OFFER, '工程③', '', 'FIX', '図解（工程③）', 0, 'GEN', L.KOTEI, '工程③', true);
  d(90, T.OFFER, '工程④', '', 'FIX', '図解（工程④）', 0, 'GEN', L.KOTEI, '工程④', true);
  d(91, T.OFFER, '成果実証済み', '', 'GEN', '', 0, 'GEN', L.IMPACT, null);
  d(92, T.OFFER, '開発費用と期間', '', 'GEN', '', 0, 'GEN', L.INVEST, null);
  d(93, T.OFFER, '理想①②', '', 'REF', '', 3, 'GEN', L.IDEAL, null);
  d(94, T.OFFER, '', '', 'REF', '', 4, 'GEN', L.IDEAL, null);
  d(95, T.OFFER, '価値ある？→ 取り消し線', '', 'FIX', 'No.92の金額に赤の取り消し線を引く。', 0, 'GEN', L.STRIKE, null);
  d(96, T.OFFER, '定価', '', 'GEN', '', 0, 'GEN', L.PRICE, null);
  d(97, T.OFFER, 'ウェビナー終了までなら', '', 'EMPTY', '', 0, 'FIX', L.S_ZOOM, null);
  d(98, T.OFFER, '特典⒈', '', 'GEN', '', 0, 'GEN', L.DIAGRAM40, '特典①');
  d(99, T.OFFER, '特典⒉', '', 'GEN', '', 0, 'GEN', L.DIAGRAM40, '特典②');
  d(100, T.OFFER, '理想③', '', 'GEN', '', 0, 'GEN', L.IDEAL, null);
  d(101, T.OFFER, '名言', '', 'EMPTY', '', 0, 'FIX', L.S_MEIGEN, null);
  d(102, T.OFFER, '値引き後の金額', '', 'GEN', '', 0, 'GEN', L.PRICE, null);
  d(103, T.OFFER, 'なぜこんな価格で？', '', 'GEN', '', 0, 'GEN', L.RINEN, null);
  d(104, T.OFFER, '分割払いも可', '', 'GEN', '', 0, 'GEN', L.TEXT55, null);
  d(105, T.OFFER, 'BYAF法', '', 'EMPTY', '', 0, 'FIX', L.S_BYAF, null);
  d(106, T.OFFER, '御礼', '', 'EMPTY', '', 0, 'FIX', L.S_OREI, null);
  return R;
}

/** 台本をAI生成する行のNo一覧（分割単位つき） */
function daihonGenRows_() {
  const defs = rowDefs_().filter(function (r) { return r.dm === 'GEN'; });
  return {
    partA: defs.filter(function (r) { return r.no <= 56; }).map(function (r) { return r.no; }),
    partB: defs.filter(function (r) { return r.no >= 57; }).map(function (r) { return r.no; })
  };
}

/** 構成テキストをAI生成する行のNo一覧 */
function kouseiGenRows_() {
  return rowDefs_().filter(function (r) { return r.km === 'GEN'; }).map(function (r) { return r.no; });
}

/** 構成テキスト中の図解タグを必須とする行 → タグ */
function requiredImgTags_() {
  const m = {};
  rowDefs_().forEach(function (r) {
    if (r.km === 'GEN' && r.img && r.img !== '*プロフ') m[r.no] = '図解（' + r.img + '）';
  });
  return m;
}

/** アイコン差し替え対象の行No一覧 */
function iconRows_() {
  return rowDefs_().filter(function (r) { return r.icon; }).map(function (r) { return r.no; });
}

// ============================================================
// 案件フォルダ（【出力】配下に会社名フォルダを作り、案件の成果物を集約）
// ============================================================
function caseCompany_(projectName) {
  const nm = String(projectName || '').replace(/[\s_＿\-]*ヒアリングシート\s*$/, '').trim();
  return nm || String(projectName || '');
}
// ============================================================
// 名前でフォルダ・ファイルを探す（濁点の分解に強い版）
// ============================================================
/**
 * ★Macの「Googleドライブ（パソコン版）」から入れたフォルダ・ファイルは、
 *   名前の濁点が分解された形（NFD。「グ」＝「ク」＋「゛」）でドライブに保存される。
 *   GASの getFoldersByName / getFilesByName は文字列の完全一致なので、
 *   合成形（NFC）で問い合わせると、そういう名前は永久に見つからない。
 *
 *   実害（2026-08-13 柳澤社）：素材フォルダ「柳澤社_ヒアリングシート」が見つからず、
 *   図解が1枚も入らないまま、表紙がテンプレート（別会社のロゴ）のスライドができた。
 *   エラーは出ない。管理シートF列に「図解未発見」が並ぶだけなので、見落とすと客先へ出る。
 *
 *   そこで、完全一致で見つからなかったときだけ、中を総なめしてNFCで比べ直す。
 *   戻り値は getFoldersByName と同じ使い勝手（hasNext / next）にしてある。
 */
function nfc_(s) {
  const t = String(s == null ? '' : s);
  return t.normalize ? t.normalize('NFC') : t;
}

function iterByName_(direct, listAll, name) {
  if (direct.hasNext()) return direct;
  const want = nfc_(name);
  const hits = [];
  const all = listAll();
  while (all.hasNext()) {
    const f = all.next();
    if (nfc_(f.getName()) === want) hits.push(f);
  }
  let i = 0;
  return {
    hasNext: function () { return i < hits.length; },
    next: function () { return hits[i++]; }
  };
}

function foldersByName_(parent, name) {
  return iterByName_(parent.getFoldersByName(name), function () { return parent.getFolders(); }, name);
}

function filesByName_(parent, name) {
  return iterByName_(parent.getFilesByName(name), function () { return parent.getFiles(); }, name);
}

/**
 * 案件フォルダを「探すだけ」。無ければ null を返し、作らない。
 * ★読み取り側は必ずこちらを使うこと。caseFolder_ は無いと作ってしまうので、
 *   案件フォルダを消しても5分後に空フォルダが復活し続ける（2026-08-12 実測）。
 *   管理シートに行が残っている限り永久に生えるので、消したつもりが消せない。
 */
function caseFolderIfExists_(cfg, projectName) {
  const parent = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  const it = foldersByName_(parent, caseCompany_(projectName));
  return it.hasNext() ? it.next() : null;
}

function caseFolder_(cfg, projectName) {
  const parent = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  const nm = caseCompany_(projectName);
  const it = foldersByName_(parent, nm);
  return it.hasNext() ? it.next() : parent.createFolder(nm);
}

// ============================================================
// 構成シートのセル（B3・B15・B23など）に出す失敗の文
//
// ★GASのエラーは英語のまま（"Exception: No item with the given ID..." など）で、
//   PMが読んでも何が起きたか・次に何をすればよいか分からなかった。
//   「何が起きたか」「次にすること」を日本語で先に書き、元の文は最後に（詳細）で添える。
//   例：失敗の文_('LP修正', e, 'B10') →
//     'LP修正: 失敗しました。ファイルかフォルダが見つかりません…。…もう一度B10にチェックを入れてください（詳細: …）'
// ============================================================
function 失敗の文_(見出し, e, 押し直すセル) {
  const 元 = String(e && e.message ? e.message : e).replace(/^(Exception|Error|例外):\s*/i, '');
  // 「B10」のようなセル名ならチェックの入れ直し、文ならそのまま使う
  const 押し直す = !押し直すセル ? 'もう一度同じ操作をしてください'
    : /^[A-Z]+\d+$/.test(押し直すセル) ? 'もう一度' + 押し直すセル + 'にチェックを入れてください'
    : 押し直すセル;
  const 相談 = '同じ文が2回続いたら、この欄の文をそのままDiscordの現場相談へ貼ってください';
  let 何 = '', 次 = '';
  // 仕組みが自分で投げた日本語のエラー（例：台本JSONがありません）は、そのままで読める。
  // GAS自身のエラーは「Exception: 」で始まるので、それ以外で日本語なら自前のものとみなす
  const 自前 = !/^Exception/i.test(String(e)) && /[぀-ヿ一-鿿]/.test(元);
  if (自前) {
    return 見出し + ': 失敗しました。' + 元 + '。書かれていることを直してから、' + 押し直す + '。' + 相談;
  }
  if (/no item with the given id|not found|could not be found|does not exist|was deleted|見つかりません|ありません|存在しません/i.test(元)) {
    何 = 'ファイルかフォルダが見つかりません（消された・名前が変わった・移動された・共有が外れた可能性があります）';
    次 = 'フォルダ名とファイル名が「◯◯社」で合っているか見てから、' + 押し直す + '。' + 相談;
  } else if (/permission|access denied|not have access|forbidden|権限|アクセス/i.test(元)) {
    何 = 'ファイルかフォルダを開く権限がありません';
    次 = 'そのファイルやフォルダが共有ドライブの中にあり、サポートアカウントが開けるか見てから、' + 押し直す + '。' + 相談;
  } else if (/exceeded maximum execution time|timed? ?out|timeout|時間切れ/i.test(元)) {
    何 = '処理に時間がかかりすぎて、途中で止まりました';
    次 = '5分ほど待ってから、' + 押し直す + '。' + 相談;
  } else if (/too many times|service invoked|rate limit|quota|service .* failed|temporar|internal error|we're sorry|server error|backend error|\block\b|混み合|一時的/i.test(元)) {
    何 = 'Googleの側が混み合っていました（入力の誤りではありません）';
    次 = '5分ほど待ってから、' + 押し直す + '。' + 相談;
  } else {
    何 = '想定していないエラーが起きました';
    次 = 'この欄の文をそのままDiscordの現場相談へ貼ってください';
  }
  return 見出し + ': 失敗しました。' + 何 + '。' + 次 + '（詳細: ' + 元 + '）';
}
