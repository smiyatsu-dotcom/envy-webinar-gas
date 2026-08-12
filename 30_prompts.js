/**
 * 30_prompts.gs（APIレス版）
 * ヒアリングシート読込と、手動実行用プロンプト（Googleドキュメント）の構築
 *
 * フロー: GASがヒアリング内容を埋め込んだプロンプトDocを自動生成
 *  → ユーザーがClaude（チャット/Claude Code）に貼って実行
 *  → 返ってきた2つのJSON（台本→構成テキスト）を構成シート「AI出力貼り付け」タブへ貼る
 */

// ============================================================
// ヒアリングシートの読込・テキスト化
// ============================================================

/** シート1枚をタブ区切りテキストへ（空行圧縮） */
function serializeSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  const lines = [];
  values.forEach(function (row) {
    const cells = row.map(function (c) { return String(c).trim(); }).filter(function (c) { return c !== ''; });
    if (cells.length > 0) lines.push(cells.join('\t'));
  });
  return lines.join('\n');
}

/** ヒアリングスプシ（3タブ）を読み込む。タブ名は前後空白を許容 */
function readHearing_(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  function findTab(name) {
    const hit = ss.getSheets().filter(function (s) { return s.getName().trim() === name.trim(); })[0];
    if (!hit) throw new Error('ヒアリングシートにタブが見つかりません: ' + name);
    return hit;
  }
  return {
    offer: serializeSheet_(findTab(HEARING_TABS.OFFER)),
    problem: serializeSheet_(findTab(HEARING_TABS.PROBLEM)),
    story: serializeSheet_(findTab(HEARING_TABS.STORY))
  };
}

// ============================================================
// プロンプト本体
// ============================================================

function rolePrompt_() {
  return [
    'あなたはウェビナー（オンラインセミナー）の構成作家です。',
    'クライアントへのヒアリング内容をもとに、決められた固定構造（No.1〜106）に沿ってウェビナーの原稿を作成します。',
    '',
    '# ウェビナーの全体構造（固定）',
    '- イントロ（No.1〜6）: 表紙 → 問いかけ → 理想①②③ → 講師紹介',
    '- 問題教育（No.7〜56）: お悩み提示 → モデル顧客の失敗ナラティブ → 旧ビリーフ（原因）→ 新ビリーフ（解決策）→ メリット3点 → 事例Before/After → よくある間違い → 正しいポイント2点 → サービス表紙',
    '- ストーリー（No.57〜78）: 講師自身の起業物語（原点 → 挫折 → 転機 → 検証 → 型化 → 他社成功と理念）',
    '- オファー（No.79〜106）: サービス詳細 → 実績 → 価格 → 特典 → クロージング',
    '',
    '# 絶対に守るルール',
    '1. 内容の創作をしない。すべての内容はヒアリング記入内容から導出すること。ヒアリングにない事実・エピソード・固有名詞を追加しない。',
    '2. 金額・社数・年数などの数値は、ヒアリングに記載された値だけを使用する。数値の変更・推定・創作は禁止。',
    '3. 文体変換と要約・分割は行ってよいが、意味を変えないこと。',
    '4. 一貫性ルール:',
    '   - 理想①②③（No.3〜5）＝ 新ビリーフのメリット3点（No.42〜44）を「もし、〜たら？」形式に変換したもの。',
    '   - お悩み①②③（No.8〜10）＝ 理想①②③の裏返し（現状の悩み）。',
    '   - 旧ビリーフ・新ビリーフはヒアリング記入値をそのまま核として使用。',
    '   - 理念（No.78）の中核フレーズを No.103 で再引用する。',
    '   - サービス概要文（サービス名＋一言説明）は No.56・79・96 で同一の文を使用する。',
    '5. 料金体系の判定: ヒアリング【オファー記入用】の料金体系欄（単一用/松竹梅用の記入欄）から「単一」か「松竹梅」かを判定する。No.96とNo.102のJSONオブジェクトには "taikei" フィールド（"単一" または "松竹梅"）を必ず含めること。松竹梅は常に3プラン（ライト/スタンダード/プレミアム）。'
  ].join('\n');
}

function daihonRulesPartA_() {
  return [
    '## イントロ',
    '- No.1: 挨拶＋自己紹介＋ウェビナータイトルの読み上げ。タイトルは「ターゲットの悩み＋理想状態」を組み合わせて生成（例:「社員20名を超えて現場対応に追われる社長の任せる組織づくり」）。',
    '- No.3〜5: 理想①②③。「もし、〜になったらどうでしょうか？」の問いかけ。メリット3点（ヒアリング⒊質問6）から逆算する。',
    '- No.6: 「その秘訣を本日シェアさせていただきます。」から始め、プロフィール（基本情報）を一人称で語り直し、「では、私の紹介は以上にして、具体的な中身に入っていこうと思います。」で締める。',
    '## 問題教育（ヒアリング【問題教育の記入用】から）',
    '- No.8〜10: お悩み①②③。「まず1つ目、〜」「続いて2つ目、〜」「最後3つ目、〜」で始め「というお悩みはありませんか。」＋「（挙手ありがとうございます。）」で締める。',
    '- No.11: モデル顧客の導入。「実は、このような3つの悩みを抱えていた、ある〇〇会社の経営者がいました。それが〇〇社長です。」',
    '- No.12〜13: ⒈質問1（以前の状況）を2文脈に分割。',
    '- No.14: ⒈質問2（課題のきっかけ）。',
    '- No.15〜16: ⒈質問3（一次課題）を2分割。',
    '- No.17〜19: ⒈質問4（二次課題）を「二次課題→連鎖→拡大」の3行に分割。',
    '- No.20〜21: ⒈質問5（経営リスクの認識）を「具体的リスク→本質的な恐れ」の2行に分割。',
    '- No.22: ⒈質問6（解決への決断）。「〜することを決意しました。」調。',
    '- No.23〜27: ⒉取り組み①を「内容(23)→エピソード(24,25)→結果=失敗(26,27)」の5行に分割。',
    '- No.28〜31: ⒉取り組み②を「内容(28)→エピソード(29,30)→結果=失敗(31)」の4行に分割。',
    '- No.32: ⒉質問7の言葉をカギ括弧の直接話法で。経営者の嘆き。',
    '- No.34: 「それは、（旧ビリーフ）だからです。」',
    '- No.35: ⒊質問2（旧ビリーフの定義）の説明。',
    '- No.36〜37: ⒊質問3（デメリット）を2行に分割。「この状態が続くと〜」「さらに〜」',
    '- No.39: 「（新ビリーフ）することです。」',
    '- No.40: ⒊質問5（新ビリーフの定義）。誤解の否定（「〜ではありません」）があれば含める。',
    '- No.41: 「そんな状態、（新ビリーフの動作）することで得られるメリットは3つあります。」',
    '- No.42〜44: メリット①②③。「まず1つ目は〜」「2つ目は〜」「最後3つ目は〜」',
    '- No.45: ⒊質問7のBeforeを「例えば、冒頭でお伝えした〇〇社長は、もともとは〜」で要約（モデル顧客の業種と社長名はNo.11と一致させる）。',
    '- No.46: Afterの要約＋「そんな事例があります。」で締める。',
    '- No.47: 「ただ、『（新ビリーフ）』を理解して実践し始めた経営者が、よく陥ってしまう間違いがあります。」',
    '- No.48: ⒋質問1〜2（よくある間違いの内容と定義）。「それが、（間違い）です。」から始める。',
    '- No.50: 「正しく（テーマの動作）するためのポイントは2つです。」',
    '- No.51: 「まず1つ目は、（ポイント①）です。」 No.52: ポイント①の解説（⒋質問4前半）。',
    '- No.53: 「2つ目は、（ポイント②）です。」 No.54: ポイント②の解説（⒋質問4後半）。',
    '- No.55: 「とはいえ、実際に取り組み始めると、次のような課題が出てきます。」＋自力実践の課題3点（ヒアリングから導出）＋「その解決策が（↓）」',
    '- No.56: 「『（サービス名）』です。」＋サービス概要文（オファー記入用の概要）。'
  ].join('\n');
}

function daihonRulesPartB_() {
  return [
    '## ストーリー（ヒアリング【ストーリーの記入用】から。講師の一人称で語る）',
    '- No.57〜59: ⒈の3記入欄（キャリアスタート／きっかけ（原点）／小さな成功体験）をそれぞれ1行に。',
    '- No.60〜63: ⒉の失敗①〜④をそれぞれ1行に。',
    '- No.64: ⒊転機（出会い）。',
    '- No.65〜66: ⒋検証過程①②。 No.67: ⒋初期成果。',
    '- No.68〜70: ⒌型化①〜③。 No.71: ⒌再現性。 No.72: ⒌サービス化。「それが（サービス名）」に繋がる形。',
    '- No.73〜74: ⒍事例① Before/After。「ここからは、実際に導入した企業の事例をご紹介します。」等の導入から。Afterは「そんな事例があります。」で締める。',
    '- No.75〜76: ⒍事例② Before/After。「他には、〜」から。Afterは「そんな事例もあります。」で締める。',
    '- No.77: ⒍成果再現済み（導入実績数はヒアリング値のみ）。',
    '- No.78: ⒍理念（想い）。中核フレーズをカギ括弧で明示する。',
    '## オファー（ヒアリング【オファー記入用】から）',
    '- No.79: 「それが『（サービス名）』です。」＋サービス概要文（No.56と同一の概要文を使用）。',
    '- No.80: 「（サービス名）には、大きく3つの特徴があります。」＋特徴3点を順に一言ずつ紹介。',
    '- No.81: 「まず1つ目の特徴は、〜です。」＋特徴①の詳細（質問2）。',
    '- No.82: 「2つ目の特徴は、〜です。」（1文のみ）。',
    '- No.83: 特徴②の詳細（質問3）。',
    '- No.84: 「最後の特徴は、〜です。」（1文のみ）。',
    '- No.85: 特徴③の詳細（質問4）。',
    '- No.86: 「そんな（サービス名）は、大きく4つのステップで進めていきます。」＋工程①〜④の列挙＋「それでは、それぞれ詳しくご紹介します。」',
    '- No.87〜90: 各工程の詳細説明（提供プロセスの質問2〜5）。「まず最初に行うのが〜」「続いて行うのが〜」等。',
    '- No.91: 導入実績（料金体系の質問1の値のみ）。',
    '- No.92: 「この（サービス名）を開発するまでに、約〇年の期間をかけ、開発費用は約〇〇円を投資しました。このくらいの時間と費用をかけても、」（ヒアリング値のみ）。',
    '- No.93〜94: 理想①②の再掲。No.93の冒頭は No.92 の「このくらいの時間と費用をかけても、」から自然に続く問いかけにする。',
    '- No.95: 「〇〇円の価値はあるでしょうか。しかし、（サービス名）は〇〇円ではありません。」（〇〇=No.92の投資額）。',
    '- No.96: 定価の提示（料金体系の質問4。複数プランがあれば全て）。',
    '- No.98: 特典1の紹介（料金体系の質問6）。 No.99: 特典2の紹介。**特典が1つしかない場合は台本・構成テキストとも空文字 "" を出力**（スライドは斜線=スキップ扱いになる）。',
    '- No.100: 理想③を「30日後、〜状態になったら、空いた時間で何に挑戦しますか？」形式に。',
    '- No.102: ウェビナー特別価格の提示（質問5）。「本日Zoom終了までなら〜」',
    '- No.103: 「なぜこんな価格でご提供しているのかと申しますと〜」＋理念（No.78の中核フレーズ）を理由として引用。',
    '- No.104: 分割払い案。定価から計算できる場合のみ（例: 24回払いで月々〇円）。計算できない料金体系（月額制など）なら台本・構成テキストとも空文字 "" を出力（スライドは斜線=スキップ扱い）。'
  ].join('\n');
}

function kouseiRules_() {
  const tags = requiredImgTags_();
  const spec = kouseiSpec_();
  // 行別の書式・字数規定（00_config.gs kouseiSpec_ が単一ソース）
  const nos = kouseiGenRows_();
  const grouped = {}; // fmt文 → no配列（同一規定をまとめて提示）
  nos.forEach(function (no) {
    const s = spec[no];
    if (!s) return;
    (grouped[s.fmt] = grouped[s.fmt] || []).push(no);
  });
  const ruleLines = Object.keys(grouped).map(function (fmt) {
    return '- No.' + grouped[fmt].join(',') + ': ' + fmt;
  });
  const tagLines = nos.filter(function (no) { return tags[no]; })
    .map(function (no) { return '- No.' + no + ': 末尾に「' + tags[no] + '」を必ず含める'; });
  return [
    '# 構成テキストの文体・書式（サンプル実測に基づく厳守ルール）',
    '- スライド見出しコピー。その行の台本の要点を圧縮したもの。台本にない内容を足さない。',
    '- 各行の書式と字数制限は以下のとおり。制限を超える場合は言い換えて短くすること。',
    ruleLines.join('\n'),
    '',
    '# 図解タグ（必須）',
    tagLines.join('\n')
  ].join('\n');
}

/** アイコン・背景の選定ルール（4フォルダのファイル名一覧を埋め込む） */
function designRules_(lists) {
  const t = designTargets_();
  const has = lists.iconP.length + lists.iconW.length + lists.bg85.length + lists.bg25.length + lists.bg15.length + lists.select.length > 0;
  if (!has) return '';
  const parts = ['## 段階3: アイコン・背景の選定',
    '各行の内容・雰囲気に最も合う素材ファイルを一覧から選ぶ。ぴったりのものが無ければ空文字 ""。'];
  if (lists.iconP.length > 0) {
    parts.push('',
      '### アイコン（紫）: 対象行 ' + t.iconPurple.join(','),
      '- No.42-44=各メリットの内容、No.51/53=各ポイントの内容、No.87-90=各工程の作業内容を象徴するアイコンを1つずつ（icon欄）。',
      '- 一覧: ' + lists.iconP.map(function(n){return assetLabel_('ICON_P', n);}).join(' / '));
  }
  if (lists.iconW.length > 0) {
    parts.push('',
      '### アイコン（白）: 対象行 80 のみ',
      '- 特徴①②③それぞれに合うアイコンを3つ選び、icon欄に「名1／名2／名3」の全角スラッシュ区切りで出力（No.82/84へは自動再掲）。',
      '- 一覧: ' + lists.iconW.map(function(n){return assetLabel_('ICON_W', n);}).join(' / '));
  }
  if (lists.bg85.length + lists.bg25.length + lists.bg15.length > 0) {
    parts.push('',
      '### 背景画像（No.別に選定元の濃度フォルダが決まっている。必ず該当濃度の一覧から選ぶこと）',
      '- 行の感情トーンに合わせて選び、bg欄にファイル名を出力。同じ背景の連続使用可。',
      '- 濃度85%の対象行: ' + BG_OPACITY_MAP['85'].join(','),
      '  一覧: ' + lists.bg85.map(function(n){return assetLabel_('BG', n);}).join(' / '),
      '- 濃度25%の対象行: ' + BG_OPACITY_MAP['25'].join(','),
      '  一覧: ' + lists.bg25.map(function(n){return assetLabel_('BG', n);}).join(' / '),
      '- 濃度15%の対象行: ' + BG_OPACITY_MAP['15'].join(','),
      '  一覧: ' + lists.bg15.map(function(n){return assetLabel_('BG', n);}).join(' / '));
  }
  if (lists.select.length > 0) {
    parts.push('',
      '### 選択画像: 対象行 72 のみ（bg欄）',
      '- 一覧: ' + lists.select.map(function(n){return assetLabel_('SELECT', n);}).join(' / '));
  }
  parts.push('',
    '【出力3: デザインJSON】対象行のみでよい',
    '```json',
    '[{"no": 3, "icon": "", "bg": "ファイル名"}, {"no": 42, "icon": "ファイル名", "bg": ""}, ...]',
    '```');
  return parts.join('\n');
}

/** 用途キーのフォルダ内 画像ファイル名一覧 */
function scanAssetNames_(cfg, key) {
  const names = [];
  const folder = resolveNamedFolder_(cfg, key); // 50_slides.gs
  if (!folder) return names;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType().indexOf('image/') === 0) names.push(f.getName());
  }
  return names.sort();
}

/**
 * 手動実行用プロンプト全文を組み立てる（プロンプトDocの中身）
 */
function buildManualPrompt_(hearing, assetLists, projectName) {
  const gen = daihonGenRows_();
  const allDaihonNos = gen.partA.concat(gen.partB);
  const kouseiNos = kouseiGenRows_();
  const design = designRules_(assetLists || { iconP: [], iconW: [], bg: [], select: [] });
  const company = projectName ? caseCompany_(projectName) : '｛会社名｝';
  const blocks = [
    '# 【プロンプト】' + company + ' — Claude Codeへの作業指示',
    '',
    'あなた(Claude Code)は、作業フォルダとして共有ドライブ「社内｜Claude Code資料生成」が選択された状態で、この指示を受け取っています。次の手順で作業してください。',
    '',
    '1. 「【固定】ウェビナー自動生成/資料改善メモ.md」を読み込み、そこに書かれた各ルールを、下記の生成ルールと同格の追加ルールとして必ず適用する（ファイルが無い場合はそのまま進める）',
    '2. 下記のルールに従って段階1→2' + (design ? '→3' : '') + 'の順に生成し、末尾の「出力形式」のとおりtxtファイルへ保存する',
    '3. 利用者が添削・修正・要望を伝えたら、必ず次の2つを行う: ①該当のtxtを修正して上書き保存する ②その添削を「【固定】ウェビナー自動生成/【受信箱】添削ログ.md」の末尾に記録する。形式は「## 日付 案件名 / No.◯◯（対象）」の見出しに続けて「- 指示:」「- 対応:」「- 状態: 未処理」の3行。ルールにするかどうかは後で人が仕分けるので、ここでは判断せず、そのまま記録する',
    '   ※「資料改善メモ.md」は配布用のコピーなので直接編集しないこと。ルールとして追加されるのは、仕分けと承認を経たあと。',
    '4. ファイルの保存・上書き・削除は、必ず作業フォルダ(ローカル同期された「社内｜Claude Code資料生成」)のファイル操作で行うこと。Google Driveコネクタ(API)でのファイル作成は禁止(削除・上書きができず重複の原因になるため)。作業フォルダが選択されていない場合は、作業を始める前に利用者へ「作業フォルダに社内｜Claude Code資料生成を選択してください」と伝えて中断すること',
    '',
    rolePrompt_(),
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━',
    '# ヒアリング【オファー記入用】',
    hearing.offer,
    '',
    '# ヒアリング【問題教育の記入用】',
    hearing.problem,
    '',
    '# ヒアリング【ストーリーの記入用】',
    hearing.story,
    '━━━━━━━━━━━━━━━━━━━━━━━━━━',
    '',
    '# あなたのタスク（' + (design ? '3' : '2') + '段階）',
    '',
    '## 段階1: 台本の作成',
    '講師が話すスピーカーノート「台本」を作成する。',
    '- 文体: 「です・ます」の語り口調。聞き手（経営者）に語りかける自然な話し言葉。1行（=スライド1枚）あたり2〜6文。',
    '- 分量: 台本全体で約13,000〜16,000字（読み上げ50〜55分が目安）。問いかけ・強調・言い換え・つなぎの言葉で講演らしく肉付けする。ただしヒアリングにない事実・数値・エピソードの追加は禁止。',
    '- 生成対象行: ' + allDaihonNos.join(','),
    '',
    '### 行別ルール（No.1〜56）',
    daihonRulesPartA_(),
    '',
    '### 行別ルール（No.57〜106）',
    daihonRulesPartB_(),
    '',
    '## 段階2: 構成テキストの作成',
    '段階1で作成した台本を圧縮し、各スライドに表示する「構成テキスト」を作成する。',
    '- 生成対象行: ' + kouseiNos.join(','),
    '',
    kouseiRules_()
  ];
  if (design) blocks.push('', design);
  blocks.push(
    '',
    '# 出力形式（厳守）',
    'JSONはチャットに貼らず、次の' + (design ? '3' : '2') + 'つのtxtファイルへ保存する。',
    '保存先は「【固定】ウェビナー自動生成/【出力】生成物/' + company + '/」（フォルダが無ければ作成）。毎回同じファイル名へ上書きする。',
    '',
    '【出力1: 台本JSON】→ 「【台本JSON】' + company + '.txt」',
    '```json',
    '[{"no": 1, "daihon": "..."}, {"no": 3, "daihon": "..."}, ...]',
    '```',
    '',
    '【出力2: 構成テキストJSON】→ 「【構成テキストJSON】' + company + '.txt」',
    '```json',
    '[{"no": 3, "kousei_text": "..."}, ...]',
    '```',
    design ? '【出力3: デザインJSON】→ 「【デザインJSON】' + company + '.txt」（段階3の指示どおりの形式）' : '',
    '',
    '- txtの中身はJSONだけにする（前置き・解説・```json などの記号は入れない）。',
    '- 1回で書き切れない場合は、同じファイルの続きへ書き足して1本にまとめる。',
    '- 全行が揃うまで書き切ること。省略・要約での置き換え（「以下同様」等）は禁止。',
    '- 保存後、ファイルごとに「保存先パス・入っている行番号・抜けの有無」を報告する。');
  return blocks.join('\n');
}

/** 背景画像フォルダ内の濃度サブフォルダのファイル名一覧 */
function scanBgNames_(cfg, opacity) {
  const names = [];
  const root = resolveNamedFolder_(cfg, 'BG');
  if (!root) return names;
  const it = foldersByName_(root, opacity);
  if (!it.hasNext()) return names;
  const files = it.next().getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType().indexOf('image/') === 0) names.push(f.getName());
  }
  return names.sort();
}

function createPromptDoc_(projectName, hearingId) {
  const cfg = getConfig_();
  const hearing = readHearing_(hearingId);
  const assetLists = {
    iconP: scanAssetNames_(cfg, 'ICON_P'),
    iconW: scanAssetNames_(cfg, 'ICON_W'),
    bg85: scanBgNames_(cfg, '85'),
    bg25: scanBgNames_(cfg, '25'),
    bg15: scanBgNames_(cfg, '15'),
    select: scanAssetNames_(cfg, 'SELECT')
  };
  const folder = caseFolder_(cfg, projectName);
  const doc = DocumentApp.create('【プロンプト】' + projectName);
  doc.getBody().setText(buildManualPrompt_(hearing, assetLists, projectName));
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(folder);
  // 見張りが拾って無人で作るための指示ファイル。資料とLPの2本。
  // ★Docも残してある。中身は指示ファイルと同じなので、自動生成が失敗したときは
  //   Docを開いてClaude Codeに貼れば手動で回せる（フォールバック）。
  createShiryoOrderFile_(projectName, hearing, assetLists, folder);
  createLpOrderFile_(projectName, hearing, folder);
  return doc;
}

// ============================================================
// LPプロンプトDoc（Claude Code向け・自己完結型）
// ============================================================
function buildLpPrompt_(projectName, hearing) {
  const company = caseCompany_(projectName);
  const head = [
    '# 【LPプロンプト】｛会社名｝ — Claude Codeへの作業指示',
    '',
    'あなた(Claude Code)は、作業フォルダとして共有ドライブ「社内｜Claude Code資料生成」が選択された状態で、この指示を受け取っています。次の手順で作業してください。',
    '',
    '1. 「【固定】ウェビナー自動生成/LP改善メモ.md」を読み込み、そこに書かれた各ルールを、下記のLP生成ルールと同格の追加ルールとして必ず適用する',
    '2. 下記のLP生成ルールに従い、末尾のヒアリング内容だけを根拠に、STEP1→2→3→4の順にすべて生成してチャットに表示する',
    '3. STEP4の作成前に「【固定】ウェビナー自動生成/【参考】サムネイル」フォルダ内の画像を実際に開いて確認し、今回の題材・ターゲット・トーンに最も合う1枚をNo.で選定する(ファイル名の番号=No.)。選定No.と選定理由を先に表示する',
    '4. ファイルとして保存するのは「確定タイトル(サブタイトル＋メインタイトル)」と「LP本文(リード文・▼こんな方におすすめ・▼学べる内容)」のみ。これをプレーンテキストで「【固定】ウェビナー自動生成/【出力】生成物/｛会社名｝/【LP】｛会社名｝.txt」として保存する(フォルダが無ければ作成)。このtxtは5分以内にシステムが自動でGoogleドキュメント「【LP】｛会社名｝」に変換する(既存のLPドキュメントがあれば同じドキュメントを上書き更新)ので、変換済みのGoogleドキュメントには触れない・削除しない。STEP1の整理・タイトル10案の表・サムネイル指示文はチャット表示のみとし、ファイルには含めない。txtの保存は常に同一ファイルへの上書きとし、「【LP】」で始まる古いtxt/md/docxが残っている場合は削除する',
    '5. 利用者が添削・修正・要望を伝えたら、必ず次の2つを行う: ①出力を修正し、タイトル・本文に関わる修正は【LP】｛会社名｝.txt を上書き保存する(新規ファイルの作成は禁止。Googleドキュメントへの反映はシステムが自動で行う) ②その添削を「【固定】ウェビナー自動生成/【受信箱】添削ログ.md」の末尾に記録する。形式は「## 日付 案件名 / LP（対象）」の見出しに続けて「- 指示:」「- 対応:」「- 状態: 未処理」の3行。ルールにするかどうかは後で人が仕分けるので、ここでは判断せず、そのまま記録する',
    '   ※「LP改善メモ.md」は配布用のコピーなので直接編集しないこと。ルールとして追加されるのは、仕分けと承認を経たあと。',
    '6. ファイルの保存・上書き・削除は、必ず作業フォルダ(ローカル同期された「社内｜Claude Code資料生成」)のファイル操作で行うこと。Google Driveコネクタ(API)でのファイル作成は禁止(削除・上書きができず重複の原因になるため)。作業フォルダが選択されていない場合は、作業を始める前に利用者へ「作業フォルダに社内｜Claude Code資料生成を選択してください」と伝えて中断すること',
    '7. 完了時に、保存先パス・選定したサムネイルNo.を報告する',
    '',
    '━━━━━━━━━━━━━━━━━━━━',
    'LP生成ルール',
    '━━━━━━━━━━━━━━━━━━━━',
    'あなたはBtoBウェビナー集客専門のコピーライターです。末尾のヒアリング内容だけを根拠に、①ウェビナータイトル案 ②LP本文 ③サムネイル画像の生成指示文 を作成します。',
    '',
    '【絶対ルール】',
    '・根拠はヒアリング内容のみ。書かれていない事実・数値・実績・効果を追加しない',
    '・煽らない。不安や危機感だけで引っ張らない。読み手を責めない',
    '・特定業界に寄せすぎず、ヒアリングに登場する人物像の言葉で書く',
    '',
    '## STEP1: 素材の整理(中間成果物として必ず表示)',
    '1. ターゲット人物像: 業界・役職・会社規模・事業フェーズ・置かれている状況',
    '2. 実際に起きている現象: 課題や感情ではなく、第三者が確認できる事実・出来事',
    '3. 強い悩み・欲求と本人の言葉: 焦り・不満・違和感と、実際に口にしそうなセリフ',
    '4. 本ウェビナーで解決できる悩み: ヒアリングの解決策で解消できるものを3つに絞る',
    '5. 解決の考え方: 新しい考え方(ビリーフ)・仕組み・ポイント(ヒアリングの記述のみ)',
    '6. ベネフィット: 参加後に日常業務レベルで何がどう変わるか(売上向上のような抽象語は禁止)',
    '',
    '## STEP2: ウェビナータイトル10案',
    '【構造】各案は「サブタイトル＋メインタイトル」の2部構成。サブタイトル=最重要価値の表現(「誰の・どんな現象が・どう解消されるか」が一読で伝わる。STEP1の2と6を接続)。メインタイトル=テーマと対象者の明示',
    '【本数】短め(合計25字前後)5本＋説明寄り(35〜45字)5本',
    '【ルール】問題提起で終わらせず解消後の状態まで含める/タイトルだけで誰向け・どんな状況向けか分かる/主人公が普段使う言葉を優先。記号・カギ括弧は使わない/重要なメッセージは左(先頭)に配置/キャッチーさより「伝わる」を優先',
    '【NG例】うまくいかない会社の共通点/失敗する人の特徴/成果が出ない原因/◯◯の進め方/◯◯で成果を出す方法',
    '【OK例の型】社員20名を超えて社長が現場を見切れなくなった会社の組織設計/紹介では決まるのに新規だと決まらない会社の営業改善法',
    '【自己チェック】10案を表にして「誰向けが明示/現象が具体/解消後の未来/字数」を○×評価し、最も優れた3案を理由付きで推薦する',
    '',
    '## STEP3: LP本文',
    '推薦1位のタイトルを前提に3部構成で作成。',
    '1. リード文: 悩みや違和感を短い文で複数提示→「こんな状態に、心当たりはありませんか?」等の問いかけ→原因は能力不足・努力不足ではないと否定→陥りがちな我流・誤解・思い込みを一般論として示す→感覚論ではなく構造的に整理することを伝える→精神論ではなく考え方や仕組みを扱うと明示→タイトルと自然につながる締め。文体: 無駄な改行なし/段落は意味のまとまりごと/煽らない/「まさに自分のことだ」と静かに刺さるトーン',
    '2. ▼こんな方におすすめ: 箇条書き3項目・各40字程度。状況や悩みを1文で具体的に描写。ヒアリングの人物像・課題感と一致させる',
    '3. ▼学べる内容: 箇条書き5項目・各50〜70字程度。「何が理解できるか」を表現。おすすめの各項目と対応関係が分かる内容。ヒアリング記載の解決策・成功要因・考え方のみ使用',
    'トーン: 断定しすぎない。過度なベネフィット表現をしない。「考え方が整理されそう」と感じる温度感',
    '',
    '## STEP4: サムネイル(バナー)生成指示文',
    'ChatGPTにそのまま貼れる指示文を1本作成する。指示文の中では選定した番号(No.)には一切触れず、「添付した参考画像」とだけ表現する。',
    '',
    '--- テンプレート ---',
    'ウェビナー集客用のバナー画像を作成してください。',
    '【サイズ】1600px × 1000px(横長)',
    '【最重要】添付した参考画像のデザイントーン・レイアウト・配色・文字組みを参照してください',
    '【テキスト要素(日本語を一字一句正確に)】',
    '・メインコピー(最も大きく): (サブタイトルをここに)',
    '・セミナー名: (メインタイトルをここに)',
    '・対象者ラベル: (◯◯の経営者向け 等)',
    '【レイアウト指定】',
    '・文字は視認性最優先。背景と文字のコントラストを強く',
    '【トーン】ビジネス向けで信頼感のある配色。過度な装飾・煽り表現は不可',
    '--- テンプレートここまで ---',
    '',
    '※運用メモ: この指示文をChatGPTに貼る際、選定された「サムネイルNo.◯」の画像1枚を必ず添付すること(【参考】サムネイルフォルダから)',
    '',
    '━━━━━━━━━━━━━━━━━━━━',
    'ヒアリング内容(これだけを根拠にする)',
    '━━━━━━━━━━━━━━━━━━━━'
  ].join('\n').replace(/｛会社名｝/g, company);
  return head + '\n\n【オファー記入用】\n' + hearing.offer +
    '\n\n【問題教育の記入用】\n' + hearing.problem +
    '\n\n【ストーリーの記入用】\n' + hearing.story + '\n';
}

/**
 * LP生成の指示ファイルを案件フォルダへ置く。
 * 宮津さんのMacの見張り（~/.claude/envy-webinar-watch/watch.sh）が1分ごとに拾い、
 * 無人でLPを生成して【LP】◯◯社.txt を書き、この指示ファイルを消す。
 *
 * ★名前を「【LPプロンプト】…」にしてはいけない。convertLpTexts_() が
 *   「【LP】で始まる.txt」をGoogleドキュメントへ変換して元txtを捨てるため、
 *   指示ファイルが処理される前に消える。だから「【指示】LP_」で始める。
 * ★Googleドキュメントではなくtxtにするのは、.gdoc はローカルには214バイトの
 *   ショートカットしか同期されず、Claude Codeが中身を読めないため。
 */
/**
 * 資料生成（構成テキスト・台本・デザインJSON）の指示ファイルを案件フォルダへ置く。
 * 見張り（宮津さんのMacの ~/.claude/envy-webinar-watch）が1分ごとに拾い、
 * 無人で3つのtxtを作る。その後は既存の processImports_ が構成シートへ流し込む。
 *
 * ★名前は「【指示】資料_」。他の見張りが拾う名前（【指示】LP_ / 【指示】修正LP_ /
 *   【指示】サムネ_）とも、convertLpTexts_ が拾う「【LP】」とも重ならない。
 */
function createShiryoOrderFile_(projectName, hearing, assetLists, folder) {
  const name = '【指示】資料_' + caseCompany_(projectName) + '.txt';
  const old = filesByName_(folder, name);
  while (old.hasNext()) old.next().setTrashed(true);
  return folder.createFile(name, buildManualPrompt_(hearing, assetLists, projectName), MimeType.PLAIN_TEXT);
}

function createLpOrderFile_(projectName, hearing, folder) {
  const name = '【指示】LP_' + caseCompany_(projectName) + '.txt';
  // 再生成のときに二重にならないよう、同名の古い指示ファイルは先に捨てる
  const old = filesByName_(folder, name);
  while (old.hasNext()) old.next().setTrashed(true);
  return folder.createFile(name, buildLpPrompt_(projectName, hearing), MimeType.PLAIN_TEXT);
}

// ============================================================
// LPの修正依頼（構成シート 設定タブ B9/B10）→ 指示ファイル
// ============================================================
/**
 * PMが構成シートに書いた修正依頼を拾い、指示ファイルにして案件フォルダへ置く。
 * 見張り（宮津さんのMacの ~/.claude/envy-webinar-watch）が1分ごとに拾って直す。
 *
 * ★いまのLP本文を指示ファイルに埋め込むのが肝。LPは既にGoogleドキュメントに
 *   変換され、元のtxtは捨てられているため、Claude Codeはローカルから中身を
 *   読めない（.gdoc は214バイトのショートカット）。GAS側で読んで渡す。
 * ★名前を「【指示】修正LP_」にする。「【LP】」で始めると convertLpTexts_ に
 *   拾われてDoc化され消える。「【指示】LP_」だと新規生成の見張りが拾ってしまう。
 */
/**
 * 構成シートの設定タブにLP修正依頼のセルが無ければ足す。
 * この機能より前に作られた案件でも使えるようにするため、毎回そっと確認する。
 * 既にあるときは何もしない（A8の見出しの有無で判定）。
 */
function ensureLpFixCells_(st) {
  try {
    // ★空のときだけ書く。既に何か書いてあるシートを上書きしない。
    //   （既存の構成シートは A1:A6 までしか使っていないが、手で書き足されている
    //     可能性があるので、無条件に setValues しない）
    const a8 = String(st.getRange('A8').getValue() || '').trim();
    if (a8 !== '') return;
    if (String(st.getRange('A9').getValue() || '').trim() !== '') return;
    if (String(st.getRange('A10').getValue() || '').trim() !== '') return;
    st.getRange('A8:A10').setValues([
      ['― LPを直したいとき ―'],
      ['直したい内容（まとめて書く）'],
      ['↑の内容で直す（チェックすると実行）']
    ]);
    st.getRange('A8').setFontWeight('bold').setBackground('#efefef');
    st.getRange('B9').setNote(
      '例: サブタイトルをもっと現象寄りに。学べる内容の3つ目が抽象的。\n' +
      '細かく何度も出すより、1回にまとめて書いてください。1回ごとにAIが動くので、\n' +
      '小刻みに出すとその分だけ費用がかかります。\n' +
      'チェックを入れると15分ほどでLPが直り、チェックは自動で外れます。'
    );
    st.getRange('B9').setWrap(true);
    if (st.getRange('B10').getDataValidation() === null) st.getRange('B10').insertCheckboxes();
    st.setRowHeight(9, 60);
  } catch (e) { /* 既存シートの体裁は壊さない。失敗しても黙って進む */ }
}

function processLpFixRequests_() {
  const cfg = getConfig_();
  const rows = mgmtSheet_(cfg).getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const hearingId = rows[i][0];
    const projectName = rows[i][1];
    const kouseiUrl = rows[i][3];
    if (!projectName || !kouseiUrl) continue;

    let st;
    try {
      st = SpreadsheetApp.openById(idFromUrl_(String(kouseiUrl))).getSheetByName(TABS.SETTINGS);
    } catch (e) { continue; }
    if (!st) continue;

    // 新しいセルが無い構成シート（この機能より前に作られたもの）には足しておく
    ensureLpFixCells_(st);

    if (st.getRange('B10').getValue() !== true) continue;

    // チェックは先に外す。途中で失敗しても同じ依頼を延々と回さないため。
    st.getRange('B10').setValue(false);

    const req = String(st.getRange('B9').getValue() || '').trim();
    if (!req) continue;

    try {
      const company = caseCompany_(projectName);
      const folder = caseFolderIfExists_(cfg, projectName);
      if (!folder) { st.getRange('B3').setValue('LP修正: 案件フォルダが見つかりません'); continue; }

      // いまのLP本文（Googleドキュメント）を読む
      let lpText = '';
      const it = filesByName_(folder, '【LP】' + company);
      while (it.hasNext()) {
        const f = it.next();
        if (f.getMimeType() === MimeType.GOOGLE_DOCS) {
          lpText = DocumentApp.openById(f.getId()).getBody().getText();
          break;
        }
      }
      if (!lpText) { st.getRange('B3').setValue('LP修正: LPドキュメントが見つかりません'); continue; }

      const hearing = readHearing_(hearingId);
      const body = [
        '# LP修正の指示 — ' + company,
        '',
        '## 直したい内容（PMからの依頼）',
        req,
        '',
        '## いまのLP本文（これを直す。1行目がサブタイトル、2行目がメインタイトル）',
        lpText,
        '',
        '## ヒアリング内容（事実の根拠。ここに書かれていないことは足さない）',
        '【オファー記入用】', hearing.offer,
        '【問題教育の記入用】', hearing.problem,
        '【ストーリーの記入用】', hearing.story
      ].join('\n');

      const fname = '【指示】修正LP_' + company + '.txt';
      const old = filesByName_(folder, fname);
      while (old.hasNext()) old.next().setTrashed(true);
      folder.createFile(fname, body, MimeType.PLAIN_TEXT);
      st.getRange('B3').setValue('LP修正: 依頼を受け付けました（15分ほどで直ります）');
    } catch (e) {
      st.getRange('B3').setValue('LP修正: エラー ' + String(e));
    }
  }
}


// ============================================================
// LPテキスト → Googleドキュメント自動変換
// （Claude Codeが保存した【LP】◯◯.txt を検知し、同名のGoogleドキュメントへ
//   変換。既存ドキュメントがあれば本文を上書き更新し、txtは削除する）
// ============================================================
function convertLpTexts_() {
  try {
    const cfg = getConfig_();
    const out = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
    const folders = out.getFolders();
    while (folders.hasNext()) {
      const f = folders.next();
      const files = f.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();
        if (name.indexOf('【LP】') !== 0) continue;
        if (!/\.(txt|md)$/i.test(name)) continue;
        let text = '';
        try { text = file.getBlob().getDataAsString('UTF-8'); } catch (e) { continue; }
        if (!String(text).trim()) continue;
        const docName = name.replace(/\.(txt|md)$/i, '');
        let doc = null;
        const it = filesByName_(f, docName);
        while (it.hasNext()) {
          const g = it.next();
          if (g.getMimeType() === MimeType.GOOGLE_DOCS) { doc = g; break; }
        }
        if (doc) {
          const d = DocumentApp.openById(doc.getId());
          d.getBody().clear();
          d.getBody().setText(text);
          d.saveAndClose();
        } else {
          const d = DocumentApp.create(docName);
          d.getBody().setText(text);
          d.saveAndClose();
          DriveApp.getFileById(d.getId()).moveTo(f);
        }
        file.setTrashed(true);
      }
    }
  } catch (e) { Logger.log('convertLpTexts_ error: ' + e); }
}


/**
 * 構成シートの設定タブにスライド不具合の報告セルが無ければ足す。
 * B13=不具合の内容 / B14=実行チェック / B15=結果（システムが書く）
 */
function ensureSlideFixCells_(st) {
  try {
    if (String(st.getRange('A12').getValue() || '').trim() !== '') return;
    if (String(st.getRange('A13').getValue() || '').trim() !== '') return;
    st.getRange('A12:A15').setValues([
      ['― スライドの不具合を直したいとき ―'],
      ['どこがどうおかしいか（まとめて書く）'],
      ['↑の原因をAIに調べてもらう（10〜35分・スライドは変わりません）'],
      ['結果（システムが書きます）']
    ]);
    st.getRange('A12').setFontWeight('bold').setBackground('#efefef');
    st.getRange('B13').setNote(
      '例: 42枚目の見出しが枠からはみ出している。79枚目のサービス名が2行に折り返している。\n' +
      '枚数と症状をセットで書いてください。スクリーンショットは要りません。\n' +
      '★これは「原因を調べる」ボタンです。スライドそのものは変わりません。\n' +
      'セルを直してスライドに反映したいだけなら、B16のほうを使ってください。\n' +
      '細かく何度も出すより、1回にまとめて書いてください。1回ごとにAIが動くので、\n' +
      '小刻みに出すとその分だけ費用がかかります。\n' +
      '10分ほどで結果がB15に入り、チェックは自動で外れます。'
    );
    st.getRange('B13').setWrap(true);
    st.getRange('B15').setWrap(true);
    if (st.getRange('B14').getDataValidation() === null) st.getRange('B14').insertCheckboxes();
    st.setRowHeight(13, 60);
    st.setRowHeight(15, 80);
  } catch (e) { /* 既存シートの体裁は壊さない */ }
}

/**
 * スライド不具合の報告を指示ファイルにし、返ってきた結果をシートへ書き戻す。
 * 調査は宮津さんのMacの見張りが行う。★本番への反映は人が手で行う。
 */
function processSlideFixRequests_() {
  const cfg = getConfig_();
  const rows = mgmtSheet_(cfg).getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const projectName = rows[i][1];
    const kouseiUrl = rows[i][3];
    const slideUrl = rows[i][4];
    if (!projectName || !kouseiUrl) continue;

    let st;
    try {
      st = SpreadsheetApp.openById(idFromUrl_(String(kouseiUrl))).getSheetByName(TABS.SETTINGS);
    } catch (e) { continue; }
    if (!st) continue;
    ensureSlideFixCells_(st);

    let folder;
    folder = caseFolderIfExists_(cfg, projectName); if (!folder) continue;

    // 見張りが書いた結果があればシートへ移して消す
    try {
      const res = filesByName_(folder, '【結果】スライド_' + caseCompany_(projectName) + '.txt');
      while (res.hasNext()) {
        const rf = res.next();
        st.getRange('B15').setValue(rf.getBlob().getDataAsString('UTF-8'));
        rf.setTrashed(true);
      }
    } catch (e) { /* 取り込み失敗でも受付は続ける */ }

    if (st.getRange('B14').getValue() !== true) continue;
    st.getRange('B14').setValue(false);
    const req = String(st.getRange('B13').getValue() || '').trim();
    if (!req) continue;

    try {
      const body = [
        '# スライド不具合の調査依頼 — ' + caseCompany_(projectName),
        '',
        '## PMからの報告',
        req,
        '',
        '## スライド',
        slideUrl ? String(slideUrl) : '（URLが管理シートにありません。案件フォルダの【スライド】を開いてください）',
        '',
        '## 案件フォルダ',
        '【固定】ウェビナー自動生成/【出力】生成物/' + projectName + '/'
      ].join('\n');
      const fname = '【指示】スライド_' + caseCompany_(projectName) + '.txt';
      const old = filesByName_(folder, fname);
      while (old.hasNext()) old.next().setTrashed(true);
      folder.createFile(fname, body, MimeType.PLAIN_TEXT);
      st.getRange('B15').setValue('調べています（10分ほどお待ちください）');
    } catch (e) {
      st.getRange('B15').setValue('エラー: ' + String(e));
    }
  }
}

/**
 * タイトル確定のセルを用意する（設定タブ A17:A19 / B18:B19）。
 * ★「タイトル案を見て、選んでからサムネイル」の関門。
 *   これが無いと、LPができた瞬間にサムネイルが自動で作られてしまい、
 *   PMがタイトルを選ぶ余地が無くなる。
 * A17=見出し / A18=タイトル案の置き場所（B18にリンク） / A19=確定チェック（B19）
 */
function ensureTitleConfirmCells_(st) {
  try {
    if (String(st.getRange('A17').getValue() || '').trim() !== '') return;
    if (String(st.getRange('A18').getValue() || '').trim() !== '') return;
    st.getRange('A17:A19').setValues([
      ['― タイトルを決めるとサムネイルを作ります ―'],
      ['タイトル案10本（システムが書きます）'],
      ['この案でよければチェック（サムネイル生成）']
    ]);
    st.getRange('A17').setFontWeight('bold').setBackground('#efefef');
    st.getRange('B19').setNote(
      'B18のタイトル案を開いて、推薦1位でよければここにチェックを入れてください。\n' +
      '10分ほどでサムネイル画像ができます。\n' +
      '別の案がよければ、ここは押さずにB9へ「タイトルを案7にしてください」と書いてB10を押してください。\n' +
      'その場合はLPが書き直され、サムネイルも新しいタイトルで自動的に作られます。'
    );
    if (st.getRange('B19').getDataValidation() === null) st.getRange('B19').insertCheckboxes();
  } catch (e) { /* 既存シートの体裁は壊さない */ }
}

/**
 * タイトル案のリンクをB18に出し、B19が押されたら
 * 【保留】サムネ指示_◯◯.txt を【指示】サムネ_◯◯.txt に改名する。
 * 改名した瞬間から、宮津さんのMacの見張りがそれを拾って画像を作る。
 */
function processTitleConfirm_() {
  const cfg = getConfig_();
  const rows = mgmtSheet_(cfg).getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const projectName = rows[i][1];
    const kouseiUrl = rows[i][3];
    if (!projectName || !kouseiUrl) continue;

    let st;
    try {
      st = SpreadsheetApp.openById(idFromUrl_(String(kouseiUrl))).getSheetByName(TABS.SETTINGS);
    } catch (e) { continue; }
    if (!st) continue;

    ensureTitleConfirmCells_(st);

    let folder;
    folder = caseFolderIfExists_(cfg, String(projectName)); if (!folder) continue;
    const company = caseCompany_(String(projectName));

    // タイトル案ができていたら、その場所をB18に出す（PMが探さなくて済むように）
    try {
      if (!String(st.getRange('B18').getValue() || '').trim()) {
        const tf = filesByName_(folder, '【タイトル案】' + company + '.txt');
        if (tf.hasNext()) st.getRange('B18').setValue(tf.next().getUrl());
      }
    } catch (e) { /* リンクは補助。失敗しても確定処理は続ける */ }

    if (st.getRange('B19').getValue() !== true) continue;
    st.getRange('B19').setValue(false);   // 先に外す。失敗しても同じ依頼を回し続けない

    try {
      const held = filesByName_(folder, '【保留】サムネ指示_' + company + '.txt');
      if (!held.hasNext()) {
        st.getRange('B3').setValue('サムネイル: 保留中の指示が見つかりません（LPがまだできていないか、既に作成済み）');
        continue;
      }
      held.next().setName('【指示】サムネ_' + company + '.txt');
      st.getRange('B3').setValue('サムネイル: 作り始めます（10分ほどで素材フォルダにできます）');
    } catch (e) {
      st.getRange('B3').setValue('サムネイル: エラー ' + String(e));
    }
  }
}

/**
 * スライドを作り直すチェック（設定タブ A16 / B16）。
 * ★不具合の調査（B14）とは別物。こちらはAIを使わず、
 *   今の4タブの中身からスライドを組み直すだけ。
 *   これまで再生成には「管理シートC列を戻す→mp3を消す→承認OFF→ON」の
 *   3手順が必要で、面倒なうえにmp3の消し忘れ事故が起きていた。
 */
function ensureSlideRegenCell_(st) {
  try {
    if (String(st.getRange('A16').getValue() || '').trim() !== '') return;
    st.getRange('A16').setValue('【別の操作】セルを直したあとスライドを作り直す（数分・AIは使いません）');
    st.getRange('B16').setNote(
      '4タブのセルを直したあと、ここにチェックを入れるとスライドを作り直します。AIは動きません。\n' +
      '古いスライド・台本txt・音声mp3は自動でゴミ箱へ移り、音声も作り直されます。\n' +
      '★スライドを手で直接編集していた場合、その編集は消えます。\n' +
      '不具合の原因を調べてほしいときは、B13とB14のほうを使ってください。'
    );
    if (st.getRange('B16').getDataValidation() === null) st.getRange('B16').insertCheckboxes();
  } catch (e) { /* 既存シートの体裁は壊さない */ }
}
