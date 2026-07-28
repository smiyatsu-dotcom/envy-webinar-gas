/**
 * 05_bootstrap.gs（実測規定反映版）
 * ワンショット初期構築: フォルダ（アイコン/背景サブフォルダ含む）・管理シート・
 * スライドテンプレート（実測フォントサイズの24レイアウト）を自動作成する。
 * 使い方: bootstrapAll() を1回実行するだけ。APIキー不要。
 */

function bootstrapAll_() {
  const p = PropertiesService.getScriptProperties();

  const root = DriveApp.createFolder('ウェビナー資料自動生成');
  const inFolder = root.createFolder('【入力】ヒアリングシート');
  const outFolder = root.createFolder('【出力】生成物');
  const assetsFolder = root.createFolder('【素材】投影資料');
  // 素材サブフォルダ（既にユーザー所有フォルダがある場合はスクリプトプロパティでID指定可）
  Object.keys(ASSET_SUB).forEach(function (k) { assetsFolder.createFolder(ASSET_SUB[k]); });

  const mgmt = SpreadsheetApp.create('【管理】ウェビナー自動生成ダッシュボード');
  DriveApp.getFileById(mgmt.getId()).moveTo(root);

  const template = buildTemplateDeck_();
  DriveApp.getFileById(template.getId()).moveTo(root);

  p.setProperties({
    INPUT_FOLDER_ID: inFolder.getId(),
    OUTPUT_FOLDER_ID: outFolder.getId(),
    ASSETS_FOLDER_ID: assetsFolder.getId(),
    MGMT_SHEET_ID: mgmt.getId(),
    TEMPLATE_SLIDE_ID: template.getId()
  });
  setup();

  Logger.log('===== 初期構築 完了 =====');
  Logger.log('ルート: ' + root.getUrl());
  Logger.log('入力フォルダ（ヒアリングを置く）: ' + inFolder.getUrl());
  Logger.log('素材フォルダ（図解を案件名サブフォルダへ、アイコン/背景を各サブフォルダへ）: ' + assetsFolder.getUrl());
  Logger.log('テンプレデッキ（デザイン調整可）: ' + template.getUrl());
  Logger.log('管理シート: ' + mgmt.getUrl());
}

// ============================================================
// テンプレデッキ生成（実測フォントサイズ）
// 色: 紺 #0F2A43 / ゴールド #C9A227 / 白 / グレー #F2F4F7
// ============================================================
// サンプル1・2実測パレット（両デッキ共通）
const C_ = {
  purple: '#7F61AD',  // メインアクセント（円・ラベル）
  ink: '#434343',     // 基本文字色
  brown: '#504036',   // 強調文字（サービス名等）
  red: '#DD4031',     // 二重線・強調
  bg: '#F2F2F2',      // ベース背景（薄グレー）
  panel: '#D9D9D9',   // 帯・パネル
  dark: '#434343',    // 問いかけ等の濃背景
  white: '#FFFFFF'
};

function buildTemplateDeck_() {
  const pres = SlidesApp.create('【テンプレ】ウェビナースライド');
  const W = pres.getPageWidth(), H = pres.getPageHeight();
  const mk = [tCover_, tQuestion_, tIdeal_, tNarrative_, tText55_, tDiagram_, tDiagram40_, tCase_, tService_,
              tMeritIntro_, tMerit_, tPointIntro_, tPoint1_, tPoint2_, tChallenge_, tFeatures_, tSteps_,
              tImpact_, tPrice_, tInvest_, tStrike_, tProfile_, tRinen_, tKotei_, tZoom_, tMeigen_, tByaf_, tOrei_];
  mk.forEach(function (f) { f(pres, W, H); });
  pres.getSlides()[0].remove();
  return pres;
}

// ---- 共通部品 ----
function base_(pres, tag, bg) {
  const s = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s.getBackground().setSolidFill(bg || C_.bg);
  const n = s.getNotesPage().getSpeakerNotesShape();
  if (n) n.getText().setText('{{TPL:' + tag + '}}');
  return s;
}
function tb_(s, ph, l, t, w, h, size, color, opts) {
  opts = opts || {};
  const box = s.insertTextBox(ph, l, t, w, h);
  const tr = box.getText();
  tr.getTextStyle().setFontSize(size).setForegroundColor(color).setBold(opts.bold !== false).setFontFamily('Noto Sans JP');
  tr.getParagraphStyle().setParagraphAlignment(opts.align || SlidesApp.ParagraphAlignment.CENTER);
  box.setContentAlignment(opts.valign || SlidesApp.ContentAlignment.MIDDLE);
  return box;
}
function img_(s, l, t, w, h) { // 図解枠
  const ph = s.insertShape(SlidesApp.ShapeType.RECTANGLE, l, t, w, h);
  ph.getFill().setSolidFill(C_.bg);
  ph.getBorder().setDashStyle(SlidesApp.DashStyle.DASH).getLineFill().setSolidFill('#9AA1A9');
  ph.setTitle('IMG');
  ph.getText().setText('{{IMG}}');
  ph.getText().getTextStyle().setFontSize(12).setForegroundColor('#9AA1A9');
  ph.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  ph.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  return ph;
}
function overlayBand_(s, W) { // テキスト背面の紺50%帯（実測: y143 高さ164pt）
  const ov = s.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 143, W, 164);
  ov.getFill().setSolidFill(C_.panel, 0.5);
  ov.getBorder().setTransparent();
  return ov;
}
function icon_(s, l, t, d) { // アイコン枠（円）
  const ph = s.insertShape(SlidesApp.ShapeType.ELLIPSE, l, t, d, d);
  ph.getFill().setSolidFill('#FFFFFF');
  ph.getBorder().setDashStyle(SlidesApp.DashStyle.DASH).getLineFill().setSolidFill(C_.purple);
  ph.setTitle('ICON');
  ph.getText().setText('{{ICON}}');
  ph.getText().getTextStyle().setFontSize(10).setForegroundColor('#9AA1A9');
  ph.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  ph.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  return ph;
}
function band_(s, W, ph) { // 上部帯見出し 20pt
  const bar = s.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 30, W - 80, 47);
  bar.getFill().setSolidFill(C_.white);
  bar.getBorder().setTransparent();
  const tr = bar.getText();
  tr.setText(ph);
  tr.getTextStyle().setFontSize(20).setForegroundColor(C_.ink).setBold(true).setFontFamily('Noto Sans JP');
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  bar.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  return bar;
}
function numCircle_(s, l, t, d, num, fs) {
  const c = s.insertShape(SlidesApp.ShapeType.ELLIPSE, l, t, d, d);
  c.getFill().setSolidFill(C_.purple);
  c.getBorder().setTransparent();
  const tr = c.getText();
  tr.setText(String(num));
  tr.getTextStyle().setFontSize(fs).setForegroundColor(C_.white).setBold(true);
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  c.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  return c;
}
function static_(s, text, l, t, w, h, size, color) {
  return tb_(s, text, l, t, w, h, size, color);
}

// ---- 各レイアウト（フォントサイズは実測規定） ----
function tCover_(p, W, H) { const s = base_(p, 'cover'); img_(s, 0, 0, W, H); }
function tQuestion_(p, W, H) { const s = base_(p, 'question', C_.dark); tb_(s, '{{P1}}', W*0.05, H*0.3, W*0.9, H*0.32, 45, C_.white); }
function tIdeal_(p, W, H) { // 45pt 1-2行（背景画像対象。紺50%帯の上に文字）
  const s = base_(p, 'ideal');
  overlayBand_(s, W);
  tb_(s, '{{P1}}', 0, 166, W, 117, 45, C_.white);
}
function tNarrative_(p, W, H) { // リード30 + メイン55（背景画像対象）
  const s = base_(p, 'narrative');
  overlayBand_(s, W);
  tb_(s, '{{P1}}', 0, 148, W, 46, 30, C_.white);
  tb_(s, '{{P2}}', 0, 196, W, 100, 55, C_.ink);
}
function tText55_(p, W, H) { const s = base_(p, 'text55'); overlayBand_(s, W); tb_(s, '{{P1}}', 0, 150, W, 150, 55, C_.ink); }
function tDiagram_(p, W, H) { const s = base_(p, 'diagram', C_.white); tb_(s, '{{P1}}', W*0.05, 24, W*0.9, 50, 20, C_.ink); img_(s, W*0.1, 90, W*0.8, H-130); }
function tDiagram40_(p, W, H) { const s = base_(p, 'diagram40', C_.white); tb_(s, '{{P1}}', W*0.05, 20, W*0.9, 66, 40, C_.ink); img_(s, W*0.1, 100, W*0.8, H-140); }
function tCase_(p, W, H) { // タイトル30 + 図解 + 社長名/業種役職20 + 本文20（実測配置）
  const s = base_(p, 'case', C_.white);
  tb_(s, '{{P1}}', 0, 0, W, 91, 30, C_.ink);
  img_(s, 48, 74, 623, 280);
  tb_(s, '{{P3}}', 429, 254, 265, 42, 20, C_.ink);                    // 社長名
  tb_(s, '{{P4}}', 429, 296, 265, 42, 20, C_.ink, { bold: false });   // 業種＋役職
  tb_(s, '{{P2}}', 20, 366, 680, 84, 20, C_.ink, { bold: false });    // 本文（最大2行・各27字）
}
function tService_(p, W, H) { // サービス名45 + 概要15 + モックアップ
  const s = base_(p, 'service');
  tb_(s, '{{P1}}', 0, 26, W, 70, 45, C_.purple);
  tb_(s, '{{P2}}', W*0.08, 100, W*0.84, 58, 15, C_.ink, { bold: false });
  img_(s, W*0.2, 168, W*0.6, H-200);
}
function tMeritIntro_(p, W, H) { // 帯 + 数字兦×3（227pt / 60pt）
  const s = base_(p, 'merit_intro');
  band_(s, W, '{{P1}}');
  [20, 247, 473].forEach(function (x, i) { numCircle_(s, x, 112, 227, i + 1, 60); });
}
function tMerit_(p, W, H) { // 帯(No.41再掲) + 円下見出し19pt×3 + 説明20pt
  const s = base_(p, 'merit');
  band_(s, W, '{{H}}');
  [20, 247, 473].forEach(function (x, i) {
    numCircle_(s, x + 38, 112, 150, i + 1, 48);
    icon_(s, x + 72, 136, 82).setTitle('ICON' + (i + 1)); // 実測: 82×82pt
    tb_(s, '{{M' + (i + 1) + '}}', x, 272, 227, 46, 19, C_.purple);
  });
  tb_(s, '{{P2}}', 20, 360, 680, 58, 20, C_.white, { bold: false });
}
function tPointIntro_(p, W, H) { // 帯 + 数字大兦×2（326pt / 86pt）
  const s = base_(p, 'point_intro');
  band_(s, W, '{{P1}}');
  numCircle_(s, 34, 106, 326, 1, 86);
  numCircle_(s, 360, 106, 326, 2, 86);
}
function tPoint1_(p, W, H) { // 円内30pt×1
  const s = base_(p, 'point1');
  band_(s, W, '{{H}}');
  const c = s.insertShape(SlidesApp.ShapeType.ELLIPSE, 34, 106, 326, 326);
  c.getFill().setSolidFill(C_.white); c.getBorder().getLineFill().setSolidFill(C_.purple);
  icon_(s, 149, 153, 96).setTitle('ICON1'); // 実測: 96×96pt
  tb_(s, '{{M1}}', 34, 220, 326, 100, 30, C_.ink);
  numCircle_(s, 360, 106, 326, 2, 86);
}
function tPoint2_(p, W, H) { // 円内30pt×2（①はNo.51再掲）
  const s = base_(p, 'point2');
  band_(s, W, '{{H}}');
  [34, 360].forEach(function (x, i) {
    const c = s.insertShape(SlidesApp.ShapeType.ELLIPSE, x, 106, 326, 326);
    c.getFill().setSolidFill(C_.white); c.getBorder().getLineFill().setSolidFill(C_.purple);
    icon_(s, x + 115, 153, 96).setTitle('ICON' + (i + 1));
    tb_(s, '{{M' + (i + 1) + '}}', x, 220, 326, 100, 30, C_.ink);
  });
}
function tChallenge_(p, W, H) { // タイトル55 + 箇条書き30
  const s = base_(p, 'challenge');
  tb_(s, '{{P1}}', 0, 46, W, 76, 55, C_.purple);
  tb_(s, '{{P2}}', W*0.12, 150, W*0.76, 220, 30, C_.white, { bold: false, align: SlidesApp.ParagraphAlignment.START });
}
function tFeatures_(p, W, H) { // サービス名45 + カード3枚（見出し14/補足11）
  const s = base_(p, 'features', C_.white);
  tb_(s, '{{P1}}', 0, 20, W, 70, 45, C_.brown);
  [57, 267, 477].forEach(function (x, i) {
    const cir = s.insertShape(SlidesApp.ShapeType.ELLIPSE, x + 45, 158, 92, 92);
    cir.getFill().setSolidFill(C_.purple); cir.getBorder().setTransparent();
    icon_(s, x + 57, 168, 72).setTitle('ICON' + (i + 1)); // 白アイコン 実測72×72pt
    const card = s.insertShape(SlidesApp.ShapeType.RECTANGLE, x, 250, 186, 139);
    card.getFill().setSolidFill(C_.bg);
    card.getBorder().setTransparent();
    tb_(s, '特徴' + ['①','②','③'][i], x, 258, 186, 20, 12, C_.purple);
    tb_(s, '{{F' + (i + 1) + 'H}}', x + 6, 282, 174, 24, 14, C_.ink);
    tb_(s, '{{F' + (i + 1) + 'B}}', x + 8, 310, 170, 66, 11, C_.ink, { bold: false });
  });
}
function tSteps_(p, W, H) { // 36 + 20 + ステップ名12×4
  const s = base_(p, 'steps');
  tb_(s, '{{P1}}', 0, 24, W, 56, 36, C_.brown);
  tb_(s, '{{P2}}', 0, 86, W, 40, 20, C_.ink, { bold: false });
  [40, 212, 384, 556].forEach(function (x, i) {
    icon_(s, x + 40, 132, 72).setTitle('ICON' + (i + 1)); // 実測72×72pt
    const cell = s.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 216, 152, 110);
    cell.getFill().setSolidFill(C_.white); cell.getBorder().setTransparent();
    tb_(s, 'ステップ' + ['①','②','③','④'][i], x, 226, 152, 22, 12, '#9AA1A9');
    tb_(s, '{{S' + (i + 1) + '}}', x + 4, 252, 144, 60, 12, C_.ink);
  });
}
function tImpact_(p, W, H) { // 70 + 40
  const s = base_(p, 'impact');
  tb_(s, '{{P1}}', 0, 130, W, 110, 70, C_.brown);
  tb_(s, '{{P2}}', 0, 252, W, 64, 40, C_.ink);
}
function tPrice_(p, W, H) { // 45 + 価格行40
  const s = base_(p, 'price');
  tb_(s, '{{P1}}', 0, 34, W, 70, 45, C_.brown);
  tb_(s, '{{P2}}', W*0.08, 130, W*0.84, 240, 40, C_.ink);
}
function tInvest_(p, W, H) { investLike_(p, 'invest'); }
function tStrike_(p, W, H) { investLike_(p, 'strike'); }
function investLike_(p, tag) {
  const W = p.getPageWidth();
  const s = base_(p, tag, C_.bg);
  tb_(s, '{{P1}}', 0, 96, W, 50, 30, C_.ink, { bold: false });
  tb_(s, '{{P2}}', 0, 156, W, 120, 80, C_.brown);
  tb_(s, '{{P3}}', 0, 290, W, 50, 30, C_.ink, { bold: false });
}
function tProfile_(p, W, H) {
  const s = base_(p, 'profile', C_.white);
  tb_(s, '講師紹介', W*0.06, 24, 200, 40, 24, C_.ink, { align: SlidesApp.ParagraphAlignment.START });
  img_(s, W*0.06, 90, W*0.32, H-130);
  tb_(s, '{{P1}}', W*0.44, 90, W*0.5, H-130, 13, C_.ink, { bold: false, align: SlidesApp.ParagraphAlignment.START, valign: SlidesApp.ContentAlignment.TOP });
}
function tRinen_(p, W, H) { // 45pt理念。上下の実績画像帯はテンプレ側で固定（差し込みなし）
  const s = base_(p, 'rinen');
  [0, 308].forEach(function (y) {
    const strip = s.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, y, W, 143);
    strip.getFill().setSolidFill(C_.bg);
    strip.getBorder().setTransparent();
    strip.getText().setText('装飾画像エリア（テンプレ側で実画像に差し替え・固定）');
    strip.getText().getTextStyle().setFontSize(10).setForegroundColor('#9AA1A9');
    strip.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    strip.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  });
  tb_(s, '{{P1}}', 0, 154, W, 143, 45, C_.ink);
}
function tKotei_(p, W, H) { // 上部ステップ帯（アイコン72pt×4+ステップ名。非現在はGASが減光）+ 図解
  const s = base_(p, 'kotei');
  [40, 212, 384, 556].forEach(function (x, i) {
    icon_(s, x + 40, 26, 72).setTitle('ICON' + (i + 1)); // 実測: y26 72×72pt
    tb_(s, '{{S' + (i + 1) + '}}', x, 102, 152, 20, 12, C_.ink);
  });
  img_(s, 0, 169, W, 280); // 実測: (0,169) 720×280pt
}
function tZoom_(p, W, H) {
  const s = base_(p, 'static_zoom');
  const b = s.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, W*0.3, 66, W*0.4, 52);
  b.getFill().setSolidFill(C_.purple); b.getBorder().setTransparent();
  b.getText().setText('本日限定');
  b.getText().getTextStyle().setFontSize(20).setBold(true).setForegroundColor(C_.white);
  b.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  b.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  static_(s, '本日ZOOM終了までに\nお申し込みいただいた方には\n追加特典をご用意しています', W*0.1, 150, W*0.8, 190, 26, C_.ink);
}
function tMeigen_(p, W, H) {
  const s = base_(p, 'static_meigen', C_.bg);
  static_(s, '「今、上手くいっているからと\n今の仕事のやり方だけを続ける経営者は\n必然的に失敗する」', W*0.08, 90, W*0.84, 190, 24, C_.ink);
  const w = s.insertTextBox('― ピーター・ドラッカー', W*0.5, 300, W*0.42, 34);
  w.getText().getTextStyle().setFontSize(16).setForegroundColor('#9AA1A9');
  w.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
}
function tByaf_(p, W, H) {
  const s = base_(p, 'static_byaf', C_.white);
  static_(s, 'どちらでも構いません。\n\n大切なのは、\nご自身の意思で決断することです。', W*0.1, 108, W*0.8, 225, 26, C_.ink);
}
function tOrei_(p, W, H) {
  const s = base_(p, 'static_orei', C_.dark);
  static_(s, '最後までご覆聴いただき\nありがとうございました', W*0.1, 144, W*0.8, 153, 30, C_.white);
}


/**
 * テンプレデッキだけを実配色で作り直す（構築済み環境用）。。
 * 旧デッキは残る（名前に[旧]を付与）。プロパティ TEMPLATE_SLIDE_ID は新デッキへ更新される。
 */
function rebuildTemplateDeck_() {
  const p = PropertiesService.getScriptProperties();
  const oldId = p.getProperty('TEMPLATE_SLIDE_ID');
  const tpl = buildTemplateDeck_();
  if (oldId) {
    try {
      const oldFile = DriveApp.getFileById(oldId);
      oldFile.setName('[旧] ' + oldFile.getName());
      const parents = oldFile.getParents();
      if (parents.hasNext()) DriveApp.getFileById(tpl.getId()).moveTo(parents.next());
    } catch (e) { /* 旧デッキなしでも続行 */ }
  }
  p.setProperty('TEMPLATE_SLIDE_ID', tpl.getId());
  Logger.log('新テンプレデッキ: ' + tpl.getUrl());
}




















// ============================================================
// 本番テンプレの切り替え（サンプル108枚デッキを差し替えエンジンの原本にする）
// ============================================================
const SAMPLE_DECK_ID = '1hU8pMzPazyeTRcr_NlF3xDKx-9Gpc2WWpWD7zdcxVKY'; // サンプル1ベース・松竹梅組み込み済み108枚

/** 1回実行: TEMPLATE_SLIDE_ID をサンプル108枚デッキに設定する（統合版の必須初期化） */
function useSampleTemplate() {
  PropertiesService.getScriptProperties().setProperty('TEMPLATE_SLIDE_ID', SAMPLE_DECK_ID);
  const f = DriveApp.getFileById(SAMPLE_DECK_ID);
  Logger.log('TEMPLATE_SLIDE_ID を設定しました: ' + f.getName() + ' / ' + f.getUrl());
}
