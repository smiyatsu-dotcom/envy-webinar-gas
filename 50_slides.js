/**
 * 50_slides.gs（統合版: pilot14エンジン + 構成シート実データ配線）
 *
 * 方式: 完成済みサンプルデッキ（108枚・松竹梅変異版込み）を複製し、
 *       テキストボックスの位置・行数から差し替え箇所を特定して行単位で置換する。
 *       元の書式（フォントサイズ・色・不透明度）はサンプルのまま継承される。
 *
 * 画像・アイコン・背景: 既存画像を Image.replace() で差し替える。
 *   ※ replaceすると不透明度は100%にリセットされるため（GASに設定APIなし）、
 *     サンプル1実測値（OPACITY_RULES）に基づき白の半透明オーバーレイを画像直上に
 *     重ねて等価な見た目を再現する。行別に「ファイル名@25」で上書き可。
 */

function generateSlides_(kouseiSheetId) {
  const cfg = getConfig_();
  const data = readKouseiSheet_(kouseiSheetId); // rows + ryoukin（単一/松竹梅)
  const warnings = [];

  // 1. サンプルデッキ（108枚）を複製（再生成時、同名の旧スライドはゴミ箱へ移す）
  const outFolder = caseFolder_(cfg, data.name);
  const slideName = '【スライド】' + data.name;
  const trashed = trashSameName_(outFolder, slideName);
  if (trashed) warnings.push('旧スライド' + trashed + '件をゴミ箱へ移しました（30日以内なら復元可）');
  const copyFile = DriveApp.getFileById(cfg.TEMPLATE_SLIDE_ID).makeCopy(slideName, outFolder);
  const pres = SlidesApp.openById(copyFile.getId());
  const slides = pres.getSlides();
  // ★Q&A（No.104.5）はテンプレに入れていない（複製して作る）ので、テンプレは108枚のまま
  if (slides.length !== 108) warnings.push('テンプレ枚数が108ではありません: ' + slides.length + '枚（TEMPLATE_SLIDE_IDがサンプル108枚版か確認）');

  const zukaiFolder = resolveAssetsFolder_(cfg, data, warnings);
  const idx = { ICON_P: buildAssetIndex_(cfg, 'ICON_P'), ICON_W: buildAssetIndex_(cfg, 'ICON_W'),
                SELECT: buildAssetIndex_(cfg, 'SELECT'),
                BG85: buildBgIndex_(cfg, '85'), BG25: buildBgIndex_(cfg, '25'), BG15: buildBgIndex_(cfg, '15') };
  const targets = designTargets_();

  // No.→スライド（108枚: 松竹梅版No.96/102が単一版の直後に挿入されている）→ slideIndexFor_
  function slideFor(no) { const i = slideIndexFor_(no); return i < 0 ? null : slides[i]; }

  // Q&A（No.104.5）: 構成シートにQ&Aの行があり、中身があるときだけ、テンプレの No.104 を複製して
  // その直後（＝No.105の前）に1枚足す。★複製はテンプレのままの No.104 から取るため、ループの前に行う
  // （ループで No.104 が斜線スキップ・文言差し替えされたあとに複製すると、それを引き継いでしまう）。
  // ★slides はここより前に取った配列なので、足してもその中の番号はずれない（No.105以降・料金バリアントの削除も同じ）。
  // ★Q&Aの行が無い構成シート（この仕組みより前の案件）では何もしない。B16・B2を押し直しても割り込まない。
  let qaSlide = null;
  const qaRow = data.rows[QA_NO];
  if (qaRow && !isSkipRow_({ no: QA_NO }, qaRow)) {
    try {
      qaSlide = slideFor(104).duplicate();
      try { qaSlide.setSkipped(false); } catch (e) { }
    } catch (e) { warnings.push('No.' + QA_NO + '（Q&A）スライドの複製に失敗: ' + e); }
  }
  function matsuFor(no) { return no === 96 ? slides[96] : (no === 102 ? slides[103] : null); } // 松竹梅版

  rowDefs_().forEach(function (def) {
    const row = data.rows[def.no];
    if (def.no === QA_NO) {   // Q&Aは記入がある案件だけ。無ければ何も入れず、警告も出さない
      if (!qaSlide) return;
      try { fillQa_(qaSlide, row, pres, warnings); }
      catch (e) { warnings.push('No.' + QA_NO + '（Q&A）テキスト差し替え失敗: ' + e); }
      setNotesText_(qaSlide, row.daihon || '');
      return;
    }
    if (!row) { warnings.push('No.' + def.no + ' が構成シートにありません'); return; }
    const isMatsuTarget = (def.no === 96 || def.no === 102);
    const slide = (isMatsuTarget && data.ryoukin === '松竹梅') ? matsuFor(def.no) : slideFor(def.no);
    if (!slide) return;

    if (isSkipRow_(def, row)) {
      fillSkipSample_(slide, def);
      setNotesText_(slide, '');
      try { slide.setSkipped(true); } catch (e) { warnings.push('No.' + def.no + ' スキップ設定失敗: ' + e); }
      return;
    }

    try { fillReal_(slide, def, row, data, warnings); }
    catch (e) { warnings.push('No.' + def.no + ' テキスト差し替え失敗: ' + e); }

    // No.92/95: 「以上」は大きくしない(大きくするのは金額の数字のみ)
    if (def.no === 92 || def.no === 95) {
      try {
        slide.getShapes().forEach(function (sh) {
          sh.getText().find('以上').forEach(function (r) { r.getTextStyle().setFontSize(30).setBold(false); });
        });
      } catch (e) { }
    }

    // No.40: 最上部のタイトルはArial
    if (def.no === 40) {
      try {
        let top = null;
        slide.getShapes().forEach(function (sh) {
          const s = txt_(sh);
          if (!s) return;
          if (!top || sh.getTop() < top.getTop()) top = sh;
        });
        if (top) top.getText().getTextStyle().setFontFamily('Arial');
      } catch (e) { }
    }

    setNotesText_(slide, row.daihon || '');

    try { replaceImages_(slide, def, row, data, zukaiFolder, idx, targets, warnings); }
    catch (e) { warnings.push('No.' + def.no + ' 画像差し替え失敗: ' + e); }
  });

  // 2. 台本txt(音声化用)を【出力】へ書き出し
  exportDaihonTxt_(cfg, data, warnings);

  // 3. 料金体系: 該当しない側の No.96/102 を削除（インデックスの大きい方から）
  try {
    if (data.ryoukin === '松竹梅') { slides[102].remove(); slides[95].remove(); }
    else { slides[103].remove(); slides[96].remove(); }
  } catch (e) { warnings.push('料金バリアント削除失敗: ' + e); }

  pres.saveAndClose();
  if (warnings.length > 0) Logger.log('スライド生成警告:\n' + warnings.join('\n'));
  return { pres: SlidesApp.openById(copyFile.getId()), warnings: warnings };
}

/**
 * No.→テンプレ（108枚）の何枚目か（0始まり）。
 * 松竹梅版のNo.96/102が単一版の直後に1枚ずつ入っているので、97〜102は+0、103以降は+1。
 * ★整数のNo.だけ。Q&A（No.104.5）はテンプレに無い（No.104を複製して作る）ので -1。
 */
function slideIndexFor_(no) {
  if (no !== Math.floor(no) || no < 1 || no > 106) return -1;
  return no <= 96 ? no - 1 : (no <= 102 ? no : no + 1);
}

/**
 * 台本を読み上げ順に連結したtxtを【出力】フォルダへ書き出す（音声化ツールへの入力用）。
 * - No.順（Q&AのNo.104.5は104と105の間）/ 斜線スキップ行は除外 / 生成台本が空の行は固定台本(FIXED_DAIHON)で補完
 * - Q&Aの行が無い構成シート（古い案件）では、Q&Aの台本は入らない（row が空なのでスキップ扱い）
 * - 進行用の書き込み「（挙手ありがとうございます。）」は読み上げから除去
 */
function exportDaihonTxt_(cfg, data, warnings) {
  try {
    const defs = rowDefs_().slice().sort(function (a, b) { return a.no - b.no; });
    const parts = [];
    defs.forEach(function (def) {
      const row = data.rows[def.no] || {};
      if (SKIP_ROWS[def.no] && isSkipRow_(def, row)) return;
      let t = String(row.daihon || '').trim();
      if (t.indexOf('斜線') !== -1) t = '';
      if (!t && typeof FIXED_DAIHON !== 'undefined' && FIXED_DAIHON[def.no]) t = String(FIXED_DAIHON[def.no]).trim();
      if (!t) return;
      t = t.replace(/（挙手ありがとうございます。?）/g, '').trim();
      if (t) parts.push(t);
    });
    const txt = parts.join('\n\n') + '\n';
    // 再生成時、同名の旧台本txtはゴミ箱へ移す（古いほうを掴む事故を防ぐ）
    const daihonFolder = caseFolder_(cfg, data.name);
    const daihonName = '【台本】' + data.name + '.txt';
    const trashedDaihon = trashSameName_(daihonFolder, daihonName);
    if (trashedDaihon) warnings.push('旧台本txt' + trashedDaihon + '件をゴミ箱へ移しました（30日以内なら復元可）');
    const file = DriveApp.createFile(daihonName, txt, MimeType.PLAIN_TEXT);
    file.moveTo(daihonFolder);
    warnings.push('台本txt(音声化用): ' + file.getUrl());
  } catch (e) { warnings.push('台本txt書き出し失敗: ' + e); }
}

/**
 * フォルダ内の同名ファイルをゴミ箱へ移し、移した件数を返す。
 * 完全削除ではないので30日以内なら復元できる。再生成で同名が並ぶのを防ぐ用。
 */
function trashSameName_(folder, name) {
  let n = 0;
  try {
    const it = filesByName_(folder, name);
    while (it.hasNext()) {
      try { it.next().setTrashed(true); n++; } catch (e) { /* 個別の失敗は無視して続ける */ }
    }
  } catch (e) { /* フォルダが読めない場合も生成自体は続ける */ }
  return n;
}

// ============================================================
// テキスト差し替えエンジン（pilot14の位置認識を実データで駆動）
// ============================================================
function txt_(sh) { try { return sh.getText ? sh.getText().asString().trim() : ''; } catch (e) { return ''; } }
function boxLines_(sh) { return sh.getText().asString().split('\n').map(function (s) { return s.trim(); }).filter(function (s) { return s; }); }
function cs_(slide) {
  const arr = [];
  slide.getShapes().forEach(function (sh) {
    const t = txt_(sh);
    if (!t || /^\d+$/.test(t)) return; // 空・数字円は対象外
    arr.push({ sh: sh, t: t, x: sh.getLeft(), y: sh.getTop(), w: sh.getWidth(), h: sh.getHeight() });
  });
  return arr;
}
function biggest_(slide) { let b = null, l = 0; cs_(slide).forEach(function (o) { if (o.t.length > l) { l = o.t.length; b = o.sh; } }); return b; }
function repl_(slide, needle, s) { if (needle) slide.replaceAllText(needle, s); }
function setWhole_(sh, s) { sh.getText().setText(s); }

/** 全角=1、半角=0.55として行の見かけ幅(em)を概算 */
function emLen_(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    n += /[\x20-\x7E\uFF61-\uFF9F]/.test(s.charAt(i)) ? 0.55 : 1;
  }
  return n;
}

/**
 * ボックス内の各行が幅に収まらない場合、その行のフォントだけ縮小して1行に収める。
 * （混在サイズの行=料金セル等はgetFontSizeがnullになるため自動的に対象外）
 */
function fitBoxText_(sh) {
  try {
    const w = (sh.getWidth() - 14) * 0.93; // インセット分を引いた幅の93%を目標（余白を残す）
    const tr = sh.getText();
    const s = tr.asString();
    // 「見た目の行」単位に分割（通常改行\nと段落内改行\u000bの両方を区切りとして扱う）
    const segs = [];
    let st = 0;
    for (let i = 0; i <= s.length; i++) {
      const ch = i < s.length ? s.charAt(i) : '\n';
      if (ch === '\n' || ch === '\u000b' || ch === '\r') { segs.push([st, i]); st = i + 1; }
    }
    segs.forEach(function (seg) {
      if (seg[1] <= seg[0]) return;
      const line = s.substring(seg[0], seg[1]);
      if (!line.trim()) return;
      const r = tr.getRange(seg[0], seg[1]);
      let size = null;
      try { size = r.getTextStyle().getFontSize(); } catch (e) { }
      if (!size) {
        // 行内にサイズ混在がある場合は最大サイズで判定
        try {
          r.getRuns().forEach(function (run) {
            const rs = run.getTextStyle().getFontSize();
            if (rs && (!size || rs > size)) size = rs;
          });
        } catch (e) { }
      }
      if (!size) return;
      if (emLen_(line) * size * 1.02 > w) {
        // 1文字でもあふれる場合は、収まるサイズまで無条件に縮小（下限10ptのみ）
        let ns = Math.floor(w / (emLen_(line) * 1.02));
        if (ns < 10) ns = 10;
        if (ns < size) r.getTextStyle().setFontSize(ns);
      }
    });
  } catch (e) { }
}

/** 構成テキスト → 表示行（図解タグ・【】除去） */
function klines_(text) {
  return String(text || '')
    .replace(/[【】]/g, '')
    .split('\n')
    .map(function (l) { return l.trim().indexOf('図解（') === 0 ? '' : l.replace(/図解（.*）/g, '').trim(); })
    .filter(function (l) { return l !== ''; });
}

/**
 * ボックス内の既存行を新しい行に置換（行ごとの書式を保持）。
 * 新行が多い場合は末尾を最終行にまとめ、少ない場合は残り既存行を空にする。
 */
function mapBoxLines_(slide, box, newLines) {
  const ls = boxLines_(box);
  if (ls.length === 0) { setWhole_(box, newLines.join('\n')); return; }
  const n = newLines.slice();
  if (n.length > ls.length) { n[ls.length - 1] = n.slice(ls.length - 1).join('\n'); n.length = ls.length; }
  for (let i = 0; i < ls.length; i++) repl_(slide, ls[i], n[i] !== undefined ? n[i] : '');
  trimEdges_(box);
  fitBoxText_(box);
}

/** テキストボックスの先頭・末尾に残った空行（改行）を削除する */
function trimEdges_(box) {
  try {
    const tr = box.getText();
    const s = tr.asString();
    const trimmedEnd = s.replace(/[\s\u3000]+$/, '');
    if (trimmedEnd.length < s.length) tr.getRange(trimmedEnd.length, s.length).clear();
    const s2 = box.getText().asString();
    const lead = s2.length - s2.replace(/^[\s\u3000]*\n/, '').length;
    if (lead > 0) box.getText().getRange(0, lead).clear();
  } catch (e) { /* 空ボックス等は無視 */ }
}

/** サイズ違い2行構成用: カギ括弧と末尾の三点リーダを除去 */
function stripQuotes_(s) {
  return String(s || '').replace(/[「」『』]/g, '').replace(/[…]+$/, '').trim();
}

/** 事例After行→Before行 */
const CASE_PAIR = { 46: 45, 74: 73, 76: 75 };
function rowKlines_(data, no) { const r = data.rows[no]; return r ? klines_(r.kouseiResolved) : []; }

function fillReal_(slide, def, row, data, warnings) {
  const L = def.layout, no = def.no;
  const parts = klines_(row.kouseiResolved);
  const SIMPLE = { cover: 0, question: 0, ideal: 1, text55: 1, rinen: 1, diagram: 1, diagram40: 1, merit_intro: 1, point_intro: 1 };

  if (L === 'cover' || L === 'question' || L.indexOf('static_') === 0) {
    return; // 固定スライド（No.1,2,7,97,101,105,106）: サンプル通り
  }
  if (SIMPLE[L]) {
    if (parts.length === 0) return;
    const box = biggest_(slide);
    if (!box) { warnings.push('No.' + no + ' 差し替えボックス未検出'); return; }
    mapBoxLines_(slide, box, [parts.join('\n')]);
    if (no === 52 || no === 54) { try { box.getText().getTextStyle().setFontFamily('Arial'); } catch (e) { } }

  } else if (L === 'service') { // No.56/79: サービス名45pt+概要14pt(2行のまま)
    const box = biggest_(slide);
    if (!box) { warnings.push('No.' + no + ' 差し替えボックス未検出'); return; }
    mapBoxLines_(slide, box, [parts[0] || '', parts.slice(1).join('\n')]);
    try {
      const ps = box.getText().getParagraphs();
      if (ps[0]) ps[0].getRange().getTextStyle().setFontSize(45);
      for (let i = 1; i < ps.length; i++) ps[i].getRange().getTextStyle().setFontSize(14);
    } catch (e) { }
    fitBoxText_(box);

  } else if (L === 'narrative') {
    const box = biggest_(slide);
    if (!box) { warnings.push('No.' + no + ' 差し替えボックス未検出'); return; }
    const ls = boxLines_(box);
    if (ls.length >= 2) {
      // 1行データ→リード空+メイン / 2行以上→リード+メイン（1行/2行の自動調整）
      mapBoxLines_(slide, box, parts.length >= 2 ? [stripQuotes_(parts[0]), stripQuotes_(parts.slice(1).join(''))] : ['', parts[0] || '']);
    } else if (parts.length >= 2) {
      // 1行ボックスに2行データ: 1行目=リード30pt / 2行目以降=元のメインサイズ
      let mainSize = null;
      try { mainSize = box.getText().getTextStyle().getFontSize(); } catch (e) { }
      setWhole_(box, stripQuotes_(parts[0]) + '\n' + stripQuotes_(parts.slice(1).join('')));
      try {
        const ps = box.getText().getParagraphs();
        ps[0].getRange().getTextStyle().setFontSize(30).setBold(false);
        if (mainSize) for (let i = 1; i < ps.length; i++) ps[i].getRange().getTextStyle().setFontSize(mainSize);
      } catch (e) { }
      fitBoxText_(box);
    } else {
      mapBoxLines_(slide, box, [parts.join('\n')]);
    }

  } else if (L === 'impact') { // No.77/91: 45pt統一
    const box = biggest_(slide);
    if (!box) return;
    setWhole_(box, parts.join('\n'));
    try { box.getText().getTextStyle().setFontSize(45); } catch (e) { }

  } else if (L === 'case') { // タイトル/社長名/業種役職はBefore行から
    const srcNo = CASE_PAIR[no] || no;
    const head = rowKlines_(data, srcNo);
    const title = head[0] || '';
    const role = head[1] || '';
    const name = title.replace(/の事例$/, '');
    const body = (CASE_PAIR[no] ? parts : parts.slice(2)).join('\n');
    const c = cs_(slide);
    const tBox = c.filter(function (o) { return o.y < 100; }).sort(function (a, b) { return a.y - b.y; })[0];
    const bBox = c.filter(function (o) { return o.y >= 340; }).sort(function (a, b) { return b.t.length - a.t.length; })[0];
    const nrBox = c.filter(function (o) { return o.y >= 150 && o.y < 340; }).sort(function (a, b) { return a.y - b.y; })[0];
    if (tBox) { setWhole_(tBox.sh, title); fitBoxText_(tBox.sh); }
    if (bBox) { setWhole_(bBox.sh, body); fitBoxText_(bBox.sh); }
    if (nrBox) {
      const nl = boxLines_(nrBox.sh);
      if (nl.length >= 2) { repl_(slide, nl[0], name); repl_(slide, nl[1], role); }
      else setWhole_(nrBox.sh, name + '\n' + role);
    }

  } else if (L === 'point1' || L === 'point2') {
    slide.getShapes().forEach(function (sh) { // 円内の番号数字はArial
      const t = txt_(sh);
      if (/^\d+$/.test(t)) { try { sh.getText().getTextStyle().setFontFamily('Arial'); } catch (e) { } }
    });
    const c = cs_(slide);
    const obi = c.filter(function (o) { return o.y < 100; }).sort(function (a, b) { return a.y - b.y; })[0];
    const circles = c.filter(function (o) { return o.y >= 200; }).sort(function (a, b) { return a.x - b.x; });
    if (obi) setWhole_(obi.sh, (rowKlines_(data, 50)[0] || ''));
    const texts = L === 'point1' ? [parts.join('\n')] : [rowKlines_(data, 51).join('\n'), parts.join('\n')];
    circles.forEach(function (o, i) { if (texts[i] !== undefined) { setWhole_(o.sh, texts[i]); fitBoxText_(o.sh); } });

  } else if (L === 'merit') { // 開示済みボックス数はサンプル構造のまま累積（42=1,43=2,44=3）
    const c = cs_(slide);
    const obi = c.filter(function (o) { return o.y < 100; }).sort(function (a, b) { return a.y - b.y; })[0];
    const setsu = c.filter(function (o) { return o.y >= 340; }).sort(function (a, b) { return b.t.length - a.t.length; })[0];
    const mid = c.filter(function (o) { return o.y >= 180 && o.y < 330; }).sort(function (a, b) { return a.x - b.x; });
    if (obi) { setWhole_(obi.sh, (rowKlines_(data, 41)[0] || '')); fitBoxText_(obi.sh); }
    if (setsu) { setWhole_(setsu.sh, parts.slice(1).join('\n')); fitBoxText_(setsu.sh); }
    mid.forEach(function (o, i) { setWhole_(o.sh, (rowKlines_(data, 42 + i)[0] || '')); fitBoxText_(o.sh); });

  } else if (L === 'steps') {
    const c = cs_(slide);
    const p1 = c.filter(function (o) { return o.y < 60; }).sort(function (a, b) { return a.y - b.y; })[0];
    const p2 = c.filter(function (o) { return o.y >= 60 && o.y < 200; }).sort(function (a, b) { return a.y - b.y; })[0];
    const st = c.filter(function (o) { return o.y >= 250; }).sort(function (a, b) { return a.x - b.x; });
    if (p1) setWhole_(p1.sh, parts[0] || '');
    if (p2) setWhole_(p2.sh, parts[1] || '');
    st.forEach(function (o, i) { if (parts[i + 2]) { setWhole_(o.sh, parts[i + 2]); fitBoxText_(o.sh); } });

  } else if (L === 'kotei') { // ステップ名はNo.86の3〜6行目
    const steps = rowKlines_(data, 86).slice(2);
    const st = cs_(slide).filter(function (o) { return o.y < 200; }).sort(function (a, b) { return a.x - b.x; });
    st.forEach(function (o, i) { if (steps[i]) { setWhole_(o.sh, steps[i]); fitBoxText_(o.sh); } });

  } else if (L === 'profile') { // [会社名・肩書, 氏名, プロフ文]
    const c = cs_(slide).filter(function (o) { return o.t !== '講師紹介'; });
    const bs = c.filter(function (o) { return o.y >= 100; }).sort(function (a, b) { return a.y - b.y; });
    const vals = [parts[0] || '', parts[1] || '', parts.slice(2).join('\n')];
    bs.forEach(function (o, i) { if (i < 3 && vals[i]) { setWhole_(o.sh, vals[i]); fitBoxText_(o.sh); } });

  } else if (L === 'features') { // No.80/82/84共通（REFはNo.80を参照済み）
    const src = def.km === 'REF' ? rowKlines_(data, def.kRef) : parts;
    const c = cs_(slide);
    const name = c.filter(function (o) { return o.y < 100; }).sort(function (a, b) { return a.y - b.y; })[0];
    if (name) { setWhole_(name.sh, src[0] || ''); fitBoxText_(name.sh); } // タイトルは1行厳守(あふれは縮小)
    const cards = c.filter(function (o) { return o.y >= 240 && (o.t.indexOf('特徴') === 0 || o.t.indexOf('特長') === 0); }).sort(function (a, b) { return a.x - b.x; });
    cards.forEach(function (o, i) {
      const seg = (src[i + 1] || '').split(/[／/|]/);
      const nl = boxLines_(o.sh);
      if (nl.length >= 2 && seg[0]) repl_(slide, nl[1], seg[0].trim());
      if (nl.length >= 3) {
        repl_(slide, nl[2], (seg.slice(1).join('／') || '').trim());
        for (let k = 3; k < nl.length; k++) repl_(slide, nl[k], '');
      }
    });

  } else if (L === 'price') {
    fillPrice_(slide, def, row, data, parts, warnings);

  } else if (L === 'challenge') {
    const box = biggest_(slide);
    if (!box) return;
    mapBoxLines_(slide, box, [parts[0] || '', parts.slice(1).join('\n')]);

  } else if (L === 'invest' || L === 'strike') {
    const src = L === 'strike' ? rowKlines_(data, 92) : parts;
    const box = biggest_(slide);
    if (!box) return;
    mapBoxLines_(slide, box, [src[0] || '', src[1] || '', src.slice(2).join('\n')]);
    try { box.getText().find('万円').forEach(function (r) { r.getTextStyle().setFontSize(50).setBold(true); }); } catch (e) { }
    if (L === 'strike') {
      try {
        box.getText().getParagraphs().forEach(function (p) {
          const s = p.getRange().asString();
          if (s.indexOf('万円') !== -1) p.getRange().getTextStyle().setStrikethrough(true).setForegroundColor('#DD4031');
        });
      } catch (e) { }
      const rm = [];
      slide.getShapes().forEach(function (sh) { try { if (!txt_(sh) && sh.getWidth() > 200 && sh.getHeight() < 20) rm.push(sh); } catch (e) { } });
      rm.forEach(function (sh) { try { sh.remove(); } catch (e) { } });
    }
  }
}

/** 料金スライド（単一=表形式 / 松竹梅=3プランカード。実測構造に基づく位置認識） */
function fillPrice_(slide, def, row, data, parts, warnings) {
  const matsu = data.ryoukin === '松竹梅';
  const els = [];
  slide.getShapes().forEach(function (sh) {
    const t = txt_(sh);
    els.push({ sh: sh, t: t, x: sh.getLeft(), y: sh.getTop(), w: sh.getWidth(), h: sh.getHeight() });
  });

  if (matsu) {
    // No.96のデータを単一ソースに（102はOFF・値引き後・注記のみ持つ）
    const p96 = def.no === 96 ? parts : rowKlines_(data, 96);
    const name = p96[0] || '';
    const planLines = p96.filter(function (l) { return l.indexOf('｜') !== -1 || l.indexOf('|') !== -1; });
    const firstPlan = planLines.length ? p96.indexOf(planLines[0]) : p96.length;
    const descLines = p96.slice(1, firstPlan).filter(function (l) { return l.indexOf('※') !== 0; });
    const desc = descLines.join('\n');
    const plans = planLines.slice(0, 3).map(function (l) { const s = l.split(/[｜|]/); return { name: (s[0]||'').trim(), desc: (s[1]||'').trim(), price: (s[2]||'').trim() }; });
    const note96 = p96.filter(function (l) { return l.indexOf('※') === 0; })[0] || '';

    // ヘッダー（サービス名45pt+概要15ptの同一ボックス）
    const head = els.filter(function (o) { return o.t && o.y < 60 && o.w > 500; }).sort(function (a, b) { return b.t.length - a.t.length; })[0];
    if (head) {
      mapBoxLines_(slide, head.sh, [name, desc]);
      // No.56/79と同じ見た目: サービス名45pt/概要14pt(複数行可)
      try {
        const ps = head.sh.getText().getParagraphs();
        if (ps[0]) ps[0].getRange().getTextStyle().setFontSize(45);
        for (let i = 1; i < ps.length; i++) ps[i].getRange().getTextStyle().setFontSize(14);
      } catch (e) { }
      fitBoxText_(head.sh);
    }
    // 3カラム: プラン名(26pt, y150-230) / 概要(15pt, y230-300) / 定価(y300-320台)
    function col3(yMin, yMax) {
      return els.filter(function (o) { return o.t && o.y >= yMin && o.y < yMax && o.w < 300; }).sort(function (a, b) { return a.x - b.x; });
    }
    const nameBoxes = [];
    col3(150, 230).forEach(function (o, i) {
      if (!plans[i] || !plans[i].name) return;
      const nm = plans[i].name.replace(/プラン$/, '\nプラン'); // 「ライト」＋改行＋「プラン」
      setWhole_(o.sh, nm);
      fitBoxText_(o.sh);
      nameBoxes.push(o.sh);
    });
    // 3プランのプラン名は同一フォントサイズに統一（最小値に合わせる）
    try {
      let minSize = null;
      nameBoxes.forEach(function (sh) {
        sh.getText().getParagraphs().forEach(function (p) {
          const s = p.getRange().getTextStyle().getFontSize();
          if (s && (minSize === null || s < minSize)) minSize = s;
        });
      });
      if (minSize) nameBoxes.forEach(function (sh) { sh.getText().getTextStyle().setFontSize(minSize); });
    } catch (e) { }
    col3(230, 300).forEach(function (o, i) { if (plans[i] && plans[i].desc) { setWhole_(o.sh, plans[i].desc); fitBoxText_(o.sh); } });
    col3(300, 316).forEach(function (o, i) { // 定価ボックス（（税別）行は保持）
      if (!plans[i] || !plans[i].price) return;
      const nl = boxLines_(o.sh);
      if (nl[0]) repl_(slide, nl[0], plans[i].price); else setWhole_(o.sh, plans[i].price);
    });
    if (def.no === 96) {
      const note = els.filter(function (o) { return o.t && o.y >= 400 && o.w > 500; })[0];
      if (note && note96) {
        setWhole_(note.sh, note96);
        try { note.sh.getText().getTextStyle().setFontSize(15).setBold(false); } catch (e) { }
      }
    } else { // No.102: OFFバッジ・値引き後×3・初期費用注記
      const off = parts[0] || '';
      const disc = parts.slice(1, 4);
      const note = parts[4] || ''; // 下部注記(例: 初期費用無料)。5行目に指定があれば表示、なければ空
      const badge = els.filter(function (o) { return o.t && (o.t.indexOf('OFF') >= 0 || o.t.indexOf('Zoom') >= 0); })[0];
      if (badge && off) {
        const nl = boxLines_(badge.sh);
        const target = nl.filter(function (l) { return l.indexOf('OFF') >= 0; })[0];
        if (target) repl_(slide, target, off); // 大文字部は1行(あふれはfitで自動縮小)
        fitBoxText_(badge.sh);
      }
      col3(316, 400).forEach(function (o, i) {
        if (!disc[i]) return;
        // 初月無料の場合は金額を併記せず「初月無料」のみ表示
        const v = disc[i].indexOf('初月無料') >= 0 ? '初月無料' : disc[i];
        setWhole_(o.sh, v);
        fitBoxText_(o.sh);
      });
      // 下部注記: 構成テキスト5行目があれば表示(例: 初期費用無料)、なければ空にする
      const noteBox = els.filter(function (o) { return o.t && o.y >= 400 && o.w < 400; })[0];
      if (noteBox) { setWhole_(noteBox.sh, note); if (note) fitBoxText_(noteBox.sh); }
    }
    fmtPrices_(slide);
    return;
  }

  // ---- 単一（従来の表形式） ----
  const title = els.filter(function (o) { return o.t && o.y < 160 && o.w > 400 && o.t.indexOf('OFF') < 0; }).sort(function (a, b) { return a.y - b.y; })[0];
  const table = els.filter(function (o) { return o.t && o.y >= 150 && o.t !== '通常価格' && o.t.indexOf('OFF') < 0; });
  const plans = table.filter(function (o) { return o.x >= 150 && o.x < 470; }).sort(function (a, b) { return a.y - b.y; });
  const prices = table.filter(function (o) { return o.x >= 470; }).sort(function (a, b) { return a.y - b.y; });
  // 各商品行の価格ボックスは、商品名ボックスとY座標が最も近いものを対応させる（余分な要素によるズレを防止）
  const priceFor = plans.map(function (pl) {
    let best = null, bd = 9999;
    prices.forEach(function (pr) { const d = Math.abs(pr.y - pl.y); if (d < bd) { bd = d; best = pr; } });
    return bd <= 150 ? best : null;
  });
  if (def.no === 102) {
    // 単一の値引き: サービス名・商品名・定価はNo.96と同一。構成テキストは1行目=バッジ表記、2行目以降=「商品名 値引き後表示」
    const p96 = rowKlines_(data, 96);
    if (title) { setWhole_(title.sh, p96[0] || parts[0] || ''); fitBoxText_(title.sh); }
    const names96 = p96.slice(1).map(function (l) { return (l.split(/[\s\u3000]+/)[0] || ''); });
    plans.forEach(function (o, i) { if (names96[i]) setWhole_(o.sh, names96[i]); });
    const pairs96b = p96.slice(1).map(function (l) { const m = l.split(/[\s\u3000]+/); return m.slice(1).join(' ') || ''; });
    const disc = parts.slice(1).map(function (l) {
      const m = l.split(/[\s\u3000]+/);
      return { name: m.length >= 2 ? m[0] : '', disp: m.length >= 2 ? m.slice(1).join(' ') : l };
    });
    // 赤い値引きボックスかどうかを塗り色で判別
    const isRed = function (o) {
      try {
        const f = o.sh.getFill();
        const sf = f && f.getSolidFill();
        const col = sf && sf.getColor();
        if (col && col.getColorType() == SlidesApp.ColorType.RGB) {
          const rc = col.asRgbColor();
          return rc.getRed() > 150 && rc.getGreen() < 130;
        }
      } catch (e) { }
      return false;
    };
    // 各行の値引き表示を事前に決定（商品名で照合、一致しなければ行順）
    const dvFor = plans.map(function (pl, i) {
      let d = disc.filter(function (x) { return x.name && names96[i] && (names96[i].indexOf(x.name) >= 0 || x.name.indexOf(names96[i]) >= 0); })[0];
      if (!d) d = disc[i];
      if (!d || !d.disp) return '';
      return d.disp.indexOf('初月無料') >= 0 ? '初月無料' : d.disp;
    });
    // 各価格ボックスを「最も近い商品行」に1対1で割り当てて書き込む（行間隔に依存しない）
    prices.forEach(function (pr) {
      let bi = -1, bd = 999999;
      plans.forEach(function (pl, i) {
        const d = Math.abs(pr.y - pl.y);
        if (d < bd) { bd = d; bi = i; }
      });
      if (bi < 0) return;
      if (isRed(pr)) {
        if (dvFor[bi]) { setWhole_(pr.sh, dvFor[bi]); fitBoxText_(pr.sh); } // 赤=値引き後
      } else {
        if (pairs96b[bi]) { setWhole_(pr.sh, pairs96b[bi]); fitBoxText_(pr.sh); } // グレー=No.96と同一の定価
      }
    });
    const badge = els.filter(function (o) { return o.t && o.t.indexOf('OFF') >= 0; })[0];
    if (badge && parts[0]) {
      const bt = badge.sh.getText();
      bt.setText('Zoom終了までなら…\n' + parts[0]);
      try { const ps = bt.getParagraphs(); if (ps[1]) ps[1].getRange().getTextStyle().setFontSize(35).setBold(true); } catch (e) { }
      fitBoxText_(badge.sh);
    }
  } else {
    if (title) { setWhole_(title.sh, parts[0] || ''); fitBoxText_(title.sh); }
    const pairs = parts.slice(1).filter(function (l) { return !/^[0-9]+$/.test(l); })
      .map(function (l) { const m = l.split(/[\s\u3000]+/); return { plan: m[0] || '', price: m.slice(1).join(' ') || '' }; });
    plans.forEach(function (o, i) { if (pairs[i] && pairs[i].plan) setWhole_(o.sh, pairs[i].plan); });
    plans.forEach(function (o, i) {
      const target = priceFor[i];
      if (target && pairs[i] && pairs[i].price) { setWhole_(target.sh, pairs[i].price); fitBoxText_(target.sh); }
    });
  }
  fmtPrices_(slide);
}

/** 料金セルの書式: 数字40太字 / 万円30 / （税別）20 */
function fmtPrices_(slide) {
  slide.getShapes().forEach(function (sh) {
    try {
      const t = sh.getText(); const str = t.asString();
      if (str.indexOf('万円') < 0 || str.indexOf('OFF') >= 0 || str.length > 24) return;
      if (str.indexOf('※') >= 0 || str.indexOf('初期費用') >= 0) return; // 注記は15pt非太字のまま
      t.find('万円').forEach(function (r) { r.getTextStyle().setFontSize(30).setBold(false); });
      const nums = str.match(/[0-9,]+(?=万円)/g);
      if (nums) nums.forEach(function (n) { t.find(n).forEach(function (r) { r.getTextStyle().setFontSize(40).setBold(true); }); });
      t.find('（税別）').forEach(function (r) { r.getTextStyle().setFontSize(20).setBold(false); });
      // （税別）以外の括弧補足（例: （6ヶ月契約））は改行して小さく・非太字で表示
      (str.match(/（[^）]*）/g) || []).forEach(function (p) {
        if (p === '（税別）') return;
        try {
          t.replaceAllText(p, '\n' + p);
          t.find(p).forEach(function (r) { r.getTextStyle().setFontSize(15).setBold(false); });
        } catch (e) { }
      });
      fitPriceLine_(sh, nums);
    } catch (e) { }
  });
}

/** 料金セルの見かけ幅（サイズ混在のランを積算し、見た目の行ごとの最大幅を返す） */
function priceLineWidth_(sh) {
  let maxW = 0, cur = 0;
  sh.getText().getRuns().forEach(function (r) {
    const s = r.getTextStyle().getFontSize() || 30;
    const segs = String(r.asString()).split(/[\n\u000b\r]/);
    for (let i = 0; i < segs.length; i++) {
      if (i > 0) { if (cur > maxW) maxW = cur; cur = 0; } // 行の区切り
      cur += emLen_(segs[i]) * s * 1.02;
    }
  });
  if (cur > maxW) maxW = cur;
  return maxW;
}

/**
 * 料金セルの1行維持フィット。あふれる場合の縮小優先順位:
 * ①「月額」を20ptに → ②「万円」を段階的に下げる(下限18pt) → ③金額の数字を段階的に下げる(下限20pt)
 */
function fitPriceLine_(sh, nums) {
  try {
    const usable = (sh.getWidth() - 12) * 0.98;
    const t = sh.getText();
    if (priceLineWidth_(sh) <= usable) return;
    // ① 月額を20ptへ
    t.find('月額').forEach(function (r) { r.getTextStyle().setFontSize(20); });
    // ② 万円を段階縮小
    let s = 30;
    while (priceLineWidth_(sh) > usable && s > 18) {
      s -= 2;
      t.find('万円').forEach(function (r) { r.getTextStyle().setFontSize(s); });
    }
    // ③ 数字を段階縮小
    let ns = 40;
    while (priceLineWidth_(sh) > usable && ns > 20 && nums) {
      ns -= 2;
      nums.forEach(function (n) { t.find(n).forEach(function (r) { r.getTextStyle().setFontSize(ns); }); });
    }
  } catch (e) { }
}

// ============================================================
// よくある質問（Q&A・No.104.5）: No.104の複製に、Q（太字）とA（その下）を最大3組
// ============================================================
/** 構成テキスト → [{q, a}]（最大3組）。「Q. 〜」「A. 〜」の行。番号付き（Q1.）や全角も可 */
function qaPairsFromKousei_(text) {
  const pairs = [];
  let cur = null;
  klines_(text).forEach(function (l) {
    const mq = l.match(/^[QＱ]\s*[0-9０-９]?\s*[\.．:：、]\s*(.*)$/);
    const ma = l.match(/^[AＡ]\s*[0-9０-９]?\s*[\.．:：、]\s*(.*)$/);
    if (mq) { cur = { q: mq[1].trim(), a: '' }; pairs.push(cur); }
    else if (ma && cur) { cur.a = cur.a ? cur.a + '\n' + ma[1].trim() : ma[1].trim(); }
    else if (cur) { if (cur.a) cur.a += '\n' + l; else cur.q += l; }   // 折り返しの続き
    else { cur = { q: l, a: '' }; pairs.push(cur); }                   // Q.の無い書き方は質問として扱う
  });
  return pairs.filter(function (p) { return p.q; }).slice(0, 3);
}

/**
 * 組の数と文字の長さから、文字の大きさを決める（1組=大きく、3組=小さく）。
 * 箱（widthPt × heightPt）に収まらないときは、Q・A・間隔を1ptずつ縮める（答えの下限12pt）。
 * @return {{title:number, q:number, a:number, gap:number}}
 */
function qaLayout_(pairs, widthPt, heightPt) {
  const base = { 1: { q: 30, a: 22, gap: 0 }, 2: { q: 24, a: 18, gap: 16 }, 3: { q: 20, a: 16, gap: 12 } }[Math.max(1, Math.min(3, pairs.length))];
  const lineH = 1.3;
  const need = function (q, a, gap) {
    let h = 0;
    pairs.forEach(function (p, i) {
      if (i > 0) h += gap;
      // fillQa_ と同じく、1組のQ・Aはそれぞれ1段落（改行は空白にして折り返しに任せる）
      h += Math.max(1, Math.ceil(emLen_('Q. ' + p.q.replace(/\n/g, ' ')) * q * 1.02 / widthPt)) * q * lineH;
      if (p.a) h += Math.max(1, Math.ceil(emLen_('A. ' + p.a.replace(/\n/g, ' ')) * a * 1.02 / widthPt)) * a * lineH;
    });
    return h;
  };
  let q = base.q, a = base.a, gap = base.gap;
  for (let n = 0; n < 20 && need(q, a, gap) > heightPt && a > 12; n++) {
    q = Math.max(13, q - 1); a = Math.max(12, a - 1); gap = Math.max(4, gap - 1);
  }
  return { title: 22, q: q, a: a, gap: gap };
}

function fillQa_(slide, row, pres, warnings) {
  const pairs = qaPairsFromKousei_(row.kouseiResolved || row.kousei);
  if (pairs.length === 0) { warnings.push('No.' + QA_NO + '（Q&A）構成テキストから質問を読み取れません'); return; }
  const box = biggest_(slide);
  if (!box) { warnings.push('No.' + QA_NO + '（Q&A）差し替えボックス未検出'); return; }
  const W = pres.getPageWidth(), H = pres.getPageHeight();

  // テンプレ（No.104）の文字色・書体を引き継ぐ（背景が濃い灰色・文字が白）
  let color = '#FFFFFF', family = null;
  try {
    const ts = box.getText().getRange(0, 1).getTextStyle();
    family = ts.getFontFamily();
    const c = ts.getForegroundColor();
    if (c && c.getColorType() == SlidesApp.ColorType.RGB) color = c.asRgbColor().asHexString();
  } catch (e) { }

  // 見出し
  const title = slide.insertTextBox('よくあるご質問', W * 0.08, H * 0.06, W * 0.84, H * 0.12);
  const tt = title.getText();
  tt.getTextStyle().setFontSize(22).setBold(true).setForegroundColor(color);
  if (family) tt.getTextStyle().setFontFamily(family);
  tt.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  try { title.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }

  // 本文（Qは太字、Aはその下に細字）
  const bx = W * 0.1, by = H * 0.2, bw = W * 0.8, bh = H * 0.74;
  box.setLeft(bx).setTop(by).setWidth(bw).setHeight(bh);
  const lay = qaLayout_(pairs, bw - 14, bh - 14);
  const parts = [];   // {text, kind, gapAbove}
  pairs.forEach(function (p, i) {
    parts.push({ text: 'Q. ' + p.q.replace(/\n/g, ' '), kind: 'q', gapAbove: i > 0 });
    if (p.a) parts.push({ text: 'A. ' + p.a.replace(/\n/g, ' '), kind: 'a', gapAbove: false });
  });
  const tr = box.getText();
  tr.setText(parts.map(function (x) { return x.text; }).join('\n'));
  tr.getTextStyle().setForegroundColor(color);
  if (family) tr.getTextStyle().setFontFamily(family);
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START).setSpaceAbove(0).setSpaceBelow(0);
  let pos = 0;
  parts.forEach(function (x) {
    const r = tr.getRange(pos, pos + x.text.length);
    r.getTextStyle().setFontSize(x.kind === 'q' ? lay.q : lay.a).setBold(x.kind === 'q');
    if (x.gapAbove) r.getParagraphStyle().setSpaceAbove(lay.gap);
    else if (x.kind === 'a') r.getParagraphStyle().setSpaceAbove(Math.round(lay.a * 0.3));
    pos += x.text.length + 1;
  });
  try { box.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) { }
}

// ============================================================
// スキップ（斜線）: テンプレ文言に戻してスキップ設定
// ============================================================
function isSkipRow_(def, row) {
  if (!SKIP_ROWS[def.no]) return false;
  const k = String(row.kousei || '').trim();
  return k === '' || k.indexOf('斜線') !== -1;
}
function fillSkipSample_(slide, def) {
  const box = biggest_(slide);
  if (box) setWhole_(box, String(SKIP_ROWS[def.no]));
}

// ============================================================
// 画像差し替え（Image.replace: 位置・サイズは継承 / 不透明度はリセットされるため
// OPACITY_RULES に基づく白オーバーレイで再現）
// ============================================================
function slidePics_(slide) {
  return slide.getImages().map(function (im) {
    return { im: im, x: im.getLeft(), y: im.getTop(), w: im.getWidth(), h: im.getHeight() };
  });
}
function whiteOverlayOver_(slide, x, y, w, h, alpha) {
  const ov = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  ov.getFill().setSolidFill('#FFFFFF', alpha);
  ov.getBorder().setTransparent();
  return ov;
}
/** 「ファイル名@25」形式 → {name, opacity(0-100)|null} */
function parseAssetRef_(raw) {
  const m = String(raw || '').trim().match(/^(.*?)(?:@(\d{1,3}))?$/);
  return { name: (m ? m[1] : '').trim(), opacity: m && m[2] ? Math.min(100, Number(m[2])) : null };
}

function replaceImages_(slide, def, row, data, zukaiFolder, idx, targets, warnings) {
  const no = def.no;
  const pics = slidePics_(slide);
  const fullbleed = pics.filter(function (p) { return p.w >= 700 && p.h >= 430; })[0];
  const nonFull = pics.filter(function (p) { return !(p.w >= 700 && p.h >= 430) && !(p.w >= 700 && p.h < 200); }); // 帯画像は固定

  // 1. 図解（案件別フォルダ・キーワード一致）: 帯以外で最大の画像を差し替え
  if (def.img) {
    const file = findZukai_(zukaiFolder, def, row);
    if (!file) { warnings.push('No.' + no + ' 図解未発見: ' + (def.img || '')); }
    else {
      const target = (def.layout === 'cover' || def.layout === 'profile')
        ? (fullbleed || nonFull.sort(function (a, b) { return b.w * b.h - a.w * a.h; })[0])
        : nonFull.sort(function (a, b) { return b.w * b.h - a.w * a.h; })[0];
      if (target) { try { target.im.replace(file.getBlob()); } catch (e) { warnings.push('No.' + no + ' 図解replace失敗: ' + e); } }
    }
  }

  // 2. アイコン（紫/白）: 60〜110pt級の画像を左から順に差し替え
  const iconPics = pics.filter(function (p) { return p.w >= 55 && p.w <= 115 && p.h >= 55 && p.h <= 115; })
    .sort(function (a, b) { return a.x - b.x; });
  if (def.layout === 'merit') {
    iconPics.forEach(function (p, i) { replaceIcon_(p, (data.rows[42 + i] || {}).icon, idx.ICON_P, warnings, no); });
  } else if (def.layout === 'point1') {
    if (iconPics[0]) replaceIcon_(iconPics[0], (data.rows[51] || {}).icon, idx.ICON_P, warnings, no);
  } else if (def.layout === 'point2') {
    [(data.rows[51] || {}).icon, (data.rows[53] || {}).icon].forEach(function (nm, i) { if (iconPics[i]) replaceIcon_(iconPics[i], nm, idx.ICON_P, warnings, no); });
  } else if (def.layout === 'features') {
    const names = String((data.rows[80] || {}).icon || '').split(/[／/|]/);
    iconPics.forEach(function (p, i) { replaceIcon_(p, names[i], idx.ICON_W, warnings, no); });
  } else if (def.layout === 'steps' || def.layout === 'kotei') {
    iconPics.forEach(function (p, i) { replaceIcon_(p, (data.rows[87 + i] || {}).icon, idx.ICON_P, warnings, no); });
    // 工程: 非現在ステップの減光。アイコン背後の白円と同サイズの円を重ねる（実測: 20% → 白80%）
    if (def.layout === 'kotei') {
      const cur = no - 87;
      const circles = [];
      slide.getShapes().forEach(function (sh) {
        try {
          if (sh.getShapeType() === SlidesApp.ShapeType.ELLIPSE && sh.getWidth() >= 60 && sh.getWidth() <= 180) circles.push(sh);
        } catch (e) { }
      });
      iconPics.forEach(function (p, i) {
        if (i === cur) return;
        let g = null;
        circles.forEach(function (e) {
          const cx = e.getLeft() + e.getWidth() / 2, cy = e.getTop() + e.getHeight() / 2;
          if (Math.abs(cx - (p.x + p.w / 2)) < 25 && Math.abs(cy - (p.y + p.h / 2)) < 50) g = e;
        });
        const x = g ? g.getLeft() : p.x - 10, y = g ? g.getTop() : p.y - 10;
        const w = g ? g.getWidth() : p.w + 20, h = g ? g.getHeight() : p.h + 20;
        const ov = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y, w, h);
        ov.getFill().setSolidFill('#FFFFFF', 1 - OPACITY_RULES.stepInactive / 100);
        ov.getBorder().setTransparent();
      });
    }
  }

  // 3. 背景（選定行のみ）: 全面画像をreplace + 実測不透明度を白オーバーレイで再現
  let bgIndex = null;
  if (targets.bg.indexOf(no) !== -1) bgIndex = idx['BG' + bgOpacityFor_(no)]; // No.→濃度対応表で選定元を決定
  else if (targets.select.indexOf(no) !== -1) bgIndex = idx.SELECT;
  if (bgIndex && fullbleed) {
    const ref = parseAssetRef_(row.bg);
    if (ref.name) {
      const file = lookupAsset_(bgIndex, ref.name);
      if (!file) { warnings.push('No.' + no + ' 背景未発見: ' + ref.name + '（濃度' + (targets.select.indexOf(no) !== -1 ? '選択' : bgOpacityFor_(no)) + 'フォルダ内）'); }
      else {
        try {
          fullbleed.im.replace(file.getBlob());
          // 不透明度は焼き込み済み素材（〜_白/〜_黒/〜_赤）で表現するのが基本。
          // 「ファイル名@NN」指定時のみ白オーバーレイで減光（応急用）。
          if (ref.opacity !== null) {
            const ov = whiteOverlayOver_(slide, 0, 0, 720, 450, 1 - ref.opacity / 100);
            ov.sendToBack();
            fullbleed.im.sendToBack();
          }
        } catch (e) { warnings.push('No.' + no + ' 背景replace失敗: ' + e); }
      }
    }
  }
}

function replaceIcon_(pic, rawName, index, warnings, no) {
  const ref = parseAssetRef_(rawName);
  if (!ref.name) return; // 未指定はサンプルのまま
  const file = lookupAsset_(index, ref.name);
  if (!file) { warnings.push('No.' + no + ' アイコン未発見: ' + ref.name); return; }
  try { pic.im.replace(file.getBlob()); } catch (e) { warnings.push('No.' + no + ' アイコンreplace失敗: ' + e); }
}

function findZukai_(folder, def, row) {
  let keyword = def.img;
  const m = String(row.kouseiResolved || '').match(/図解（([^）]+)）/);
  if (m) keyword = m[1];
  const partialOnly = def.img === '*プロフ';
  const key = normalizeKey_(partialOnly ? 'プロフ' : keyword);
  // ★GIFは使わない。図解はPNGで統一し、GIFはそもそも作らない運用になった（2026-08-12）。
  //   古い案件のフォルダにGIFが残っていても拾わない。
  //   PNGが1枚も無ければ見つからない扱いになり、管理シートF列に「図解未発見」が出る。
  //   黙って動く画像が入るより、見つからないと言われるほうがよい。
  let exact = null, partial = null;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const mime = f.getMimeType();
    if (mime.indexOf('image/') !== 0) continue;
    if (mime === 'image/gif') continue;
    const nk = normalizeKey_(f.getName());
    if (!partialOnly && nk === key) { exact = f; break; }
    if (nk.indexOf(key) !== -1 && !partial) partial = f;
  }
  return exact || partial;
}

// ============================================================
// ノート・素材フォルダ・索引（既存互換）
// ============================================================
function setNotesText_(slide, text) {
  const sh = slide.getNotesPage().getSpeakerNotesShape();
  if (sh) sh.getText().setText(text);
}
function resolveAssetsFolder_(cfg, data, warnings) {
  if (data.assetsFolderId) {
    try { return DriveApp.getFolderById(data.assetsFolderId); } catch (e) { /* fallthrough */ }
  }
  const root = DriveApp.getFolderById(cfg.ASSETS_FOLDER_ID);
  const sub = foldersByName_(root, data.name);
  if (sub.hasNext()) return sub.next();
  // ★ここで黙って素材ルートを返すと、図解が1枚も入らないまま
  //   「テンプレートのまま（＝別会社のロゴ入り）」のスライドが完成してしまう。
  //   図解未発見が全行に出るだけでは気づきにくいので、先頭で名指しする。
  if (warnings) {
    warnings.push('★素材フォルダが見つかりません: 【素材】投影資料/' + data.name +
      '（図解は1枚も入りません。フォルダ名が構成シートB1と同じか確認してください）');
  }
  return root;
}
function resolveNamedFolder_(cfg, key) {
  const pid = PropertiesService.getScriptProperties().getProperty(ASSET_PROP_KEYS[key]);
  if (pid) { try { return DriveApp.getFolderById(pid); } catch (e) { /* next */ } }
  try {
    const it = foldersByName_(DriveApp.getFolderById(cfg.ASSETS_FOLDER_ID), ASSET_SUB[key]);
    if (it.hasNext()) return it.next();
  } catch (e) { /* next */ }
  try { return DriveApp.getFolderById(DEFAULT_ASSET_IDS[key]); } catch (e) { return null; }
}
/** 背景画像フォルダ内の濃度サブフォルダ（85/25/15）の画像索引 */
function buildBgIndex_(cfg, opacity) {
  const index = {};
  const root = resolveNamedFolder_(cfg, 'BG');
  if (!root) return index;
  const it = foldersByName_(root, opacity);
  const folder = it.hasNext() ? it.next() : root; // サブフォルダ未作成時は直下を見る（暫定互換）
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType().indexOf('image/') === 0) index[normalizeKey_(f.getName())] = f;
  }
  return index;
}

function buildAssetIndex_(cfg, key) {
  const index = {};
  const folder = resolveNamedFolder_(cfg, key);
  if (!folder) return index;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType().indexOf('image/') === 0) index[normalizeKey_(f.getName())] = f;
  }
  return index;
}
function lookupAsset_(index, name) {
  if (!name) return null;
  const key = normalizeKey_(name);
  if (index[key]) return index[key];
  const keys = Object.keys(index);
  for (let i = 0; i < keys.length; i++) {
    if (keys[i].indexOf(key) !== -1 || key.indexOf(keys[i]) !== -1) return index[keys[i]];
  }
  return null;
}
function normalizeKey_(s) {
  const circled = { '①': '1', '②': '2', '③': '3', '④': '4', '⑤': '5', '⑥': '6' };
  let t = String(s);
  if (t.normalize) t = t.normalize('NFC'); // Macアップロード由来の濁点・半濁点分解（NFD）を合成形式に統一
  return t
    .replace(/\.[A-Za-z0-9]+$/, '')
    .replace(/[①-⑥]/g, function (c) { return circled[c]; })
    .replace(/[（(]/g, '(').replace(/[）)]/g, ')')
    .replace(/[\s\u3000_-]+/g, '')
    .toLowerCase();
}
