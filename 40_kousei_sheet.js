/**
 * 40_kousei_sheet.gs（アイコン/背景対応版）
 * 構成シートの雛形生成・AI出力の取り込み・読み戻し
 *
 * タブ構成:
 *  設定           : B1=案件名, B2=承認, B3=ステータス, B4=ヒアリングID, B5=素材フォルダID, B6=プロンプトDocURL
 *  AI出力貼り付け  : B2=取り込みチェック, A4〜=台本JSON, B4〜=構成テキストJSON, C4〜=デザインJSON(任意)
 *  コンテンツ4タブ : A=No. B=基本構成 C=小分類 D=構成テキスト E=台本 F=アイコン G=背景
 */

const PASTE_TAB = 'AI出力貼り付け';
const PASTE_ROWS = 200;

function createKouseiScaffold_(projectName, hearingId, promptDocUrl) {
  const cfg = getConfig_();
  const ss = SpreadsheetApp.create('【構成シート】' + projectName);
  DriveApp.getFileById(ss.getId()).moveTo(caseFolder_(cfg, projectName));

  // ---- 設定タブ ----
  const st = ss.getSheets()[0];
  st.setName(TABS.SETTINGS);
  st.getRange('A1:A6').setValues([['案件名'], ['承認（済にするとスライド生成）'], ['ステータス'], ['ヒアリングファイルID'], ['素材フォルダID（空なら既定）'], ['プロンプトDoc']]);
  st.getRange('B1').setValue(projectName);
  st.getRange('B2').insertCheckboxes();
  st.getRange('B3').setValue('AI出力待ち');
  st.getRange('B4').setValue(hearingId);
  st.getRange('B6').setValue(promptDocUrl || '');

  // ---- 修正依頼の入口（Claude Codeを触らずにLPとスライドを直すため） ----
  // ★ここで文言を書き直さないこと。30_prompts.js の ensure〜 が正本で、
  //   既存シートの自己修復にも同じ関数が使われる。2か所に書くと必ず片方が古くなる。
  ensureLpFixCells_(st);        // A8:A10
  ensureSlideFixCells_(st);     // A12:A15
  ensureSlideRegenCell_(st);    // A16
  ensureTitleConfirmCells_(st); // A17:A19
  ensureDeliverCells_(st);      // A21:A23

  // A列は説明文がそのまま読める幅に（240だと「↑の原因をAIに調べてもらう（10〜35分…」で切れて意味が変わる）
  st.setColumnWidth(1, 500).setColumnWidth(2, 480);

  // ---- AI出力貼り付けタブ ----
  const pt = ss.insertSheet(PASTE_TAB);
  pt.getRange('A1').setValue(
    '使い方: 設定タブB6のプロンプトDocを開き全文コピー → Claude Codeに貼って実行すると、' +
    '案件フォルダに【台本JSON】【構成テキストJSON】【デザインJSON】のtxtが保存され、' +
    '最大5分で各タブへ自動反映されます（このタブへの貼り付けは不要）。 ' +
    '手で貼りたいときだけ: 台本JSONをA4以下、構成テキストJSONをB4以下、デザインJSONをC4以下に貼り' +
    '（長い場合は下のセルに分割可・上から連結されます）、B2の取り込みチェックをONにしてください');
  pt.getRange('A1:E1').merge().setWrap(true).setBackground('#FFF3CD');
  pt.setRowHeight(1, 90);
  pt.getRange('A2').setValue('取り込み実行 →');
  pt.getRange('B2').insertCheckboxes();
  pt.getRange('D2').setValue('結果:');
  pt.getRange('A3').setValue('① 台本JSON').setFontWeight('bold').setBackground('#efefef');
  pt.getRange('B3').setValue('② 構成テキストJSON').setFontWeight('bold').setBackground('#efefef');
  pt.getRange('C3').setValue('③ デザインJSON（任意）').setFontWeight('bold').setBackground('#efefef');
  pt.setColumnWidth(1, 380).setColumnWidth(2, 380).setColumnWidth(3, 300).setColumnWidth(4, 60).setColumnWidth(5, 380);
  pt.setFrozenRows(3);

  writeContentTabs_(ss, {}, {}, {});
  return ss;
}

/** 4タブへ全106行（＋Q&AのNo.104.5）を書き込む */
function writeContentTabs_(ss, kouseiMap, daihonMap, designMap, taikeiMap) {
  designMap = designMap || {};
  taikeiMap = taikeiMap || {};
  const defs = rowDefs_();
  const byTab = {};
  defs.forEach(function (r) { (byTab[r.tab] = byTab[r.tab] || []).push(r); });
  [TABS.INTRO, TABS.PROBLEM, TABS.STORY, TABS.OFFER].forEach(function (tabName) {
    let sh = ss.getSheetByName(tabName);
    if (!sh) sh = ss.insertSheet(tabName);
    const rows = byTab[tabName];
    const values = [['No.', '基本構成', '小分類', '構成テキスト', '台本', 'アイコン', '背景']];
    rows.forEach(function (r) {
      const dz = designMap[r.no] || {};
      let k = resolveKousei_(r, kouseiMap);
      let dh = resolveDaihon_(r, daihonMap);
      const imported = Object.keys(kouseiMap).length > 0;
      if (SKIP_ROWS[r.no] && imported && String(k).trim() === '' && String(dh).trim() === '') {
        k = '（斜線）'; dh = '（斜線: このスライドはテンプレのままスキップ設定で出力されます）';
      }
      // No.96/102 は「（単一）/（松竹梅）」の2行構造（実物の構成シートと同形式）
      if (r.no === 96 || r.no === 102) {
        const base = r.no === 96 ? '定価' : '値引き後の金額';
        const tk = taikeiMap[r.no] || '';
        const single = tk !== '松竹梅';
        values.push([r.no, base + '（単一）', r.sub, single ? k : (imported ? '（斜線）' : ''), single ? dh : (imported ? '（斜線）' : ''), '', dz.bg || '']);
        values.push(['', base + '（松竹梅）', r.sub, !single ? k : (imported ? '（斜線）' : ''), !single ? dh : (imported ? '（斜線）' : ''), '', '']);
        return;
      }
      values.push([r.no, r.label, r.sub, k, dh, r.icon ? (dz.icon || '') : '', dz.bg || '']);
    });
    sh.getRange(1, 1, values.length, 7).setValues(values);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#efefef');
    sh.setColumnWidth(1, 50).setColumnWidth(2, 160).setColumnWidth(3, 160).setColumnWidth(4, 300).setColumnWidth(5, 460).setColumnWidth(6, 140).setColumnWidth(7, 140);
    sh.getRange(2, 4, values.length - 1, 2).setWrap(true).setVerticalAlignment('top');
    sh.setFrozenRows(1);
  });
}

function resolveKousei_(def, kouseiMap) {
  switch (def.km) {
    case 'FIX': return def.kFix;
    case 'EMPTY': return '';
    case 'REF': return 'No.' + def.kRef + 'と同じ';
    default: return kouseiMap[def.no] !== undefined ? kouseiMap[def.no] : '';
  }
}
function resolveDaihon_(def, daihonMap) {
  if (def.dm === 'FIX') return FIXED_DAIHON[def.no] || '';
  return daihonMap[def.no] !== undefined ? daihonMap[def.no] : '';
}

// ============================================================
// AI出力の取り込み
// ============================================================
function isImportRequested_(ss) {
  const pt = ss.getSheetByName(PASTE_TAB);
  return pt ? pt.getRange('B2').getValue() === true : false;
}
function readPasteColumn_(pt, col) {
  return pt.getRange(4, col, PASTE_ROWS, 1).getValues()
    .map(function (r) { return String(r[0]); })
    .filter(function (v) { return v.trim() !== ''; })
    .join('\n');
}

/**
 * 貼り付けタブ（手貼り）または案件フォルダのtxt（Claude Codeの書き出し）を取り込む。
 * texts を渡した場合はそちらを使い、貼り付けタブは読まない。
 */
function importAiOutput_(ss, texts) {
  const pt = ss.getSheetByName(PASTE_TAB);
  if (!pt) throw new Error(PASTE_TAB + ' タブがありません');
  const daihonText = texts ? texts.daihon : readPasteColumn_(pt, 1);
  const kouseiText = texts ? texts.kousei : readPasteColumn_(pt, 2);
  const designText = texts ? texts.design : readPasteColumn_(pt, 3); // 任意
  if (!daihonText) throw new Error('台本JSONがありません（【台本JSON】のtxt、またはA4以下）');
  if (!kouseiText) throw new Error('構成テキストJSONがありません（【構成テキストJSON】のtxt、またはB4以下）');

  const warnings = [];
  const gen = daihonGenRows_();
  const daihonMap = validateRows_(parseJsonArray_(daihonText), gen.partA.concat(gen.partB), 'daihon', warnings);
  const kouseiRows = parseJsonArray_(kouseiText);
  const kouseiMap = validateRows_(kouseiRows, kouseiGenRows_(), 'kousei_text', warnings);
  dropQaIfHearingEmpty_(ss, daihonMap, kouseiMap, warnings);
  const taikeiMap = {};
  kouseiRows.forEach(function (r) {
    const no = Number(r.no);
    if ((no === 96 || no === 102) && r.taikei) taikeiMap[no] = String(r.taikei);
  });
  if (!taikeiMap[96]) { taikeiMap[96] = '単一'; warnings.push('No.96 taikei未指定→単一とみなしました'); }
  if (!taikeiMap[102]) taikeiMap[102] = taikeiMap[96];
  if (taikeiMap[96] !== taikeiMap[102]) warnings.push('No.96とNo.102のtaikeiが不一致です（96に合わせて生成されます）');
  checkKouseiLimits_(kouseiMap, warnings); // 文字数規定の照合（警告のみ）

  const designMap = {};
  if (designText) {
    try {
      parseJsonArray_(designText).forEach(function (r) {
        const no = Number(r.no);
        if (!isNaN(no)) designMap[no] = { icon: String(r.icon || ''), bg: String(r.bg || '') };
      });
    } catch (e) { warnings.push('デザインJSONの解析に失敗（アイコン/背景は未設定で続行）: ' + e); }
  }

  writeContentTabs_(ss, kouseiMap, daihonMap, designMap, taikeiMap);
  return warnings;
}

/**
 * ★ヒアリングシートにQ&Aの記入が無いのに、AIが No.104.5 を書いてきたら捨てる。
 *   「記入が無ければ何も入らない」を、AIの出来に頼らずここで守る（作り話をお客様の前に出さない）。
 *   ヒアリングシートが読めないときは、AIの出力をそのまま使い、警告だけ残す。
 */
function dropQaIfHearingEmpty_(ss, daihonMap, kouseiMap, warnings) {
  const has = String(daihonMap[QA_NO] || '').trim() !== '' || String(kouseiMap[QA_NO] || '').trim() !== '';
  if (!has) return;
  try {
    const st = ss.getSheetByName(TABS.SETTINGS);
    const hid = st ? String(st.getRange('B4').getValue() || '').trim() : '';
    if (!hid) { warnings.push('No.' + QA_NO + '（Q&A）: ヒアリングシートのIDが設定タブB4に無く、記入の有無を確かめられませんでした'); return; }
    if (readHearingQa_(hid).length > 0) return;
    daihonMap[QA_NO] = '';
    kouseiMap[QA_NO] = '';
    warnings.push('No.' + QA_NO + '（Q&A）: ヒアリングシートにQ&Aの記入が無いので、AIの出力を捨てました');
  } catch (e) {
    warnings.push('No.' + QA_NO + '（Q&A）: ヒアリングシートを読めず、記入の有無を確かめられませんでした（' + e + '）');
  }
}

/**
 * 案件フォルダに置かれたJSONのtxt（Claude Codeが書き出したもの）を読む。
 * 台本と構成テキストの2つがそろっているときだけ取り込み対象とする。
 * stamp はファイルの最終更新時刻。前回と同じなら取り込み済みとみなす。
 */
function readCaseJsonTexts_(cfg, projectName) {
  const want = { daihon: '【台本JSON】', kousei: '【構成テキストJSON】', design: '【デザインJSON】' };
  let folder;
  folder = caseFolderIfExists_(cfg, projectName); if (!folder) return null;
  const found = {};
  let stamp = 0;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const nm = f.getName();
    if (!/\.txt$/i.test(nm)) continue;
    Object.keys(want).forEach(function (k) {
      if (found[k] || nm.indexOf(want[k]) !== 0) return;
      let text = '';
      try { text = f.getBlob().getDataAsString('UTF-8'); } catch (e) { return; }
      if (!String(text).trim()) return;
      found[k] = text;
      const t = f.getLastUpdated().getTime();
      if (t > stamp) stamp = t;
    });
  }
  if (!found.daihon || !found.kousei) return null;
  return { daihon: found.daihon, kousei: found.kousei, design: found.design || '', stamp: String(stamp) };
}
/** txtから取り込んだときの目印（同じ内容で何度も取り込まないため） */
function getAutoStamp_(ssId) {
  return PropertiesService.getScriptProperties().getProperty('AUTOIMPORT_' + ssId) || '';
}
function setAutoStamp_(ssId, stamp) {
  PropertiesService.getScriptProperties().setProperty('AUTOIMPORT_' + ssId, String(stamp));
}
function setImportResult_(ss, message) {
  const pt = ss.getSheetByName(PASTE_TAB);
  if (!pt) return;
  pt.getRange('B2').setValue(false);
  pt.getRange('E2').setValue(message).setWrap(true);
}

// ============================================================
// 読み戻し・承認・ステータス
// ============================================================
function readKouseiSheet_(kouseiSheetId) {
  const ss = SpreadsheetApp.openById(kouseiSheetId);
  const st = ss.getSheetByName(TABS.SETTINGS);
  const out = {
    name: String(st.getRange('B1').getValue()),
    assetsFolderId: String(st.getRange('B5').getValue()).trim(),
    ryoukin: '単一',
    rows: {}
  };
  [TABS.INTRO, TABS.PROBLEM, TABS.STORY, TABS.OFFER].forEach(function (tabName) {
    const sh = ss.getSheetByName(tabName);
    if (!sh) throw new Error('構成シートにタブがありません: ' + tabName);
    const values = sh.getDataRange().getValues();
    let lastNo = 0;
    for (let i = 1; i < values.length; i++) {
      const noCell = Number(values[i][0]);
      const label = String(values[i][1] || '');
      const no = noCell || ((lastNo === 96 || lastNo === 102) && label ? lastNo : 0);
      if (!no) continue;
      lastNo = noCell || lastNo;
      const rec = {
        kousei: String(values[i][3]), daihon: String(values[i][4]),
        icon: String(values[i][5] || ''), bg: String(values[i][6] || '')
      };
      const isSlash = rec.kousei.trim() === '' || rec.kousei.indexOf('斜線') !== -1;
      if ((no === 96 || no === 102) && out.rows[no]) {
        // 2行目（松竹梅側）: 記入がある側を採用
        if (!isSlash) { out.rows[no] = rec; if (label.indexOf('松竹梅') !== -1) out.ryoukin = '松竹梅'; }
        continue;
      }
      if ((no === 96 || no === 102) && !isSlash && label.indexOf('松竹梅') !== -1) out.ryoukin = '松竹梅';
      out.rows[no] = rec;
    }
  });
  rowDefs_().forEach(function (def) {
    const row = out.rows[def.no];
    if (!row) return;
    const m = row.kousei.match(/^No\.?(\d+)と同じ$/);
    row.kouseiResolved = (m && out.rows[Number(m[1])]) ? out.rows[Number(m[1])].kousei : row.kousei;
  });
  return out;
}

function isApproved_(kouseiSheetId) {
  return SpreadsheetApp.openById(kouseiSheetId).getSheetByName(TABS.SETTINGS).getRange('B2').getValue() === true;
}
function setKouseiStatus_(kouseiSheetId, status) {
  SpreadsheetApp.openById(kouseiSheetId).getSheetByName(TABS.SETTINGS).getRange('B3').setValue(status);
}