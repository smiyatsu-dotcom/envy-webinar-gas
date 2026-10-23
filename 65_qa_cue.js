// Q&Aカンペ: FAQを補完せず、出力フォルダの1つのGoogle Docへ最新化する。
function qaCueFileName_(company) { return '【Q&Aカンペ】' + company; }
function qaCuePdfFileName_(company) { return '【Q&Aカンペ】' + company + '.pdf'; }
function qaCuePairsFromValues_(values) {
  const out = [], cells = r => (r || []).map(x => String(x == null ? '' : x).trim());
  const label = (x, s) => x.replace(/[:：\s]/g, '') === s;
  let start = -1;
  for (let i = 0; i < values.length; i++) if (cells(values[i]).some(x => /よくある質問[（(]\s*Q\s*[&＆]\s*A\s*[）)]/i.test(x))) { start = i + 1; break; }
  let q = null;
  for (let i = start; i >= 0 && i < values.length && out.length < 3; i++) {
    const row = cells(values[i]), non = row.filter(Boolean);
    if (non.length && /^([⒈-⒛]|[0-9０-９]+[\.．])\s*/.test(non[0]) && !label(non[0], 'よくある質問')) break;
    const qi = row.findIndex(x => label(x, 'よくある質問')), ai = row.findIndex(x => label(x, '記入欄'));
    if (qi >= 0) q = row.slice(qi + 1).filter(Boolean).join('\n');
    if (ai >= 0 && q !== null) { out.push({ q: q, a: row.slice(ai + 1).filter(Boolean).join('\n') }); q = null; }
  }
  while (out.length < 3) out.push({ q: '', a: '' });
  return out;
}
function qaCueText_(pairs) {
  const lines = ['Q&Aカンペ', '', '司会と登壇者が共通で利用する確認用です。画面共有はしません。', 'ヒアリングシートの確定FAQをそのまま記載しています。', ''];
  pairs.forEach((p, i) => { lines.push('FAQ ' + (i + 1), '司会が読む質問：' + (p.q || '未記入'), '登壇者の回答：' + (p.a || '未記入'), ''); });
  return lines.join('\n');
}
function syncQaCue_(cfg, projectName) {
  const src = caseFolderIfExists_(cfg, projectName); if (!src) throw new Error('【出力】生成物の案件フォルダがありません');
  const row = mgmtSheet_(cfg).getDataRange().getValues().slice(1).filter(r => String(r[1]) === String(projectName))[0];
  if (!row || !row[0]) throw new Error('ヒアリングシートIDが見つかりません');
  const ss = SpreadsheetApp.openById(String(row[0]));
  const offer = ss.getSheets().filter(s => s.getName().trim() === HEARING_TABS.OFFER.trim())[0];
  if (!offer) throw new Error('ヒアリングシートにオファー記入用タブがありません');
  const name = qaCueFileName_(caseCompany_(projectName)), it = filesByName_(src, name);
  let file;
  if (it.hasNext()) file = it.next();
  else {
    const created = DocumentApp.create(name);
    file = DriveApp.getFileById(created.getId());
    file.moveTo(src);
  }
  const doc = DocumentApp.openById(file.getId()); doc.getBody().setText(qaCueText_(qaCuePairsFromValues_(offer.getDataRange().getValues()))); doc.saveAndClose();
  // 同じ内容をPDFにもする。既存版だけを置換し、共有範囲は変更しない。
  const oldPdf = filesByName_(src, qaCuePdfFileName_(caseCompany_(projectName)));
  while (oldPdf.hasNext()) oldPdf.next().setTrashed(true);
  const pdf = file.getAs(MimeType.PDF).setName(qaCuePdfFileName_(caseCompany_(projectName)));
  src.createFile(pdf);
  return file;
}
function syncAllQaCues_() {
  const cfg = getConfig_();
  mgmtSheet_(cfg).getDataRange().getValues().slice(1).forEach(r => {
    // 生成物フォルダがある案件だけ。まだ生成を始めていない案件を作らない。
    if (r[0] && r[1] && caseFolderIfExists_(cfg, String(r[1]))) try { syncQaCue_(cfg, String(r[1])); } catch (e) { Logger.log('Q&Aカンペを更新できません: ' + r[1] + ' / ' + e); }
  });
}
