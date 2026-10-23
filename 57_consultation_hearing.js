/**
 * PM相談室が「ヒアリング原本」と呼べるのは、案件の管理行から解決した
 * ヒアリングシートだけです。出力物から推測してはいけません。
 *
 * このスナップショットは同一案件の【出力】フォルダに置く内部の読取用資料で、
 * 作成・更新に既存の Drive/Spreadsheet 権限以外は使いません。原本の値が
 * 変わらない限り同じ revision を保つため、Bridge は承認前に版違いを検出できます。
 */
const 相談室_原本ファイル名_ = '【相談室】ヒアリング原本.json';
const 相談室_原本最大文字数_ = 180000;

function 相談室_ハッシュ_(text) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return digest.map(function (b) { const n = b < 0 ? b + 256 : b; return ('0' + n.toString(16)).slice(-2); }).join('');
}

function 相談室_原本テキスト_(ss) {
  const out = [];
  ss.getSheets().forEach(function (sheet) {
    out.push('【' + sheet.getName() + '】');
    const rows = sheet.getDataRange().getValues();
    rows.forEach(function (row) {
      const line = row.map(function (value) { return String(value == null ? '' : value).replace(/[\t\r\n]+/g, ' ').trim(); }).join('\t').replace(/\t+$/g, '');
      if (line) out.push(line);
    });
  });
  const text = out.join('\n');
  if (!text) throw new Error('ヒアリング原本に読める内容がありません');
  if (text.length > 相談室_原本最大文字数_) throw new Error('ヒアリング原本が相談室の安全な読取上限を超えています');
  return text;
}

/** 管理行で解決済みの projectName/hearingId だけを受け取る。任意IDは受け取らない。 */
function 相談室_原本スナップショットを書く_(folder, projectName, hearingId) {
  const id = String(hearingId || '');
  if (!/[-\w]{25,}/.test(id)) throw new Error('ヒアリングIDが正規の管理行から確認できません');
  const sourceModifiedAt = DriveApp.getFileById(id).getLastUpdated().toISOString();
  const text = 相談室_原本テキスト_(SpreadsheetApp.openById(id));
  const revision = 相談室_ハッシュ_(text);
  const data = { schema: 'consultation-hearing/v1', source: 'hearing-sheet', caseName: caseCompany_(projectName), hearingId: id, revision: revision, sourceModifiedAt: sourceModifiedAt, text: text };
  const serialized = JSON.stringify(data, null, 2);
  const it = filesByName_(folder, 相談室_原本ファイル名_);
  if (it.hasNext()) {
    const file = it.next();
    if (it.hasNext()) throw new Error('相談室のヒアリング原本スナップショットが複数あります');
    try { const old = JSON.parse(file.getBlob().getDataAsString('UTF-8')); if (old && old.revision === revision && old.sourceModifiedAt === sourceModifiedAt && old.hearingId === id && old.caseName === data.caseName) return data; } catch (e) { /* replace the damaged internal snapshot */ }
    file.setContent(serialized);
  } else folder.createFile(相談室_原本ファイル名_, serialized, 'application/json');
  return data;
}

/**
 * 承認ボタン直前のための固定読取。任意のヒアリングID・フォルダ・関数名は
 * 受け取らず、案件名から既存の管理シートを一意に引いて原本を再解決する。
 * 原本自体やスナップショットは変更しない。
 */
function 相談室_管理行から原本を解決_(cfg, caseName) {
  const wanted = caseCompany_(caseName);
  if (!wanted || wanted.length > 120) throw new Error('案件名が不正です');
  const rows = mgmtSheet_(cfg).getDataRange().getValues();
  const hits = [];
  for (let i = 1; i < rows.length; i++) {
    const hearingId = String(rows[i][0] || '').trim();
    const projectName = String(rows[i][1] || '').trim();
    if (caseCompany_(projectName) === wanted && /[-\w]{25,}/.test(hearingId)) hits.push({ hearingId: hearingId, projectName: projectName });
  }
  if (hits.length !== 1) throw new Error(hits.length ? '案件に対応するヒアリング原本が一意ではありません' : '案件に対応するヒアリング原本が管理シートにありません');
  return hits[0];
}

/** Apps Script API から呼ぶ公開入口。既存のDrive/Spreadsheet読取だけを使う。 */
function スライド_相談室承認前ヒアリングを読む(caseName) {
  const cfg = getConfig_();
  const resolved = 相談室_管理行から原本を解決_(cfg, caseName);
  const source = DriveApp.getFileById(resolved.hearingId);
  const text = 相談室_原本テキスト_(SpreadsheetApp.openById(resolved.hearingId));
  return {
    ok: true,
    caseName: caseCompany_(resolved.projectName),
    hearingId: resolved.hearingId,
    revision: 相談室_ハッシュ_(text),
    sourceModifiedAt: source.getLastUpdated().toISOString()
  };
}
