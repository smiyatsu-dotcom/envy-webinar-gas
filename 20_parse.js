/**
 * 20_parse.gs
 * AI出力（貼り付けられたJSON）のパース・検証ユーティリティ
 * ※ APIレス版: Claude API呼び出しは行わない。ユーザーがチャット/Claude Codeで生成した
 *   JSONを構成シート「AI出力貼り付け」タブに貼り、GASが取り込む。
 */

/**
 * テキストからJSON配列を抽出してパースする。
 * - コードフェンス・前置きを除去
 * - 分割出力を連結した「][」は「,」に置換して1配列にマージ
 */
function parseJsonArray_(text) {
  let t = String(text).replace(/```json|```/g, '').trim();
  const s = t.indexOf('[');
  const e = t.lastIndexOf(']');
  if (s === -1 || e === -1 || e <= s) throw new Error('JSON配列が見つかりません。貼り付け内容を確認してください: ' + t.slice(0, 120));
  t = t.slice(s, e + 1).replace(/\]\s*,?\s*\[/g, ',');
  return JSON.parse(t);
}

/**
 * 出力行の検証・正規化
 * @param {Array<Object>} rows  [{no, ...}]
 * @param {Array<number>} expectedNos 期待する行No一覧
 * @param {string} field 'daihon' | 'kousei_text'
 * @param {Array<string>} warnings 警告の書き込み先
 * @return {Object} no -> テキスト のマップ
 */
function validateRows_(rows, expectedNos, field, warnings) {
  const map = {};
  rows.forEach(function (r) {
    const no = Number(r.no);
    if (!isNaN(no) && r[field] !== undefined) map[no] = String(r[field]);
  });
  const missing = expectedNos.filter(function (no) { return map[no] === undefined; });
  // No.99（特典⒉）と No.104（分割払い）は空を許容
  const hardMissing = missing.filter(function (no) { return no !== 99 && no !== 104; });
  if (hardMissing.length > 0) {
    throw new Error(field + ' の欠落行があります: No.' + hardMissing.join(',') + '（チャットで「欠落行だけ再出力して」と依頼し、追記貼り付けしてください）');
  }
  missing.forEach(function (no) {
    map[no] = '';
    warnings.push(field + ' No.' + no + ' は空で補完（特典/分割なしの想定）');
  });
  // 図解タグの必須チェック（構成テキストのみ）: 欠落は自動補完
  if (field === 'kousei_text') {
    const tags = requiredImgTags_();
    expectedNos.forEach(function (no) {
      const tag = tags[no];
      if (tag && map[no] && map[no].indexOf('図解（') === -1) {
        map[no] = map[no] + '\n\n' + tag;
        warnings.push('No.' + no + ' に図解タグを自動補完: ' + tag);
      }
    });
  }
  return map;
}

/**
 * 文字数規定（kouseiSpec_）との照合。超過は警告のみ（停止しない）。
 * 図解タグ行・空行は対象外。
 */
function checkKouseiLimits_(kouseiMap, warnings) {
  const spec = kouseiSpec_();
  Object.keys(kouseiMap).forEach(function (k) {
    const no = Number(k);
    const s = spec[no];
    if (!s) return;
    const lines = String(kouseiMap[no]).split('\n')
      .map(function (l) { return l.replace(/図解（[^）]*）/g, '').trim(); })
      .filter(function (l) { return l !== ''; });
    if (lines.length > s.maxLines) {
      warnings.push('No.' + no + ' 行数超過: ' + lines.length + '行（規定' + s.maxLines + '行）');
    }
    lines.forEach(function (l, i) {
      if (l.length > s.maxLine) {
        warnings.push('No.' + no + ' ' + (i + 1) + '行目が' + l.length + '字（規定' + s.maxLine + '字以内）: 「' + l.slice(0, 15) + '…」');
      }
    });
  });
}