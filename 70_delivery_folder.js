// ============================================================
// 【納品】フォルダ
//
// お客様に共有する納品物だけを入れる、案件ごとのフォルダ。
//
// なぜ分けるか:
//   【出力】生成物 の案件フォルダには、プロンプトDocや中間JSONなど
//   社内用のファイルが入っている。フォルダごと共有すると、それも
//   すべてお客様に見えてしまう。納品用は最初から別の場所に置く。
//
// 中身は「コピー」を入れる。原本を移動しない。
//   → 社内で原本を直しても、お客様に見えている物は勝手に変わらない。
//     直したものを届けたいときは、入れ直す。
// ============================================================

const DELIVERY_ROOT_NAME = '【納品】';

// 【納品】フォルダのID。共有ドライブの中で【出力】生成物と並んでいる。
//
// 名前や場所から探さずIDで直接開く。以前は【出力】生成物の親をたどっていたが、
// 共有ドライブでは親をたどれないアカウントがあり、そのときマイドライブに
// 別の【納品】を作ってしまう危険があった。
const DELIVERY_ROOT_ID = '1l3qvbW2u-4RLBvNov4ZUj-opEGapdZnJ';

/**
 * 【納品】ルートフォルダを返す。
 */
function deliveryRoot_(cfg) {
  try {
    return DriveApp.getFolderById(DELIVERY_ROOT_ID);
  } catch (e) {
    throw new Error('【納品】フォルダを開けません。ID（DELIVERY_ROOT_ID）と共有設定を確認してください: ' + e);
  }
}

/**
 * 【納品】の中に会社名フォルダを作って返す。すでにあればそれを返す。
 * 案件フォルダと同じ会社名の付け方（caseCompany_）を使うので、名前は必ずそろう。
 */
function deliveryFolder_(cfg, projectName) {
  const root = deliveryRoot_(cfg);
  const nm = caseCompany_(projectName);
  const it = root.getFoldersByName(nm);
  return it.hasNext() ? it.next() : root.createFolder(nm);
}

/**
 * 【テスト用】納品フォルダを試しに作る。本番の案件には触らない。
 * 「【納品】__テスト」ができる。確認したら消してよい。
 */
function testDeliveryFolder() {
  const f = deliveryFolder_(getConfig_(), '__テスト_ヒアリングシート');
  Logger.log('作成しました: ' + f.getUrl());
}
