// ============================================================
// 納品ボタン
//
// 構成シートの設定タブでチェックを入れると、顧客に渡す4点を
// 【納品】→ 会社名 フォルダへコピーする。
//
// ★フォルダは作らない。宮津さんがヒアリングシートを作るときに
//   自分で用意するもの。無ければエラーを返して止まる。
//   （以前はGASが自動生成していたが、二重になるのでやめた）
//
// ★移動ではなくコピー。原本を動かすと、社内で直したものが
//   そのまま顧客に見えてしまう。直したものを届けたいときは、
//   もう一度チェックを入れる。
// ============================================================

// 共有ドライブの【納品】フォルダ。名前で探すと、共有ドライブの親を
// たどれないアカウントでマイドライブに別物を作る事故が起きるのでIDで開く。
const DELIVERY_ROOT_ID = '1l3qvbW2u-4RLBvNov4ZUj-opEGapdZnJ';

/**
 * 納品のセルを用意する（設定タブ A21:A23 / B22:B23）。
 * A21=見出し / A22=実行チェック（B22） / A23=結果（B23）
 */
function ensureDeliverCells_(st) {
  try {
    if (String(st.getRange('A21').getValue() || '').trim() !== '') return;
    if (String(st.getRange('A22').getValue() || '').trim() !== '') return;
    st.getRange('A21:A23').setValues([
      ['― 検品が終わったら納品する ―'],
      ['4点を【納品】へコピー（チェックすると実行）'],
      ['結果（システムが書きます）']
    ]);
    st.getRange('A21').setFontWeight('bold').setBackground('#efefef');
    st.getRange('B22').setNote(
      'スライド・音声mp3・サムネイル・LPの4点を、【納品】→ 会社名フォルダへコピーします。\n' +
      'フォルダは宮津さんが事前に社名で作っているものを使います。無いときはエラーになります。\n' +
      '同じ名前のファイルが既にあれば、古いほうはゴミ箱へ移します（30日以内なら復元可）。\n' +
      '直して作り直したあとは、もう一度チェックを入れてください。'
    );
    st.getRange('B23').setWrap(true);
    if (st.getRange('B22').getDataValidation() === null) st.getRange('B22').insertCheckboxes();
    st.setRowHeight(23, 60);
  } catch (e) { /* 既存シートの体裁は壊さない */ }
}

/** 【納品】の中の会社名フォルダを返す。無ければ null（作らない）。 */
function deliveryFolderOf_(company) {
  let root;
  try { root = DriveApp.getFolderById(DELIVERY_ROOT_ID); } catch (e) { return null; }
  const it = root.getFoldersByName(company);
  return it.hasNext() ? it.next() : null;
}

/** コピー先の同名ファイルをゴミ箱へ移してから、1件コピーする。 */
function deliverOne_(file, dest, newName) {
  const old = dest.getFilesByName(newName);
  while (old.hasNext()) { try { old.next().setTrashed(true); } catch (e) { /* 個別の失敗は無視 */ } }
  file.makeCopy(newName, dest);
}

/**
 * チェックが入った案件の納品物4点をコピーする。
 * スライド / 音声mp3 / サムネイル / LP。台本txtは入れない
 * （中身はスライドの話者ノートに入っており、顧客には要らないため）。
 */
function processDelivery_() {
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

    ensureDeliverCells_(st);

    if (st.getRange('B22').getValue() !== true) continue;
    st.getRange('B22').setValue(false);   // 先に外す。失敗しても同じ依頼を回し続けない

    const company = caseCompany_(String(projectName));
    const dest = deliveryFolderOf_(company);
    if (!dest) {
      st.getRange('B23').setValue(
        '【納品】の中に「' + company + '」フォルダが見つかりません。' +
        'システムは納品フォルダを作りません。先に作ってから、もう一度チェックを入れてください。');
      continue;
    }

    let src;
    try { src = caseFolderIfExists_(cfg, String(projectName)); if (!src) throw new Error('案件フォルダがありません'); } catch (e) {
      st.getRange('B23').setValue('案件フォルダを開けません: ' + String(e));
      continue;
    }

    const done = [];
    const miss = [];

    // ① スライド / ② 音声 / ③ LP … 案件フォルダから
    [['【スライド】' + String(projectName), 'スライド'],
     ['【音声】' + company + '.mp3', '音声'],
     ['【LP】' + company, 'LP']].forEach(function (pair) {
      const it = src.getFilesByName(pair[0]);
      if (!it.hasNext()) { miss.push(pair[1]); return; }
      try { deliverOne_(it.next(), dest, pair[0]); done.push(pair[1]); }
      catch (e) { miss.push(pair[1] + '(' + e + ')'); }
    });

    // ④ サムネイル … 【素材】投影資料 の会社フォルダから。名前を付け替えて入れる
    try {
      const aroot = DriveApp.getFolderById(cfg.ASSETS_FOLDER_ID);
      const asub = aroot.getFoldersByName(String(projectName));
      if (!asub.hasNext()) { miss.push('サムネイル(素材フォルダなし)'); }
      else {
        const tf = asub.next().getFilesByName('サムネイル.png');
        if (!tf.hasNext()) miss.push('サムネイル');
        else { deliverOne_(tf.next(), dest, '【サムネイル】' + company + '.png'); done.push('サムネイル'); }
      }
    } catch (e) { miss.push('サムネイル(' + e + ')'); }

    st.getRange('B23').setValue(
      (done.length ? '納品しました: ' + done.join('・') : '納品できたものはありません') +
      (miss.length ? ' ／ 見つからなかったもの: ' + miss.join('・') : '') +
      ' → 【納品】/' + company + '/');
  }
}
