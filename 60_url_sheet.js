// ============================================================
// 【URL管理】シート
//
// 案件フォルダができたときに、4章（集客〜開催）で使うURLと日程を
// 書き込むための空のシートを1枚だけ入れておく。
//
// ねらい: 運営担当が「発行したURLをどこに保存するか」で迷わないようにする。
//         マニュアル4-1の「案件フォルダに保存しておく」の受け皿。
//
// 行を足す・減らす・文言を変えるときは URL_SHEET_ROWS だけを直す。
// 「■」で始まる行は見出し行として自動で色が付く。
// ============================================================

const URL_SHEET_PREFIX = '【URL管理】';

const URL_SHEET_ROWS = [
  ['■ 日程', '', ''],
  ['登壇日（1回目・ライブ）', '', '4-1で開催回を登録する／4-3の開催案内'],
  ['配信日（2回目・擬似ライブ）', '', '4-1で開催回を登録する／4-3'],
  ['配信日（3回目・擬似ライブ）', '', '4-1で開催回を登録する／4-3'],
  ['納品予定日', '', 'この前日までに集客ページURLを制作担当へ渡す'],
  ['集客開始日', '', '4-2で集客ページを公開したあと'],

  ['■ 4-1 配信設定', '', ''],
  ['ライブ配信のパネリストURL', '', '4-3の開催案内で登壇者へ送る入室URL'],
  ['擬似ライブURL（2回目）', '', 'いまは使わない（参加URLは申込ごとに自動で発行される）'],
  ['擬似ライブURL（3回目）', '', 'いまは使わない（参加URLは申込ごとに自動で発行される）'],
  ['アンケート回答者用URL', '', '4-1でアンケートを作る／4-3で参加者へ送付'],
  ['アンケート結果スプレッドシートURL', '', '4-1でフォームと連携して作る／4-3の御礼メッセージでお客様へ送る'],

  ['■ 4-2 集客ページ', '', ''],
  ['ウェビナータイトル', '', '★自社の集客ページ・メディアレーダー・Peatixで同じ文字にする。申込の自動取り込みの照合に使う'],
  // ★行名を変えるときは tools/定型連絡.gs の 連絡_差し込みを作る_ と
  //   moushikomi-gas の writeUrlSheet_ の呼び出し名を必ず一緒に見ること。
  //   ここだけ変えると、3-3の納品連絡が {集客ページURL} を埋められず止まる。
  ['集客ページURL', '', '納品物。公開すると自動で入る。納品予定日の前日までに制作担当へ渡す'],
  ['集客者一覧リストURL', '', '納品物。お客様へは閲覧者権限で共有。申込の自動取り込み先にもなる'],
  ['メディアレーダー掲載URL', '', '社内確認用。お客様への共有は不要'],
  ['Peatix掲載URL', '', '社内確認用。お客様への共有は不要'],

  ['■ 4-3 開催', '', ''],
  ['録画データURL', '', '4-3で共有／4-4の最終納品'],
  ['録画データのパスコード', '', '4-3で共有／4-4の最終納品'],
  ['録画データの視聴期限', '', '4-4の最終納品メールに書く']
];

/**
 * 【出力】生成物にある案件フォルダのうち、【URL管理】がまだ無いものに作る。
 *
 * この仕組みを入れる前に始まった案件は【URL管理】を持っていないため、
 * 一度だけ実行して追いつかせる。すでにあるものは触らない。
 */
function backfillUrlSheets() {
  const cfg = getConfig_();
  const it = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID).getFolders();
  let made = 0, had = 0;
  while (it.hasNext()) {
    const folder = it.next();
    if (filesByName_(folder, URL_SHEET_PREFIX + folder.getName()).hasNext()) {
      Logger.log('・' + folder.getName() + '：すでにあります');
      had++;
      continue;
    }
    try {
      createUrlSheet_(cfg, folder.getName());
      Logger.log('・' + folder.getName() + '：作成しました');
      made++;
    } catch (e) {
      Logger.log('⚠ ' + folder.getName() + '：作成に失敗 ' + e);
    }
  }
  Logger.log('新しく作成 ' + made + '件 / すでにあった ' + had + '件');
}

/**
 * 【テスト用】URL管理シートを1枚だけ試しに作る。本番の案件には触らない。
 *
 * 実行すると「【出力】生成物」の中に「__テスト」フォルダができ、
 * その中に「【URL管理】__テスト」が入る。中身を確認したらフォルダごと消してよい。
 */
function testUrlSheet() {
  const file = createUrlSheet_(getConfig_(), '__テスト_ヒアリングシート');
  Logger.log('作成しました: ' + file.getUrl());
}

/**
 * 案件フォルダに【URL管理】シートを1枚作る。
 * すでに同じ名前のファイルがあれば何もしない（入力済みの内容を消さないため）。
 */
function createUrlSheet_(cfg, projectName) {
  const folder = caseFolder_(cfg, projectName);
  const name = URL_SHEET_PREFIX + caseCompany_(projectName);

  const exist = filesByName_(folder, name);
  if (exist.hasNext()) return exist.next();

  const ss = SpreadsheetApp.create(name);
  const sh = ss.getSheets()[0];
  sh.setName('URL管理');

  const rows = [['項目', '入力欄', 'いつ使うか']].concat(URL_SHEET_ROWS);
  sh.getRange(1, 1, rows.length, 3).setValues(rows);

  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#d9d9d9');
  for (let i = 0; i < URL_SHEET_ROWS.length; i++) {
    if (String(URL_SHEET_ROWS[i][0]).indexOf('■') === 0) {
      sh.getRange(i + 2, 1, 1, 3).setFontWeight('bold').setBackground('#f0f0f0');
    }
  }

  // 入力欄（B列）だけ色を変えて、どこに書くかを一目で分かるようにする
  for (let i = 0; i < URL_SHEET_ROWS.length; i++) {
    if (String(URL_SHEET_ROWS[i][0]).indexOf('■') !== 0) {
      sh.getRange(i + 2, 2).setBackground('#fff8e1');
    }
  }

  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 420);
  sh.setColumnWidth(3, 320);
  sh.setFrozenRows(1);
  SpreadsheetApp.flush();

  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(folder);
  return file;
}
