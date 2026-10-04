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
//
// ★手で直した納品版は上書きしない（2026-09-25・宮津さん「手で直した分を上書きしない」）。
//   それまでは押し直すたびに、納品フォルダの同じ名前のファイルをゴミ箱へ移してコピーし直していた。
//   9/25、制作担当がホスピタソン社の納品版スライドを手で直した（p46・p64・p87-90の図解差し替え、
//   p104のQ&A追加）。これは意図的な手直しで、押し直しで消えてはいけない。
//   そこで、この処理がコピーした時刻を覚えておき、納品版がそれより後に更新されていたら
//   （＝人かお客様が直した）作り直さずに止める。→ deliveredKept_
//   ※Q&Aは2026-09-25から自動で入る（No.104.5・構成シートにQ&Aの行がある新しい案件だけ）。
//     ホスピタソン社のように手で足した案件は構成シートにQ&Aの行が無いので、組み直しても二重にならない。
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
    st.getRange('B22').setNote(DELIVER_NOTE);
    st.getRange('B23').setWrap(true);
    if (st.getRange('B22').getDataValidation() === null) st.getRange('B22').insertCheckboxes();
    st.setRowHeight(23, 60);
  } catch (e) { /* 既存シートの体裁は壊さない */ }
}

/** B22のメモ。★動きを変えたらここを直す（押したときに古いメモを書き替える。→ processDelivery_） */
const DELIVER_NOTE =
  'スライド・音声mp3・サムネイル・LPの4点を、【納品】→ 会社名フォルダへコピーします。\n' +
  'フォルダは案件開始のときに自動で作られています。無いときはエラーになります。\n' +
  '前回この処理が入れたままの納品版は、ゴミ箱へ移して入れ直します（30日以内なら復元可）。\n' +
  '★納品版があとから手で直されていたら（お客様の編集も含む）、上書きせずに止めます。\n' +
  '作り直すなら、手直しを先にセルへ移し、納品版を手でゴミ箱へ移してから、もう一度チェックを入れてください。';

/** 【納品】の中の会社名フォルダを返す。無ければ null（作らない）。 */
function deliveryFolderOf_(company) {
  let root;
  try { root = DriveApp.getFolderById(DELIVERY_ROOT_ID); } catch (e) { return null; }
  const it = foldersByName_(root, company);
  return it.hasNext() ? it.next() : null;
}

/* ---------- 手直しした納品版を守る（2026-09-25） ----------
 *
 * ■ 前回コピーした時刻の置き場所
 *   スクリプトプロパティに「DELIVERED_<納品フォルダのID>_<種類>」＝{id, at} で持つ。
 *   ・Driveのファイルの説明欄には書かない（お客様に見える。お客様の編集で消えることもある）
 *   ・appProperties は Drive の高度なサービスが要り、このプロジェクトでは有効にしていない
 *   ・ファイルIDも持つので、納品版の名前が変えられていても（ウェビナータイトル名に変えるなど）見つけられる
 *
 * ■ 手直しありとみなすもの
 *   ・記録があり、そのファイルの最終更新がコピーした時刻より後（DELIVERY_EDIT_MARGIN_MS の余裕つき）
 *   ・★記録が無いGoogle形式のファイル（スライド・LPのドキュメント）。この仕組みより前に入った
 *     納品版（9/25に手で直したホスピタソン社のスライドなど）は、直されたかどうか分からないので**守る側に倒す**
 *   ・記録が無い mp3・png は守らない（今までどおり入れ直す）。Drive上では中身を直せず、
 *     手で差し替えるなら「新しい版をアップロード」になるが、その場合も次回からは記録で見分けられる
 */
const DELIVERED_PROP_PREFIX = 'DELIVERED_';
/** コピー直後にDrive側で最終更新が少し動くことがあるので、その分は手直しとみなさない */
const DELIVERY_EDIT_MARGIN_MS = 2 * 60 * 1000;
/** 画面（B23）とログに出す文。★文面は宮津さんの指示どおり（2026-09-25） */
const DELIVERY_KEEP_MESSAGE =
  '納品版に手直しがあるため作り直しません。作り直すなら手直しを先にセルへ移し、納品版を手でゴミ箱へ移してから押してください';

function deliveredKey_(dest, kind) {
  return DELIVERED_PROP_PREFIX + dest.getId() + '_' + kind;
}

function deliveredRecord_(dest, kind) {
  try {
    const v = PropertiesService.getScriptProperties().getProperty(deliveredKey_(dest, kind));
    return v ? JSON.parse(v) : null;
  } catch (e) { return null; }
}

function saveDeliveredRecord_(dest, kind, copy) {
  PropertiesService.getScriptProperties().setProperty(deliveredKey_(dest, kind),
    JSON.stringify({ id: copy.getId(), at: copy.getLastUpdated().getTime() }));
}

/** 手直しがあるか。★迷ったら「ある」（上書きしない側）に倒す */
function handEdited_(f, rec) {
  if (rec && rec.id === f.getId()) return f.getLastUpdated().getTime() > rec.at + DELIVERY_EDIT_MARGIN_MS;
  return String(f.getMimeType()).indexOf('application/vnd.google-apps.') === 0;
}

/**
 * いま納品フォルダにある、その種類の納品版（ゴミ箱の中のものは除く）。
 * 同じ名前のものに加えて、前回コピーしたファイル（名前が変えられていても）も拾う。
 */
function deliveredNow_(dest, names, rec) {
  const out = [];
  const seen = {};
  // ★名前は1つとは限らない。スライドは2026-09-25にウェビナータイトル名へ変えたので、
  //   それより前の「【スライド】◯◯」の名前でも探す（古い版が横に残って2枚にならないように）
  [].concat(names).forEach(function (name) {
    const it = filesByName_(dest, name);
    while (it.hasNext()) {
      const f = it.next();
      if (f.isTrashed() || seen[f.getId()]) continue;
      seen[f.getId()] = true;
      out.push(f);
    }
  });
  if (rec && rec.id && !seen[rec.id]) {
    try {
      const f = DriveApp.getFileById(rec.id);
      if (!f.isTrashed() && f.getParents().hasNext() && f.getParents().next().getId() === dest.getId()) out.push(f);
    } catch (e) { /* もう無い（人が消した）。拾わない */ }
  }
  return out;
}

/**
 * 1件を納品フォルダへ入れる。前回この処理が入れたままのものはゴミ箱へ移して入れ直す。
 * ★手直しがあれば、何も消さず・何も入れずに { kept: [名前…] } を返す。
 */
function deliverOne_(file, dest, newName, kind, oldNames) {
  const rec = deliveredRecord_(dest, kind);
  const now = deliveredNow_(dest, [newName].concat(oldNames || []), rec);
  const edited = now.filter(function (f) { return handEdited_(f, rec); });
  if (edited.length) return { kept: edited.map(function (f) { return f.getName(); }) };

  now.forEach(function (f) { try { f.setTrashed(true); } catch (e) { /* 個別の失敗は無視 */ } });
  const copy = file.makeCopy(newName, dest);
  saveDeliveredRecord_(dest, kind, copy);
  return { copied: true };
}

/**
 * サムネイル画像は【素材】投影資料にしかできず、これまで【出力】生成物には
 * 3-2f（この下の processDelivery_、B22チェック）を押すまで一切入らなかった。
 * だが 32a_LPとタイトル.md 自身が「3-2a〜3-2fで開くものは、ぜんぶ【出力】生成物
 * の中にある」と説明しており、実際にはサムネイルだけそこに無い状態だった。
 * 検品（B22）を待たずに済むよう、サムネイルができ次第ここへも複製しておく。
 * ★【納品】へは入れない。あちらは検品後にお客様へ渡す場所なので、
 *   processDelivery_ の役割（B22で明示的に押す）とは変えない。
 * （2026-09-24・ホスピタソン社で、運営担当が【納品】【出力】どちらにも
 *   見つからず「【素材】投影資料の中にあります」と個別に案内する事態になった）
 */
function syncThumbnailToOutput_() {
  const cfg = getConfig_();
  let aroot;
  try { aroot = DriveApp.getFolderById(cfg.ASSETS_FOLDER_ID); } catch (e) { return; }
  const rows = mgmtSheet_(cfg).getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const projectName = rows[i][1];
    if (!projectName) continue;

    let thumb;
    try {
      const asub = foldersByName_(aroot, String(projectName));
      if (!asub.hasNext()) continue;
      const tf = filesByName_(asub.next(), 'サムネイル.png');
      if (!tf.hasNext()) continue;
      thumb = tf.next();
    } catch (e) { continue; }

    const dest = caseFolderIfExists_(cfg, String(projectName));
    if (!dest) continue;

    const newName = '【サムネイル】' + caseCompany_(String(projectName)) + '.png';
    try {
      const existing = filesByName_(dest, newName);
      if (existing.hasNext()) {
        const old = existing.next();
        // ★素材のほうが新しければ（作り直された）差し替える。同じなら何もしない
        if (old.getLastUpdated().getTime() >= thumb.getLastUpdated().getTime()) continue;
        old.setTrashed(true);
      }
      thumb.makeCopy(newName, dest);
    } catch (e) { /* 1件の失敗で他の案件を止めない */ }
  }
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
    // ★前からある構成シートのメモは古い説明（押せば必ず入れ直す）のままなので、押されたときに書き替える
    try { if (st.getRange('B22').getNote() !== DELIVER_NOTE) st.getRange('B22').setNote(DELIVER_NOTE); } catch (e) { /* 体裁の失敗で止めない */ }

    // ★新しい流れ（管理シート H列＝新）は、承認（B2＝「確認して承認」）のあとだけ受け付ける。旧の流れは今までどおり
    const 新記 = 新の流れか_(rows[i]) ? 進み_記録を読む_(String(rows[i][0])) : null;
    let 現在revision = '';
    if (新の流れか_(rows[i])) {
      try {
        const 出力フォルダ = caseFolderIfExists_(cfg, String(projectName));
        if (出力フォルダ) 現在revision = 成果物revision_(観測する_(cfg, rows[i], 出力フォルダ));
      } catch (e) { 現在revision = ''; }
    }
    if (新の流れか_(rows[i]) && (!現在revision || !新の流れの承認済みか_(String(rows[i][0]), st, 新記, 現在revision))) {
      st.getRange('B23').setValue('まだ承認されていないので、納品フォルダへは入れていません。全部できたら「確認して承認」（操作パネルかB2）をしてから、もう一度押してください');
      continue;
    }

    const r = 納品する_(cfg, String(projectName), st);
    if (新の流れか_(rows[i]) && r.入れた) 進み_納品を記録_(String(rows[i][0]), r.文, 現在revision);
  }
}

/**
 * 1件ぶんの納品（B22 の中身）。結果は B23 に書き、{文, 入れた} を返す。
 * ★processDelivery_（B22）と操作パネルの「納品フォルダへ入れる」の両方から呼ぶ。中身は B22 のときと同じ。
 */
function 納品する_(cfg, projectName, st) {
    const company = caseCompany_(String(projectName));
    const dest = deliveryFolderOf_(company);
    if (!dest) {
      const 文 = '【納品】の中に「' + company + '」フォルダが見つかりません。' +
        'システムは納品フォルダを作りません。先に作ってから、もう一度チェックを入れてください。';
      st.getRange('B23').setValue(文);
      return { 文: 文, 入れた: false };
    }

    let src;
    try { src = caseFolderIfExists_(cfg, String(projectName)); if (!src) throw new Error('「【出力】生成物」の中に、この案件のフォルダがありません（フォルダ名が「◯◯社」で合っているか見てください）'); } catch (e) {
      const 文 = 失敗の文_('納品', e, 'B22');
      st.getRange('B23').setValue(文);
      return { 文: 文, 入れた: false };
    }

    const done = [];
    const miss = [];
    const kept = [];   // 手直しがあるので作り直さなかったもの（→ deliverOne_）

    const put = function (f, newName, kind, oldNames) {
      const r = deliverOne_(f, dest, newName, kind, oldNames);
      if (r.kept) {
        kept.push(kind);
        Logger.log('★' + projectName + '｜' + kind + '：' + DELIVERY_KEEP_MESSAGE + '（' + r.kept.join('・') + '）');
      } else {
        done.push(kind);
      }
    };

    // ① スライド / ② 音声 / ③ LP … 案件フォルダから
    // ★納品フォルダのスライドは、ウェビナータイトル名で入れる（2026-09-25・宮津さん。1-2・3-3の表と同じ）。
    //   タイトルが読めないときだけ、これまでの「【スライド】◯◯」の名前で入れる
    const 旧スライド名 = '【スライド】' + String(projectName);
    const タイトル = webinarTitleOf_(src, company);
    [[旧スライド名, 'スライド', タイトル || 旧スライド名, タイトル ? [旧スライド名] : []],
     ['【音声】' + company + '.mp3', '音声'],
     ['【LP】' + company, 'LP']].forEach(function (pair) {
      const it = filesByName_(src, pair[0]);
      if (!it.hasNext()) { miss.push(pair[1]); return; }
      try { put(it.next(), pair[2] || pair[0], pair[1], pair[3]); }
      catch (e) { miss.push(pair[1] + '(' + e + ')'); }
    });

    // ④ サムネイル … 【素材】投影資料 の会社フォルダから。名前を付け替えて入れる
    try {
      const aroot = DriveApp.getFolderById(cfg.ASSETS_FOLDER_ID);
      const asub = foldersByName_(aroot, String(projectName));
      if (!asub.hasNext()) { miss.push('サムネイル(素材フォルダなし)'); }
      else {
        const tf = filesByName_(asub.next(), 'サムネイル.png');
        if (!tf.hasNext()) miss.push('サムネイル');
        else put(tf.next(), '【サムネイル】' + company + '.png', 'サムネイル');
      }
    } catch (e) { miss.push('サムネイル(' + e + ')'); }

    const 文 =
      (kept.length ? '★' + kept.join('・') + '：' + DELIVERY_KEEP_MESSAGE + '。 ／ ' : '') +
      (done.length ? '納品しました: ' + done.join('・') : '入れ直したものはありません') +
      (miss.length ? ' ／ 見つからなかったもの: ' + miss.join('・') : '') +
      ' → 【納品】/' + company + '/';
    st.getRange('B23').setValue(文);
    return { 文: 文, 入れた: done.length > 0 || kept.length > 0, 残した: kept };
}

/**
 * ウェビナータイトル（【LP】◯◯ の1行目＝サブタイトルと2行目＝メインタイトルをつなげたもの）。
 * ★【LP】の書き方は 30_prompts.js の「1行目がサブタイトル、2行目がメインタイトル」に合わせている。
 * 読めなければ ''（呼ぶ側が【スライド】◯◯ の名前に戻す）。
 */
function webinarTitleOf_(src, company) {
  try {
    const it = filesByName_(src, '【LP】' + company);
    if (!it.hasNext()) return '';
    const f = it.next();
    if (f.getMimeType() !== MimeType.GOOGLE_DOCS) return '';
    const lines = DocumentApp.openById(f.getId()).getBody().getText()
      .split('\n').map(function (s) { return s.trim(); }).filter(function (s) { return s; });
    return lines.slice(0, 2).join('').replace(/[\\/]/g, '／').slice(0, 120);
  } catch (e) {
    Logger.log('ウェビナータイトルを読めません（【スライド】の名前で入れます）: ' + e);
    return '';
  }
}
