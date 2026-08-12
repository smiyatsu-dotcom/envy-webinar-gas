/**
 * 10_main.gs（APIレス版）
 * エントリポイント・状態管理（管理シート）・トリガー
 *
 * 管理シート列:
 *  A:ヒアリングファイルID B:案件名 C:ステータス D:構成シートURL E:スライドURL F:メッセージ G:更新日時
 *
 * 状態遷移:
 *  （ヒアリング検知）→ AI出力待ち →（貼り付け＋取り込みチェック）→ レビュー待ち
 *  →（承認チェック）→ スライド生成中 → 完了 ／ 各段階で失敗すると エラー
 */

const MGMT_HEADER = ['ヒアリングファイルID', '案件名', 'ステータス', '構成シートURL', 'スライドURL', 'メッセージ', '更新日時'];

/**
 * 初回セットアップ: 管理シートのヘッダ整備＋5分毎トリガー作成
 */
function setup() {
  const cfg = getConfig_();
  const sh = mgmtSheet_(cfg);
  if (String(sh.getRange(1, 1).getValue()) !== MGMT_HEADER[0]) {
    sh.getRange(1, 1, 1, MGMT_HEADER.length).setValues([MGMT_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'main') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('main').timeBased().everyMinutes(5).create();
  Logger.log('セットアップ完了: 5分毎トリガーを作成しました');
}

/** トリガー本体（5分毎）。多重起動はロックで防止 */
function main() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    processNewHearings_();
    processImports_();
    processApprovals_();
    processSlideRegen_();
    convertLpTexts_();
    processTitleConfirm_();
    processLpFixRequests_();
    processSlideFixRequests_();
    processDelivery_();
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// フェーズ1: 新規ヒアリング → 構成シート雛形＋プロンプトDoc生成
// ============================================================
function processNewHearings_() {
  const cfg = getConfig_();
  const known = registry_(cfg);
  const files = DriveApp.getFolderById(cfg.INPUT_FOLDER_ID).getFiles();

  while (files.hasNext()) {
    const f = files.next();
    if (known[f.getId()]) continue;
    if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) {
      if (/spreadsheet|excel/i.test(f.getMimeType())) {
        upsertRow_(cfg, f.getId(), f.getName(), 'エラー', '', '', 'Google スプレッドシート形式ではありません（ドライブの自動変換をONにして再アップロードしてください）');
      }
      continue;
    }
    // ★案件名は必ず前後の空白を落とす。Driveで「コピーを作成」してリネームすると
    //   末尾に空白が残ることがあり、そのまま使うと【台本】◯◯ .txt のように
    //   目に見えない空白入りのファイル名が量産される（2026-08-11 実測）。
    const pname = String(f.getName()).trim();
    try {
      const doc = createPromptDoc_(pname, f.getId());       // 30_prompts.gs
      const ss = createKouseiScaffold_(pname, f.getId(), doc.getUrl()); // 40_kousei_sheet.gs
      // 4章で使うURLの入れ物。補助なので、失敗しても本体は止めない
      try { createUrlSheet_(cfg, pname); } catch (e) { Logger.log('URL管理シートの作成に失敗: ' + e); }  // 60_url_sheet.gs
      // ★【納品】フォルダは作らない。宮津さんがヒアリングシートを作るときに
      //   自分で用意するもので、GASが作ると二重になる（2026-08-12 指摘）。
      //   そもそも中身をコピーする処理はどこにも無く、空フォルダが増えるだけだった。
      upsertRow_(cfg, f.getId(), pname, 'AI出力待ち', ss.getUrl(), '',
        '自動生成中です。LPは20分・サムネイルは30分・資料は50分ほど（手順書と同じ目安）。'
        + '1時間半たっても資料が揃わないときは宮津さんへ。手動で回すときのプロンプトDoc: ' + doc.getUrl());
    } catch (e) {
      upsertRow_(cfg, f.getId(), f.getName(), 'エラー', '', '', String(e));
    }
  }
}

// ============================================================
// フェーズ2: 取り込みチェック検知 → 貼り付けJSONを各タブへ反映
// ============================================================
function processImports_() {
  const cfg = getConfig_();
  const sh = mgmtSheet_(cfg);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const [hearingId, name, status, kouseiUrl] = values[i];
    if ((status !== 'AI出力待ち' && status !== 'レビュー待ち') || !kouseiUrl) continue; // レビュー待ちでも再取り込み可
    let ss;
    try { ss = SpreadsheetApp.openById(idFromUrl_(String(kouseiUrl))); } catch (e) { continue; }
    // 手貼りのチェックが入っていないときは、案件フォルダのtxtを見て自動で取り込む
    if (!isImportRequested_(ss)) {
      const auto = readCaseJsonTexts_(cfg, String(name));
      if (!auto || auto.stamp === getAutoStamp_(ss.getId())) continue;
      setAutoStamp_(ss.getId(), auto.stamp);  // 失敗しても同じ内容で繰り返さない（txtを直せば再実行される）
      try {
        const w = importAiOutput_(ss, auto);
        setImportResult_(ss, 'txtから自動取り込み' + (w.length ? '（警告: ' + w.join(' / ') + '）' : '') +
          ' → 各タブをレビューし、設定タブの承認チェックをONにしてください');
        setKouseiStatus_(ss.getId(), 'レビュー待ち');
        upsertRow_(cfg, hearingId, name, 'レビュー待ち', kouseiUrl, '', w.join(' / '));
      } catch (e) {
        setImportResult_(ss, 'txtの自動取り込みに失敗: ' + String(e));
        upsertRow_(cfg, hearingId, name, 'AI出力待ち', kouseiUrl, '', 'txt自動取り込み失敗: ' + String(e));
      }
      continue;
    }

    try {
      const warnings = importAiOutput_(ss);
      const msg = '取り込み成功' + (warnings.length ? '（警告: ' + warnings.join(' / ') + '）' : '') +
        ' → 各タブをレビューし、設定タブの承認チェックをONにしてください';
      setImportResult_(ss, msg);
      setKouseiStatus_(ss.getId(), 'レビュー待ち');
      upsertRow_(cfg, hearingId, name, 'レビュー待ち', kouseiUrl, '', warnings.join(' / '));
    } catch (e) {
      setImportResult_(ss, '取り込み失敗: ' + String(e));
      upsertRow_(cfg, hearingId, name, 'AI出力待ち', kouseiUrl, '', '取り込み失敗: ' + String(e));
    }
  }
}

// ============================================================
// フェーズ3: 承認検知 → スライド生成
// ============================================================
function processApprovals_() {
  const cfg = getConfig_();
  const sh = mgmtSheet_(cfg);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const [hearingId, name, status, kouseiUrl] = values[i];
    if (status !== 'レビュー待ち' || !kouseiUrl) continue;
    const kouseiId = idFromUrl_(String(kouseiUrl));
    let approved = false;
    try { approved = isApproved_(kouseiId); } catch (e) { continue; }
    if (!approved) continue;

    upsertRow_(cfg, hearingId, name, 'スライド生成中', kouseiUrl, '', '');
    setKouseiStatus_(kouseiId, 'スライド生成中');
    try {
      const result = generateSlides_(kouseiId); // 50_slides.gs
      upsertRow_(cfg, hearingId, name, '完了', kouseiUrl, result.pres.getUrl(), result.warnings.join(' / '));
      setKouseiStatus_(kouseiId, 'スライド生成済み: ' + result.pres.getUrl());
    } catch (e) {
      upsertRow_(cfg, hearingId, name, 'エラー', kouseiUrl, '', String(e));
      setKouseiStatus_(kouseiId, 'エラー: ' + String(e));
    }
  }
}

// ============================================================
// 管理シートユーティリティ
// ============================================================
function mgmtSheet_(cfg) {
  return SpreadsheetApp.openById(cfg.MGMT_SHEET_ID).getSheets()[0];
}

function registry_(cfg) {
  const values = mgmtSheet_(cfg).getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) if (values[i][0]) map[String(values[i][0])] = i + 1;
  return map;
}

function upsertRow_(cfg, hearingId, name, status, kouseiUrl, slideUrl, message) {
  const sh = mgmtSheet_(cfg);
  const reg = registry_(cfg);
  const row = [hearingId, name, status, kouseiUrl, slideUrl, message, new Date()];
  if (reg[hearingId]) {
    sh.getRange(reg[hearingId], 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}

function idFromUrl_(url) {
  const m = url.match(/[-\w]{25,}/);
  if (!m) throw new Error('URLからIDを抽出できません: ' + url);
  return m[0];
}

// ============================================================
// 単体テスト（手動実行用）
// ============================================================

/** ヒアリングID直指定でプロンプトDoc＋構成シート雛形を生成 */
function testCreateScaffold() {
  const hearingId = 'ここにヒアリングスプシIDを入れる';
  const doc = createPromptDoc_('テスト案件', hearingId);
  const ss = createKouseiScaffold_('テスト案件', hearingId, doc.getUrl());
  Logger.log('プロンプトDoc: ' + doc.getUrl());
  Logger.log('構成シート: ' + ss.getUrl());
}

/** 構成シートID直指定で貼り付けJSONを取り込み */
function testImport() {
  const kouseiId = 'ここに構成シートIDを入れる';
  const ss = SpreadsheetApp.openById(kouseiId);
  const warnings = importAiOutput_(ss);
  Logger.log('取り込み完了。警告: ' + warnings.join('\n'));
}

/** 構成シートID直指定でスライド生成（承認チェック不要） */
function testGenerateSlides() {
  const kouseiId = 'ここに構成シートIDを入れる';
  const result = generateSlides_(kouseiId);
  Logger.log('スライド: ' + result.pres.getUrl());
  Logger.log('警告: ' + result.warnings.join('\n'));
}
// ============================================================
// フェーズ3b: 「スライドを作り直す」チェック → その場で再生成
//
// これまでは PM が3手順（管理シートC列を戻す → mp3を消す → 承認OFF→ON）を
// 踏む必要があり、面倒なうえに mp3 の消し忘れで古い音声が残る事故が起きていた。
// チェック1つで全部やる。AIは動かないので費用はかからない。
// ============================================================
function processSlideRegen_() {
  const cfg = getConfig_();
  const rows = mgmtSheet_(cfg).getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const hearingId = rows[i][0];
    const projectName = rows[i][1];
    const kouseiUrl = rows[i][3];
    if (!projectName || !kouseiUrl) continue;

    const kouseiId = idFromUrl_(String(kouseiUrl));
    let st;
    try { st = SpreadsheetApp.openById(kouseiId).getSheetByName(TABS.SETTINGS); } catch (e) { continue; }
    if (!st) continue;

    ensureSlideRegenCell_(st);

    if (st.getRange('B16').getValue() !== true) continue;
    st.getRange('B16').setValue(false);   // 先に外す。失敗しても回し続けない

    // 古い音声を消す。残っていると見張りが「もうある」と判断して作り直さない
    try {
      const folder = caseFolder_(cfg, String(projectName));
      const mp3 = folder.getFilesByName('【音声】' + caseCompany_(String(projectName)) + '.mp3');
      while (mp3.hasNext()) mp3.next().setTrashed(true);
    } catch (e) { /* 音声が無くても再生成は進める */ }

    upsertRow_(cfg, hearingId, projectName, 'スライド生成中', kouseiUrl, '', '作り直し中');
    setKouseiStatus_(kouseiId, 'スライド生成中（作り直し）');
    try {
      const result = generateSlides_(kouseiId);   // 50_slides.gs
      upsertRow_(cfg, hearingId, projectName, '完了', kouseiUrl, result.pres.getUrl(), result.warnings.join(' / '));
      setKouseiStatus_(kouseiId, 'スライド生成済み: ' + result.pres.getUrl());
    } catch (e) {
      upsertRow_(cfg, hearingId, projectName, 'エラー', kouseiUrl, '', String(e));
      setKouseiStatus_(kouseiId, 'エラー: ' + String(e));
    }
  }
}
