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
    convertLpTexts_();
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
    try {
      const doc = createPromptDoc_(f.getName(), f.getId());       // 30_prompts.gs
      const ss = createKouseiScaffold_(f.getName(), f.getId(), doc.getUrl()); // 40_kousei_sheet.gs
      upsertRow_(cfg, f.getId(), f.getName(), 'AI出力待ち', ss.getUrl(), '',
        'プロンプトDoc: ' + doc.getUrl() + ' → Claudeで実行し、構成シート「' + PASTE_TAB + '」タブへ貼り付けてください');
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
    if (!isImportRequested_(ss)) continue;

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