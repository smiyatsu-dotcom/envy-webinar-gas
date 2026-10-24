// Fixed, no-argument integration test. Properties are administrator-set once;
// chat, runner arguments, and case files cannot choose IDs or folders.
// This setup is intentionally separate from the runner: it only creates the
// synthetic fixture under the fixed isolation root.  It never reads a
// customer sheet, writes the regular input/output roots, or sets properties.
function スライド_隔離fixtureを作る() {
  const rootId = '1GUmxo4J9cj0WYdh4IPj2lBi3RSxiMw73';
  const caseName = '【隔離試験】架空配置永続化案件';
  const root = DriveApp.getFolderById(rootId);
  if (root.getName() !== '【隔離試験】配置永続化') throw new Error('隔離専用の出力rootではありません');
  const existing = filesByName_(root, '【隔離試験】構成シート_架空配置永続化');
  if (existing.hasNext()) {
    const file = existing.next();
    const ss = SpreadsheetApp.openById(file.getId());
    if (ss.getSheetByName(TABS.SETTINGS) && ss.getSheetByName(TABS.INTRO) &&
        ss.getSheetByName(TABS.PROBLEM) && ss.getSheetByName(TABS.STORY) && ss.getSheetByName(TABS.OFFER)) {
      return { created: false, sheetId: ss.getId(), sheetName: ss.getName(), caseName: caseName, outputRootId: rootId };
    }
    throw new Error('同名の隔離fixtureが生成仕様を満たしません。手動で確認してください');
  }
  const ss = SpreadsheetApp.create('【隔離試験】構成シート_架空配置永続化');
  DriveApp.getFileById(ss.getId()).moveTo(root);
  const settings = ss.getSheets()[0];
  settings.setName(TABS.SETTINGS);
  settings.getRange('A1:A5').setValues([['案件名'], ['承認'], ['ステータス'], ['ヒアリングファイルID'], ['素材フォルダID']]);
  settings.getRange('B1').setValue(caseName);
  settings.getRange('B2').setValue(true);
  settings.getRange('B3').setValue('隔離fixture（架空案件・顧客入力なし）');
  settings.getRange('B4').setValue('ISOLATION_SYNTHETIC_HEARING_ONLY');
  settings.getRange('B5').setValue('');
  writeContentTabs_(ss, {}, {}, {}, {});
  SpreadsheetApp.flush();
  return { created: true, sheetId: ss.getId(), sheetName: ss.getName(), caseName: caseName, outputRootId: rootId };
}

// 2026-10-23: 失われた旧fixtureの復元ではなく、新しい固定隔離用の再実行入口。
// 引数・任意ID・任意関数は受け付けず、通常 main/顧客root/通知を通さない。
var 隔離通し_生成_案件名 = '隔離通し診断テスト社';
var 隔離通し_生成_rootID = '1oQSEwyopPcvX20xuxNDHy5Kh4TfvtCBO';
var 隔離通し_生成_親rootID = '1GUmxo4J9cj0WYdh4IPj2lBi3RSxiMw73';
var 隔離通し_生成_root名 = '【隔離通し】診断から資料音声_20261023';
var 隔離通し_生成_台帳ID = '16UgS3scd3X2D2Qqhzd1WfJuGkuNcOt5WbekQZLyKDpg';
var 隔離通し_生成_開始箱ID = '1Mxch8d4SEbddMfx-C2slbmQWWa1Uz16u';
var 隔離通し_生成_生成物rootID = '18-xhLrskU8M9DvW2Rurqg5ST_NrOiSM2';
var 隔離通し_生成_ヒアリングID = '1ur8hPUElZMrKuitnDo1j1wzr54EkmRY9efNMaEQFErs';
var 隔離通し_生成_素材root名 = '【素材】投影資料';

function 隔離通し_生成の親を確かめる_(fileOrFolder, parentId, label) {
  var parents = fileOrFolder.getParents();
  if (!parents.hasNext() || parents.next().getId() !== parentId || parents.hasNext()) {
    throw new Error(label + ' が固定隔離root直下ではありません');
  }
}

function 隔離通し_生成の子フォルダ1件_(parent, name, label) {
  var it = parent.getFoldersByName(name);
  if (!it.hasNext()) throw new Error(label + ' が固定隔離root直下にありません');
  var folder = it.next();
  if (it.hasNext()) throw new Error(label + ' が固定隔離root直下に複数あります');
  隔離通し_生成の親を確かめる_(folder, parent.getId(), label);
  return folder;
}

function 隔離通し_生成設定を作る_() {
  var root = DriveApp.getFolderById(隔離通し_生成_rootID);
  if (root.getName() !== 隔離通し_生成_root名) throw new Error('隔離通しroot名が一致しません');
  隔離通し_生成の親を確かめる_(root, 隔離通し_生成_親rootID, '隔離通しroot');
  var startBox = DriveApp.getFolderById(隔離通し_生成_開始箱ID);
  隔離通し_生成の親を確かめる_(startBox, root.getId(), '隔離通し開始箱');
  if (startBox.getName() !== '【出力】') throw new Error('隔離通し開始箱名が一致しません');
  var output = DriveApp.getFolderById(隔離通し_生成_生成物rootID);
  隔離通し_生成の親を確かめる_(output, root.getId(), '隔離通し生成物root');
  if (output.getName() !== '【出力】生成物') throw new Error('隔離通し生成物root名が一致しません');
  var assets = 隔離通し_生成の子フォルダ1件_(root, 隔離通し_生成_素材root名, '隔離通し素材root');
  var ledgerFile = DriveApp.getFileById(隔離通し_生成_台帳ID);
  隔離通し_生成の親を確かめる_(ledgerFile, root.getId(), '隔離通し台帳');

  // 実行前に専用台帳がこの一件だけを新フローとして持つことを実物照合する。
  // これにより processNewFlow_ の通常走査が顧客案件へ広がらない。
  var rows = SpreadsheetApp.openById(隔離通し_生成_台帳ID).getSheets()[0].getDataRange().getValues();
  var matches = rows.slice(1).filter(function (row) {
    return String(row[0] || '').trim() === 隔離通し_生成_ヒアリングID &&
      String(row[1] || '').trim() === 隔離通し_生成_案件名 + '_ヒアリングシート' &&
      String(row[流れ_列 - 1] || '').trim() === 流れ_新;
  });
  if (matches.length !== 1 || rows.slice(1).filter(function (row) { return String(row[流れ_列 - 1] || '').trim() === 流れ_新; }).length !== 1) {
    throw new Error('隔離通し台帳が固定の新フロー1件ではありません');
  }

  // 通常設定からテンプレートだけを借りる。入出力・素材・台帳は必ず専用fixtureへ差し替える。
  var normal = getConfig_();
  if (normal.INPUT_FOLDER_ID === startBox.getId() || normal.OUTPUT_FOLDER_ID === output.getId() ||
      normal.ASSETS_FOLDER_ID === assets.getId() || normal.MGMT_SHEET_ID === 隔離通し_生成_台帳ID) {
    throw new Error('通常設定と隔離fixtureが分離されていません');
  }
  return Object.assign({}, normal, {
    ISOLATION_TEST: true,
    INPUT_FOLDER_ID: startBox.getId(),
    OUTPUT_FOLDER_ID: output.getId(),
    ASSETS_FOLDER_ID: assets.getId(),
    MGMT_SHEET_ID: 隔離通し_生成_台帳ID,
    ISOLATION_START_BOX_ID: startBox.getId()
  });
}

function 隔離通し_同名ファイル1件_(folder, name, label) {
  var it = folder.getFilesByName(name);
  if (!it.hasNext()) return null;
  var file = it.next();
  if (it.hasNext()) throw new Error(label + ' が複数あります');
  return file;
}

function 隔離通し_正規開始を引き渡す_(cfg) {
  var startBox = DriveApp.getFolderById(cfg.ISOLATION_START_BOX_ID);
  var startName = '【開始】' + 隔離通し_生成_案件名 + '.json';
  var start = 隔離通し_同名ファイル1件_(startBox, startName, '隔離通し開始JSON');
  var generatedRoot = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  if (!start) {
    // 前回は台帳の旧行によって登録だけがskipされた。新成果物が無いときだけ、
    // 同じ正規JSONを本文不変で開始状態へ戻して固定fixtureを再初期化できる。
    var pending = generatedRoot.getFoldersByName(隔離通し_生成_案件名);
    if (!pending.hasNext()) return { handed: false, reason: '開始箱に未処理の正規開始JSONがありません' };
    var fresh = pending.next();
    if (pending.hasNext()) throw new Error('隔離通し生成物の案件フォルダが複数あります');
    var done = 隔離通し_同名ファイル1件_(fresh, '【済】開始_' + 隔離通し_生成_案件名 + '.json', '隔離通し済開始JSON');
    var files = fresh.getFiles(); var count = 0;
    while (files.hasNext()) { files.next(); count++; }
    if (!done || count !== 1) throw new Error('隔離通しの新案件フォルダは未生成の開始JSONだけでなければ再初期化できません');
    done.setName(startName);
    start = done;
  }
  var payload = 開始を読む_(start.getBlob().getDataAsString('UTF-8'));
  if (payload.案件名 !== 隔離通し_生成_案件名 || payload.ヒアリングシートID !== 隔離通し_生成_ヒアリングID) {
    throw new Error('隔離通し開始JSONの案件またはヒアリングIDが固定fixtureと一致しません');
  }

  var old = generatedRoot.getFoldersByName(隔離通し_生成_案件名);
  if (old.hasNext()) {
    var prior = old.next();
    if (old.hasNext()) throw new Error('隔離通し生成物の案件フォルダが複数あります');
    var root = DriveApp.getFolderById(隔離通し_生成_rootID);
    var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
    var archive = root.createFolder('【再試験控え】' + 隔離通し_生成_案件名 + '_' + stamp);
    // 削除・上書きではなく、旧済開始・状態・成果物を丸ごと控えへ保全する。
    prior.moveTo(archive);
  }
  var destination = generatedRoot.createFolder(隔離通し_生成_案件名);
  // JSON本文は変えず、正規開始処理が読む案件フォルダへだけ移す。
  start.moveTo(destination);
  return { handed: true, startFileId: start.getId(), outputFolderId: destination.getId(), outputUrl: destination.getUrl() };
}

function 隔離通し_台帳を新規登録前に戻す_(cfg) {
  var sh = SpreadsheetApp.openById(cfg.MGMT_SHEET_ID).getSheets()[0];
  var rows = sh.getDataRange().getValues();
  var hits = [];
  rows.slice(1).forEach(function (row, index) {
    if (String(row[0] || '').trim() === 隔離通し_生成_ヒアリングID && String(row[1] || '').trim() === 隔離通し_生成_案件名 + '_ヒアリングシート') hits.push(index + 2);
  });
  if (hits.length !== 1) throw new Error('固定隔離案件の旧台帳行を一意に確認できません');
  // 専用台帳の固定一行だけを外し、通常 processStartOrders_ が構成Sheet・URL管理・指示Docを作り直す。
  sh.deleteRow(hits[0]);
}

/** 固定fixtureの正規開始JSONを既存登録処理へ安全に引き渡す。 */
function 隔離通し生成を開始する() {
  var cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function () {
    var handoff = 隔離通し_正規開始を引き渡す_(cfg);
    if (!handoff.handed) return Object.assign({ ok: true, phase: 'start' }, handoff);
    隔離通し_台帳を新規登録前に戻す_(cfg);
    processStartOrders_();
    return Object.assign({ ok: true, phase: 'start', caseName: 隔離通し_生成_案件名, outputRootId: cfg.OUTPUT_FOLDER_ID, managementSheetId: cfg.MGMT_SHEET_ID }, handoff);
  });
}

/**
 * 品質ルールを更新した後だけ使う固定隔離案件の再試験入口。
 * 直前の成果物は日時付き控えへ丸ごと保全し、診断が作成した開始JSONの
 * 本文を変えずに通常の processStartOrders_ へ渡し直す。通常案件、通知、
 * 通常の入力・出力rootを通らない。
 */
function 隔離通し品質修正後に再試験する() {
  var cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function () {
    var output = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
    var cases = output.getFoldersByName(隔離通し_生成_案件名);
    if (!cases.hasNext()) throw new Error('隔離通しの現行案件フォルダがありません');
    var current = cases.next();
    if (cases.hasNext()) throw new Error('隔離通しの現行案件フォルダが複数あります');
    隔離通し_生成の親を確かめる_(current, output.getId(), '隔離通し現行案件フォルダ');
    var doneName = '【済】開始_' + 隔離通し_生成_案件名 + '.json';
    var done = 隔離通し_同名ファイル1件_(current, doneName, '隔離通し済開始JSON');
    if (!done) throw new Error('隔離通しの現行案件に正規済開始JSONがありません');
    var payload = 開始を読む_(done.getBlob().getDataAsString('UTF-8'));
    if (payload.案件名 !== 隔離通し_生成_案件名 || payload.ヒアリングシートID !== 隔離通し_生成_ヒアリングID) {
      throw new Error('隔離通し済開始JSONの案件またはヒアリングIDが固定fixtureと一致しません');
    }
    var startBox = DriveApp.getFolderById(cfg.ISOLATION_START_BOX_ID);
    var startName = '【開始】' + 隔離通し_生成_案件名 + '.json';
    if (隔離通し_同名ファイル1件_(startBox, startName, '隔離通し開始JSON')) {
      throw new Error('隔離通し開始箱に未処理の開始JSONがあります');
    }
    var root = DriveApp.getFolderById(隔離通し_生成_rootID);
    var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
    var archive = root.createFolder('【再試験控え】' + 隔離通し_生成_案件名 + '_' + stamp);
    current.moveTo(archive);
    // 診断済み開始JSONは本文不変で取り出し、通常開始処理へ渡す。
    done.moveTo(startBox);
    done.setName(startName);
    var handoff = 隔離通し_正規開始を引き渡す_(cfg);
    if (!handoff.handed) throw new Error('隔離通しの正規開始JSONを引き渡せません');
    隔離通し_台帳を新規登録前に戻す_(cfg);
    processStartOrders_();
    return { ok: true, phase: 'quality-rerun-start', caseName: 隔離通し_生成_案件名,
      archiveFolderId: archive.getId(), outputFolderId: handoff.outputFolderId,
      startFileId: handoff.startFileId, managementSheetId: cfg.MGMT_SHEET_ID };
  });
}

// Apps Script editor function pickers can fail to index non-ASCII names after
// an external sync. Keep one no-argument, public ASCII entrypoint for this
// fixed fixture only; it delegates to the fully guarded Japanese function.
function runIsolationQualityRerun() {
  return 隔離通し品質修正後に再試験する();
}

/**
 * 固定fixtureの1件だけを通常の新フロー順で進める。
 * main の import → new-flow の順だけを再現する。開始の引渡し・台帳初期化はしない。
 */
function 隔離通し生成を進める() {
  var cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function () {
    // Match the normal five-minute main() pass: LP text is converted before
    // imports, then a queued slide build runs first in its own pass.  The
    // flag is execution-local; do not set the shared Script Property used by
    // production main(), because that could affect a customer case.
    convertLpTexts_();
    // This fixed fixture may resume after a validated watcher recovery.  Let
    // the normal importer re-evaluate its unchanged, already-validated JSON;
    // no production AUTOIMPORT key is touched.
    PropertiesService.getScriptProperties().deleteProperty('AUTOIMPORT_1YvJszV8Mz1tN1DVZ9xOIY5D7XGW754QQFVU_oyi34AQ');
    processImports_();
    新_先に組む回_ = true;
    processNewFlow_();
    return { ok: true, phase: 'advance', caseName: 隔離通し_生成_案件名, outputRootId: cfg.OUTPUT_FOLDER_ID, managementSheetId: cfg.MGMT_SHEET_ID };
  });
}

/**
 * Read-only probe for the one fixed isolation deck left by a failed build.
 * It deliberately neither renames nor saves the presentation, and is not
 * reachable from production callers.
 */
function 隔離通しスライド読取診断() {
  var cfg = 隔離通し_生成設定を作る_();
  var deckId = '11mln8JISLVwQKsjeVqHjctXeGEnk8UeFBRZqCBJvEYQ';
  return withIsolationTestConfig_(cfg, function () {
    var root = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
    if (root.getName() !== '【出力】生成物') throw new Error('隔離生成物root以外には診断できません');
    var f = DriveApp.getFileById(deckId);
    if (f.getMimeType() !== MimeType.GOOGLE_SLIDES || f.isTrashed()) throw new Error('隔離中間deckがありません');
    var parents = f.getParents();
    if (!parents.hasNext() || parents.next().getName() !== 隔離通し_生成_案件名) throw new Error('隔離案件外のdeckには診断できません');
    var pres = SlidesApp.openById(deckId);
    var slides = pres.getSlides();
    return { ok: true, deckId: deckId, name: f.getName(), slideCount: slides.length,
      pageWidth: pres.getPageWidth(), pageHeight: pres.getPageHeight() };
  });
}

/** Save/close probe for the same isolation-only deck; it performs no edits. */
function 隔離通しスライド保存診断() {
  var cfg = 隔離通し_生成設定を作る_();
  var deckId = '11mln8JISLVwQKsjeVqHjctXeGEnk8UeFBRZqCBJvEYQ';
  return withIsolationTestConfig_(cfg, function () {
    var f = DriveApp.getFileById(deckId);
    var parents = f.getParents();
    if (f.getMimeType() !== MimeType.GOOGLE_SLIDES || f.isTrashed() || !parents.hasNext() || parents.next().getName() !== 隔離通し_生成_案件名) throw new Error('隔離中間deckだけを診断します');
    var pres = SlidesApp.openById(deckId);
    var count = pres.getSlides().length;
    pres.saveAndClose();
    return { ok: true, deckId: deckId, slideCount: count, savedWithoutMutation: true };
  });
}

/**
 * Fixed-fixture counterpart of the existing manual "スライドだけ組み直す".
 * It invokes the same internal builder once, but refuses any ledger other
 * than this single isolated new-flow row.
 */
function 隔離通しスライドだけを組み直す() {
  var cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function () {
    var rows = mgmtSheet_(cfg).getDataRange().getValues();
    var matches = rows.slice(1).filter(function (row) {
      return String(row[0] || '') === 隔離通し_生成_ヒアリングID &&
        String(row[1] || '') === 隔離通し_生成_案件名 + '_ヒアリングシート' &&
        String(row[流れ_列 - 1] || '') === 流れ_新;
    });
    if (matches.length !== 1 || rows.slice(1).filter(function (row) { return String(row[流れ_列 - 1] || '') === 流れ_新; }).length !== 1) {
      throw new Error('隔離fixture以外の台帳行には再組版できません');
    }
    // This only bypasses the time-budget defer used by main(); it does not
    // touch the diagram-failure marker, approval state, or completion state.
    新_先に組む回_ = true;
    var ok = スライドを組む_(cfg, matches[0], false);
    if (ok !== true) throw new Error('正規スライド再組版が完了しませんでした: ' + ok);
    return { ok: true, caseName: 隔離通し_生成_案件名, mode: 'manual-slide-regeneration' };
  });
}

// J の最終台本を取り込むための、一度限りの固定隔離入口。
// 引数で案件・フォルダ・台帳行を選ばせず、指定された非顧客 fixture 以外には
// 一切書かない。通常 watcher / processNewFlow_ / 納品は呼ばない。
var 隔離通し_J台本_案件フォルダID = '1Ko7cbJ2nW2C4uQbZfDnZjW6mfz_qmHZ0';
var 隔離通し_J台本_ファイル名 = '【台本JSON】隔離通し診断テスト社.txt';
var 隔離通し_J台本_旧SHA256 = 'a3d6e2cd4eaaaf57dacf64818ca33486cac14d907b6de562a48868da9404918a';
var 隔離通し_J台本_SHA256 = '7eb361136315b71a1505d50e94cbeef12d28feb1846631ec67370072e5331645';
var 隔離通し_J台本_行数 = 106;
var 隔離通し_J台本_保全ファイル名 = '【保全】J台本取込前台本.txt';
var 隔離通し_J台本_記録ファイル名 = '【不確定】J台本取込記録.json';

function 隔離通し_J台本_sha256_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return bytes.map(function (value) {
    var hex = (value < 0 ? value + 256 : value).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

function 隔離通し_J台本_本文を検証する_(text) {
  if (隔離通し_J台本_sha256_(text) !== 隔離通し_J台本_SHA256) {
    throw new Error('J最終台本のSHA-256が指定済み本文と一致しません');
  }
  var rows;
  try { rows = JSON.parse(String(text)); } catch (e) { throw new Error('J最終台本がJSONとして読めません'); }
  if (!Array.isArray(rows) || rows.length !== 隔離通し_J台本_行数) throw new Error('J最終台本は106行の配列だけを受け付けます');
  var seen = {};
  rows.forEach(function (row) {
    if (!row || typeof row !== 'object' || Array.isArray(row) || Object.keys(row).sort().join(',') !== 'daihon,no') {
      throw new Error('J最終台本の行形式が違います');
    }
    var no = Number(row.no);
    if (!Number.isInteger(no) || no < 1 || no > 隔離通し_J台本_行数 || seen[no]) throw new Error('J最終台本のNo.が1〜106で一意ではありません');
    if (typeof row.daihon !== 'string' || !row.daihon.trim()) throw new Error('J最終台本に空のdaihonがあります: No.' + no);
    seen[no] = true;
  });
  for (var no = 1; no <= 隔離通し_J台本_行数; no++) if (!seen[no]) throw new Error('J最終台本にNo.' + no + 'がありません');
  return rows;
}

function 隔離通し_J台本_記録を読む_(folder) {
  var snapshot = 隔離通し_同名ファイル1件_(folder, 隔離通し_J台本_保全ファイル名, 'J台本取込前の保全');
  var record = 隔離通し_同名ファイル1件_(folder, 隔離通し_J台本_記録ファイル名, 'J台本取込の記録');
  if (!snapshot && !record) return null;
  var status = 'snapshot-only';
  if (record) {
    try { status = String(JSON.parse(record.getBlob().getDataAsString('UTF-8')).status || 'unknown'); }
    catch (e) { status = 'unreadable'; }
  }
  // A prior preparation, uncertain result, or completed one-shot run must
  // never be retried automatically.  A human can inspect this durable pair.
  throw new Error('J台本取込の保全または記録が既にあります（状態: ' + status + '）。自動再実行せず手動精査してください');
}

function 隔離通し_J台本_記録本文_(target, state, extra) {
  var value = {
    schema: 'isolation-j-script-import/v1',
    status: state,
    caseName: 隔離通し_生成_案件名,
    folderId: target.folder.getId(),
    hearingId: 隔離通し_生成_ヒアリングID,
    managementRow: target.ledger.rowNumber,
    registryRow: target.registryRow,
    kouseiSheetId: target.kouseiId,
    previousScriptSha256: target.currentHash,
    targetScriptSha256: 隔離通し_J台本_SHA256,
    recordedAt: new Date().toISOString()
  };
  Object.keys(extra || {}).forEach(function (key) { value[key] = extra[key]; });
  return JSON.stringify(value);
}

function 隔離通し_J台本_保全を作る_(target) {
  // Both records are written before the source txt, auto-import stamp, or
  // ledger can change. If either creation fails, its remaining counterpart
  // still makes the next automatic attempt fail closed.
  var snapshot = target.folder.createFile(隔離通し_J台本_保全ファイル名, target.currentText, MimeType.PLAIN_TEXT);
  var record;
  try {
    record = target.folder.createFile(隔離通し_J台本_記録ファイル名,
      隔離通し_J台本_記録本文_(target, 'prepared', { snapshotFileId: snapshot.getId() }), MimeType.PLAIN_TEXT);
  } catch (e) {
    throw new Error('J台本取込前の保全は作成されましたが記録を作成できません。自動再実行せず手動精査してください: ' + e);
  }
  return { snapshot: snapshot, record: record };
}

function 隔離通し_J台本_不確定にする_(target, record, cause) {
  var currentHash = '';
  try { currentHash = 隔離通し_J台本_sha256_(target.file.getBlob().getDataAsString('UTF-8')); } catch (ignore) { currentHash = 'unreadable'; }
  var details = String(cause && cause.message ? cause.message : cause).replace(/[\r\n]+/g, ' ').slice(0, 240);
  try {
    record.setContent(隔離通し_J台本_記録本文_(target, 'uncertain', {
      currentScriptSha256: currentHash,
      failure: details,
      automaticRetry: false
    }));
  } catch (recordError) {
    // The already-written "prepared" record remains as the fail-closed marker.
  }
  throw new Error('J台本取込の結果が不確定です。自動再実行せず、固定隔離案件の保全・記録を手動精査してください: ' + details);
}

function 隔離通し_J台本_固定対象を読む_(cfg) {
  var output = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  if (output.getId() !== 隔離通し_生成_生成物rootID || output.getName() !== '【出力】生成物') throw new Error('J台本取込の生成物rootが固定fixtureと一致しません');
  var folder = caseFolderIfExists_(cfg, 隔離通し_生成_案件名);
  if (!folder || folder.getId() !== 隔離通し_J台本_案件フォルダID || folder.getName() !== 隔離通し_生成_案件名) throw new Error('J台本取込の案件フォルダが固定隔離fixtureと一致しません');
  隔離通し_生成の親を確かめる_(folder, output.getId(), 'J台本取込の案件フォルダ');

  隔離通し_J台本_記録を読む_(folder);
  var sheet = mgmtSheet_(cfg);
  var rows = sheet.getDataRange().getValues();
  var dataRows = rows.slice(1).filter(function (row) { return row.some(function (cell) { return String(cell == null ? '' : cell).trim() !== ''; }); });
  if (dataRows.length !== 1) throw new Error('J台本取込の固定台帳は非空データ行が1件だけでなければなりません');
  var matched = [];
  var sameHearing = [];
  var importable = [];
  rows.slice(1).forEach(function (row, index) {
    var status = String(row[2] || '').trim();
    var exact = String(row[0] || '').trim() === 隔離通し_生成_ヒアリングID &&
      String(row[1] || '').trim() === 隔離通し_生成_案件名 + '_ヒアリングシート' &&
      String(row[流れ_列 - 1] || '').trim() === 流れ_新;
    if (exact) matched.push({ row: row, rowNumber: index + 2 });
    if (String(row[0] || '').trim() === 隔離通し_生成_ヒアリングID) sameHearing.push(index + 2);
    if (status === 'AI出力待ち' || status === 'レビュー待ち') importable.push({ exact: exact, rowNumber: index + 2 });
  });
  if (matched.length !== 1) throw new Error('J台本取込の固定台帳行を一意に確認できません');
  if (sameHearing.length !== 1) throw new Error('J台本取込のヒアリングIDが台帳で一意ではありません');
  if (importable.length && (importable.length !== 1 || !importable[0].exact)) throw new Error('J台本取込時に固定fixture以外の取込対象台帳行があります');
  var ledger = matched[0];
  var registryRow = registry_(cfg)[隔離通し_生成_ヒアリングID];
  if (registryRow !== ledger.rowNumber) throw new Error('J台本取込のregistry更新先が固定台帳行と一致しません');
  var kouseiId = idFromUrl_(String(ledger.row[3] || ''));
  var ss = SpreadsheetApp.openById(kouseiId);
  var settings = ss.getSheetByName(TABS.SETTINGS);
  if (!settings || String(settings.getRange('B1').getValue()).trim() !== 隔離通し_生成_案件名 + '_ヒアリングシート' ||
      String(settings.getRange('B4').getValue()).trim() !== 隔離通し_生成_ヒアリングID ||
      settings.getRange('B2').getValue() !== false || settings.getRange('B22').getValue() !== false) {
    throw new Error('J台本取込の固定構成シートが未承認・未納品のfixtureではありません');
  }
  var blocked = [];
  var files = folder.getFiles();
  while (files.hasNext()) {
    var name = nfc_(files.next().getName());
    if (name.indexOf('【指示】資料_') === 0 || name.indexOf('【承認】') === 0 || name.indexOf('【納品指示】') === 0) blocked.push(name);
  }
  if (blocked.length) throw new Error('J台本取込の隔離案件に未完了・承認・納品の印があります: ' + blocked.join(', '));
  var file = 隔離通し_同名ファイル1件_(folder, 隔離通し_J台本_ファイル名, 'J台本JSON');
  if (!file) throw new Error('J台本取込の固定台本JSONがありません');
  var current = file.getBlob().getDataAsString('UTF-8');
  var currentHash = 隔離通し_J台本_sha256_(current);
  if (currentHash !== 隔離通し_J台本_旧SHA256) throw new Error('J台本取込の既存台本が保全した旧版と一致しません');
  return { folder: folder, file: file, ss: ss, ledger: ledger, kouseiId: kouseiId, currentHash: currentHash, currentText: current, registryRow: registryRow };
}

/** 読み取りだけの事前監査。取込・再生成・納品・通知はしない。 */
function 隔離通し_J台本取込を検査() {
  var cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function () {
    var target = 隔離通し_J台本_固定対象を読む_(cfg);
    return { ok: true, caseName: 隔離通し_生成_案件名, folderId: target.folder.getId(),
      managementRow: target.ledger.rowNumber, registryRow: target.registryRow, kouseiSheetId: target.kouseiId,
      currentScriptSha256: target.currentHash, expectedNextScriptSha256: 隔離通し_J台本_SHA256,
      deliveryOrApproval: false, mode: 'read-only-preflight' };
  });
}

/**
 * 指定済みSHA-256のJ最終台本だけを、固定隔離案件の既存JSONへ反映して
 * 正規 processImports_ に一度渡す。processNewFlow_・watcher・納品は呼ばない。
 */
function 隔離通し_J台本を正規取込する(daihonText) {
  隔離通し_J台本_本文を検証する_(daihonText);
  var cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function () {
    var target = 隔離通し_J台本_固定対象を読む_(cfg);
    var preservation = 隔離通し_J台本_保全を作る_(target);
    try {
      // 現物の旧版を保全・記録してから、同一の固定隔離txtだけを置き換える。
      target.file.setContent(String(daihonText));
      PropertiesService.getScriptProperties().deleteProperty('AUTOIMPORT_' + target.kouseiId);
      upsertRow_(cfg, 隔離通し_生成_ヒアリングID, 隔離通し_生成_案件名 + '_ヒアリングシート', 'AI出力待ち', String(target.ledger.row[3]), '', 'J最終台本の固定隔離取込');
      if (registry_(cfg)[隔離通し_生成_ヒアリングID] !== target.ledger.rowNumber) throw new Error('J台本取込前にregistry更新先が変わりました');
      processImports_();
      var after = mgmtSheet_(cfg).getDataRange().getValues()[target.ledger.rowNumber - 1];
      if (String(after[2] || '').trim() !== 'レビュー待ち') throw new Error('J最終台本を正規取込後にレビュー待ちへ遷移しません');
      var settings = target.ss.getSheetByName(TABS.SETTINGS);
      if (!settings || settings.getRange('B2').getValue() !== false || settings.getRange('B22').getValue() !== false) {
        throw new Error('J最終台本取込が承認または納品状態を変更しました');
      }
      preservation.record.setContent(隔離通し_J台本_記録本文_(target, 'completed', {
        snapshotFileId: preservation.snapshot.getId(), currentScriptSha256: 隔離通し_J台本_sha256_(daihonText), automaticRetry: false
      }));
      return { ok: true, caseName: 隔離通し_生成_案件名, folderId: target.folder.getId(), managementRow: target.ledger.rowNumber,
        scriptSha256: 隔離通し_J台本_sha256_(daihonText), status: 'レビュー待ち', mode: 'processImports-only' };
    } catch (e) {
      return 隔離通し_J台本_不確定にする_(target, preservation.record, e);
    }
  });
}

function 隔離通しスライド失敗を読む() {
  var cfg = 隔離通し_生成設定を作る_();
  return withIsolationTestConfig_(cfg, function () {
    var rows = mgmtSheet_(cfg).getDataRange().getValues();
    var row = rows.slice(1).filter(function (r) { return String(r[0] || '') === 隔離通し_生成_ヒアリングID; })[0];
    if (!row) throw new Error('隔離台帳行がありません');
    return { status: String(row[2] || ''), message: String(row[5] || ''), slideUrl: String(row[4] || '') };
  });
}

function isolationUntouchedFingerprint_(slide, omit) {
  return slide.getPageElements().filter(function(e){return !omit(e);}).map(function(e){
    return JSON.stringify({label:elementLabel_(e),geometry:geometryOf_(e)});
  }).sort();
}

function スライド_隔離配置永続化を試す() {
  const p=PropertiesService.getScriptProperties();
  const sheetId=p.getProperty('GEOMETRY_ISOLATION_KOUSEI_ID');
  const rootId=p.getProperty('GEOMETRY_ISOLATION_OUTPUT_ROOT_ID');
  const caseName=p.getProperty('GEOMETRY_ISOLATION_CASE_NAME');
  if(!sheetId||!rootId||!caseName) throw new Error('隔離配置試験の固定fixture設定がありません');
  const prod=getConfig_();
  if(rootId===prod.OUTPUT_FOLDER_ID) throw new Error('顧客の【出力】生成物は隔離試験に使えません');
  const root=DriveApp.getFolderById(rootId);
  if(root.getName()!=='【隔離試験】配置永続化') throw new Error('隔離専用の出力rootではありません');
  const ss=SpreadsheetApp.openById(sheetId);
  if(String(ss.getName()).indexOf('【隔離試験】')!==0) throw new Error('隔離fixtureの構成シートではありません');
  const delivery=DriveApp.getFolderById(prod.OUTPUT_FOLDER_ID).getParents().next().getFoldersByName('【納品】');
  if(delivery.hasNext()&&delivery.next().getId()===rootId) throw new Error('納品フォルダを隔離試験に使えません');
  const folder=caseFolder_(Object.assign({},prod,{OUTPUT_FOLDER_ID:rootId}),caseName);
  // A prior isolated run may have left its own test record.  Neutralize only
  // that fixed file before recording the first deck's untouched template
  // geometry; customer records and all other case files remain untouched.
  const geometryName='【配置指定】'+caseCompany_(caseName)+'.json';
  const priorGeometry=filesByName_(folder,geometryName);
  if(priorGeometry.hasNext()){const f=priorGeometry.next();if(priorGeometry.hasNext())throw new Error('隔離配置指定が複数あります');f.setContent(JSON.stringify({schema:GEOMETRY_SCHEMA,caseName:caseCompany_(caseName),records:[]}));}
  const first=generateSlides_(sheetId,{安全:true,isolationOutputFolderId:rootId});
  const mapFile=filesByName_(folder,'【スライド対応】'+caseCompany_(caseName)+'.json');
  if(!mapFile.hasNext()) throw new Error('1回目の対応表がありません');
  const map=JSON.parse(mapFile.next().getBlob().getDataAsString('UTF-8'));
  if(map.presentationId!==first.pres.getId()||!Array.isArray(map.pages)||!map.pages.length) throw new Error('1回目の対応表がdeckと一致しません');
  // Author one deliberately small, fixed-only layout change from the live first
  // deck.  The selector carries template geometry, so the second deck must
  // prove the same template object before the override can apply.
  const page1=map.pages.filter(function(x){return x.slideNo===1;})[0];
  const slide1=page1&&first.pres.getSlides().filter(function(s){return s.getObjectId()===page1.pageId;})[0];
  if(!slide1) throw new Error('No.1の生成済み対応を読めません');
  let target=null, selector=null;
  slide1.getPageElements().some(function(e){
    try {
      const g=geometryOf_(e), s=Object.assign(elementLabel_(e),{templateGeometry:g});
      if(!validGeometry_(g)) return false;
      if(slide1.getPageElements().filter(function(x){return selectorMatches_(x,s);}).length!==1) return false;
      target=e; selector=s; return true;
    } catch(ignore) { return false; }
  });
  if(!target) throw new Error('No.1に一意な配置検証対象がありません');
  const after=Object.assign({},selector.templateGeometry,{left:selector.templateGeometry.left+1});
  const untouchedBefore=map.pages.map(function(page){
    const slide=first.pres.getSlides().filter(function(s){return s.getObjectId()===page.pageId;})[0];
    if(!slide) throw new Error('1回目の対象外要素を読めません: No.'+page.slideNo);
    return {slideNo:page.slideNo,fingerprint:isolationUntouchedFingerprint_(slide,function(e){return page.slideNo===1&&selectorMatches_(e,selector);})};
  });
  const geometryText=JSON.stringify({schema:GEOMETRY_SCHEMA,caseName:caseCompany_(caseName),records:[{status:'active',slideNo:1,selector:selector,after:after}]});
  const oldGeometry=filesByName_(folder,geometryName);
  if(oldGeometry.hasNext()){const f=oldGeometry.next();if(oldGeometry.hasNext())throw new Error('隔離配置指定が複数あります');f.setContent(geometryText);}else{folder.createFile(geometryName,geometryText,MimeType.PLAIN_TEXT);}
  const second=generateSlides_(sheetId,{安全:true,isolationOutputFolderId:rootId});
  const latest=JSON.parse(filesByName_(folder,'【スライド対応】'+caseCompany_(caseName)+'.json').next().getBlob().getDataAsString('UTF-8'));
  if(latest.presentationId!==second.pres.getId()||latest.presentationId===map.presentationId) throw new Error('2回目の対応表を読み戻せません');
  if(latest.pages.length!==map.pages.length||latest.pages.map(function(p){return p.slideNo;}).join(',')!==map.pages.map(function(p){return p.slideNo;}).join(',')) throw new Error('対象外ページ数または順序が変わりました');
  const page2=latest.pages.filter(function(x){return x.slideNo===1;})[0];
  const slide2=page2&&second.pres.getSlides().filter(function(s){return s.getObjectId()===page2.pageId;})[0];
  // selectorMatches_ is deliberately used only before mutation. Once applied,
  // templateGeometry is no longer expected to match. Read the one post-apply
  // element by its stable labels plus its exact requested final geometry.
  const labeled=slide2?slide2.getPageElements().filter(function(e){var l=elementLabel_(e);return l.type===selector.type&&l.title===selector.title&&l.description===selector.description;}):[];
  // The SlidesApp getters and setters both expose point values.  Do not hide a
  // transform/round-trip difference behind an arbitrary epsilon: capture the
  // exact readback and require numeric equality for every requested field.
  const readbacks=labeled.map(function(e){return geometryOf_(e);});
  const applied=labeled.filter(function(e){var g=geometryOf_(e);return ['left','top','width','height'].every(function(k){return Number(g[k])===Number(after[k]);});});
  if(applied.length!==1) throw new Error('2回目の現物geometryが承認済み配置と一致しません: expected='+JSON.stringify(after)+', actual='+JSON.stringify(readbacks)+', labels='+labeled.length+', expectedGeometry='+applied.length);
  untouchedBefore.forEach(function(before){
    const page=latest.pages.filter(function(p){return p.slideNo===before.slideNo;})[0];
    const slide=page&&second.pres.getSlides().filter(function(s){return s.getObjectId()===page.pageId;})[0];
    if(!slide) throw new Error('2回目の対象外要素を読めません: No.'+before.slideNo);
    const actual=isolationUntouchedFingerprint_(slide,function(e){
      if(before.slideNo!==1) return false;
      const l=elementLabel_(e),g=geometryOf_(e);
      return l.type===selector.type&&l.title===selector.title&&l.description===selector.description&&['left','top','width','height'].every(function(k){return Number(g[k])===Number(after[k]);});
    });
    if(JSON.stringify(actual)!==JSON.stringify(before.fingerprint)) throw new Error('対象外要素のgeometryまたは構成が変わりました: No.'+before.slideNo);
  });
  return {ok:true,caseName:caseName,outputRootId:rootId,firstDeckId:map.presentationId,secondDeckId:latest.presentationId,pageCount:latest.pages.length,expectedGeometry:after,actualGeometry:geometryOf_(applied[0]),geometryVerified:true};
}


// Fixed fixture only: snapshot, mutate the live object, prove the readback,
// then persist the selector for a normal regeneration.  No input selects IDs.
const ISOLATION_GEOMETRY_CASE = '隔離通し診断テスト社';
const ISOLATION_GEOMETRY_HEARING = '1ur8hPUElZMrKuitnDo1j1wzr54EkmRY9efNMaEQFErs';
const ISOLATION_GEOMETRY_DECK = '12jfYwhluLE4r6DQrueSl45x8BnZprbaMtgJBVCugU84';
const ISOLATION_GEOMETRY_PAGE = 'g38554c11d2e_1_113';
const ISOLATION_GEOMETRY_OBJECT = 'g38554c11d2e_1_114';
const ISOLATION_GEOMETRY_NO = 40;
const ISOLATION_GEOMETRY_BEFORE = {left:40.44094488188976,top:37.21555118110236,width:639.1181102362204,height:47.40944881889764};
function isolationGeometryEqual_(a,b){return !!a&&!!b&&['left','top','width','height'].every(function(k){return Number(a[k])===Number(b[k]);});}
function isolationGeometry_(e){return {left:e.getLeft(),top:e.getTop(),width:e.getWidth(),height:e.getHeight()};}
function isolationHash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8).map(function(b){return ('0'+(b<0?b+256:b).toString(16)).slice(-2);}).join('');}
function isolationCanonical_(v){if(Array.isArray(v))return '['+v.map(isolationCanonical_).join(',')+']';if(v&&typeof v==='object')return '{'+Object.keys(v).sort().map(function(k){return JSON.stringify(k)+':'+isolationCanonical_(v[k]);}).join(',')+'}';return JSON.stringify(v);}
function isolationPlanHash_(p){const q={};Object.keys(p).forEach(function(k){if(k!=='planHash')q[k]=p[k];});return isolationHash_(isolationCanonical_(q));}
function isolationFolder_(){const root=DriveApp.getFolderById(隔離通し_生成_生成物rootID);if(root.getName()!=='【出力】生成物')throw new Error('固定隔離rootではありません');const it=root.getFoldersByName(ISOLATION_GEOMETRY_CASE);if(!it.hasNext())throw new Error('固定隔離案件がありません');const f=it.next();if(it.hasNext())throw new Error('固定隔離案件が複数あります');return f;}
function isolationGeometryMap_(){const f=activeGeometryFile_(isolationFolder_(),ISOLATION_GEOMETRY_CASE);if(!f)return null;const x=JSON.parse(f.getBlob().getDataAsString('UTF-8'));if(!x||x.schema!==GEOMETRY_SCHEMA||x.caseName!==ISOLATION_GEOMETRY_CASE||!Array.isArray(x.records))throw new Error('固定隔離配置指定が不正です');return {file:f,data:x};}
function スライド_承認用geometryを読む(){
  const folder=isolationFolder_(),it=filesByName_(folder,GEOMETRY_MAP_PREFIX+ISOLATION_GEOMETRY_CASE+'.json');if(!it.hasNext())throw new Error('正規対応表がありません');const mapFile=it.next();if(it.hasNext())throw new Error('正規対応表が複数あります');const map=JSON.parse(mapFile.getBlob().getDataAsString('UTF-8'));
  const p=(map.pages||[]).filter(function(x){return x.slideNo===ISOLATION_GEOMETRY_NO&&x.pageId===ISOLATION_GEOMETRY_PAGE;});if(map.caseName!==ISOLATION_GEOMETRY_CASE||map.presentationId!==ISOLATION_GEOMETRY_DECK||p.length!==1)throw new Error('固定対象と正規対応表が一致しません');
  const deck=SlidesApp.openById(ISOLATION_GEOMETRY_DECK),slide=deck.getSlides().filter(function(s){return s.getObjectId()===ISOLATION_GEOMETRY_PAGE;});if(slide.length!==1)throw new Error('正規ページが現物と一致しません');const hit=slide[0].getPageElements().filter(function(e){return e.getObjectId()===ISOLATION_GEOMETRY_OBJECT;});if(hit.length!==1)throw new Error('対象要素が一意ではありません');const g=isolationGeometry_(hit[0]);if(!isolationGeometryEqual_(g,ISOLATION_GEOMETRY_BEFORE))throw new Error('対象要素の元geometryが一致しません');
  const selector=Object.assign(elementLabel_(hit[0]),{templateGeometry:g});return {ok:true,caseName:ISOLATION_GEOMETRY_CASE,hearingId:ISOLATION_GEOMETRY_HEARING,presentationId:ISOLATION_GEOMETRY_DECK,slideNo:ISOLATION_GEOMETRY_NO,pageId:ISOLATION_GEOMETRY_PAGE,objectId:ISOLATION_GEOMETRY_OBJECT,before:g,persistentSelector:selector,revision:isolationHash_(isolationCanonical_({map:map,geometry:g,selector:selector}))};
}
function スライド_承認済み配置を適用(plan){
  if(!plan||plan.caseName!==ISOLATION_GEOMETRY_CASE||!/^\d{12}-[a-f0-9]{4}$/.test(String(plan.proposalId||''))||plan.revision!==1||plan.planHash!==isolationPlanHash_(plan))throw new Error('固定隔離の承認済み試験planではありません');
  const snap=スライド_承認用geometryを読む(),a=plan.after;if(plan.presentationId!==snap.presentationId||plan.pageId!==snap.pageId||plan.objectId!==snap.objectId||!isolationGeometryEqual_(plan.before,snap.before)||plan.presentationRevision!==snap.revision||isolationCanonical_(plan.persistentSelector)!==isolationCanonical_(snap.persistentSelector))throw new Error('対象、元revision、またはselectorが一致しません');if(!a||a.left!==snap.before.left||a.width!==snap.before.width||a.height!==snap.before.height||!Number.isFinite(a.top)||a.top<20||a.top>55)throw new Error('許可された小さなtop変更ではありません');
  const existing=isolationGeometryMap_();if(existing&&existing.data.records.some(function(r){return r&&r.proposalId===plan.proposalId;}))throw new Error('この試験planは既に適用済みです');
  const deck=SlidesApp.openById(snap.presentationId),slide=deck.getSlides().filter(function(s){return s.getObjectId()===snap.pageId;});const hits=slide.length===1?slide[0].getPageElements().filter(function(e){return e.getObjectId()===snap.objectId;}):[];if(hits.length!==1)throw new Error('現物の固定対象が一意ではありません');const target=hits[0];
  target.setLeft(a.left).setTop(a.top).setWidth(a.width).setHeight(a.height);const actual=isolationGeometry_(target);if(!isolationGeometryEqual_(actual,a)){target.setLeft(snap.before.left).setTop(snap.before.top).setWidth(snap.before.width).setHeight(snap.before.height);throw new Error('現物geometryの読戻しが一致しません');}
  const record={proposalId:plan.proposalId,revision:1,planHash:plan.planHash,status:'active',slideNo:ISOLATION_GEOMETRY_NO,selector:plan.persistentSelector,after:a};const data=existing?existing.data:{schema:GEOMETRY_SCHEMA,caseName:ISOLATION_GEOMETRY_CASE,records:[]};data.records.push(record);
  try{if(existing)existing.file.setContent(JSON.stringify(data));else isolationFolder_().createFile(GEOMETRY_FILE_PREFIX+ISOLATION_GEOMETRY_CASE+'.json',JSON.stringify(data),MimeType.PLAIN_TEXT);const saved=isolationGeometryMap_();if(!saved||saved.data.records.filter(function(r){return r&&r.proposalId===plan.proposalId&&r.planHash===plan.planHash;}).length!==1)throw new Error('配置指定の読戻しが一致しません');}catch(e){target.setLeft(snap.before.left).setTop(snap.before.top).setWidth(snap.before.width).setHeight(snap.before.height);throw e;}
  return {ok:true,proposalId:plan.proposalId,revision:1,planHash:plan.planHash,before:snap.before,after:actual,geometryVerified:true,durableGeometryVerified:true};
}

