/**
 * 15_new_flow.gs（資料生成の新しい流れ・2026-09-26）
 * 設計の正本: operations リポジトリ tools/資料生成_新しい流れ.md（10. ファイルのつなぎ が受け渡しの正本）
 *
 * ★「流れ＝新」（管理シート H列が「新」）の案件だけを動かす。H列が空の案件（旧の流れ）には
 *   ここから1つも触らない。旧の流れの関数（10_main.js ほか）は、H列を見て新の案件だけ避ける。
 *
 * 新の流れの順番:
 *   【開始】◯◯社.json（案件の下地.gs が置く）→ processStartOrders_ が登録し、プロンプトDoc・構成シート・
 *   【指示】LP_・【指示】資料_ を作る（【入力】もヒアリングシートの複製も通さない）
 *   → LPができたら1位のタイトルを採用して【保留】サムネ指示 → 【指示】サムネ_（人は押さない）
 *   → 取り込み（AIの出力）と【図解済み】.json がそろったら、承認を待たずにスライドを組む（音声は見張りが作る）
 *   → 全部できたら「確認して承認」（操作パネルが置く【承認】◯◯社.json か B2）
 *   → そのあとだけ「納品フォルダへ入れる」（操作パネルが置く【納品指示】◯◯社.json か B22）
 *   → 直しは PM相談室 → ボットが【指示】直し_ → Mac が直して【済】直し_ → スライドなら、ここが音声を残して組み直す
 *
 * 進みは 16_progress.js が 【状態】進み.json に書く（書くのはこのスクリプトだけ）。
 * ★操作パネルは別のGASプロジェクト（~/webinar-panel-gas）。このスクリプトの権限（appsscript.json）は変えない
 *   （権限を足すと再許可まで時間トリガーが止まり、進行中の案件の生成が止まるため・2026-09-26）。
 *   だからこのスクリプトは外へ送らない（UrlFetchApp を使わない）。知らせは【知らせ待ち】に置き、Mac が送る。
 */

const 流れ_列 = 8;          // 管理シート H列
const 流れ_新 = '新';
const 開始_頭 = '【開始】';
const 開始_済の頭 = '【済】開始_';
const 開始_失敗の上限 = 3;   // 読めても登録に失敗したとき、何回まで次の回にやり直すか

/** 管理シートの1行（getValues の1行）が新の流れか。★H列が空なら旧の流れ（今までどおり） */
function 新の流れか_(row) {
  return String((row && row[流れ_列 - 1]) || '').trim() === 流れ_新;
}

/** 台帳の案件名（◯◯社）→ このスクリプトで持つ案件名（◯◯社_ヒアリングシート。素材フォルダ名と同じ） */
function 新の流れの案件名_(案件名) {
  const 名 = String(案件名 || '').trim();
  if (!名) return '';
  return /ヒアリングシート\s*$/.test(名) ? 名 : 名 + '_ヒアリングシート';
}

/**
 * 【開始】◯◯社.json の中身を読む。形がおかしければ日本語で投げる（呼ぶ側は次の回に回す）。
 * 形の正本: 設計書 10.「【開始】◯◯社.json」
 */
function 開始を読む_(text) {
  let o;
  try { o = JSON.parse(String(text || '')); } catch (e) { throw new Error('【開始】のファイルがJSONとして読めません'); }
  if (!o || typeof o !== 'object') throw new Error('【開始】のファイルの中身が空です');
  const 案件名 = String(o['案件名'] || '').trim();
  const hid = String(o['ヒアリングシートID'] || '').trim();
  if (!案件名) throw new Error('【開始】のファイルに「案件名」がありません');
  const m = hid.match(/[-\w]{25,}/);
  if (!m) throw new Error('【開始】のファイルの「ヒアリングシートID」が読めません');
  return { 案件名: 案件名, ヒアリングシートID: m[0], 台帳の行: o['台帳の行'] || '', 置いた日時: String(o['置いた日時'] || '') };
}

/** ファイル名が【開始】◯◯社.json か（【済】開始_ は除く） */
function 開始ファイルか_(name) {
  const n = nfc_(name);
  return n.indexOf(開始_頭) === 0 && /\.json$/i.test(n);
}

/** 管理シートの H1 に見出し「流れ」が無ければ足す（旧の行には何も書かない） */
function 流れの見出しを整える_(sh) {
  try { if (String(sh.getRange(1, 流れ_列).getValue()) !== '流れ') sh.getRange(1, 流れ_列).setValue('流れ').setFontWeight('bold'); }
  catch (e) { /* 見出しは体裁。失敗しても止めない */ }
}

/** 管理シートの、そのヒアリングシートIDの行の H列を書く */
function 流れを書く_(cfg, hearingId, 値) {
  const reg = registry_(cfg);
  if (!reg[hearingId]) return false;
  const sh = mgmtSheet_(cfg);
  流れの見出しを整える_(sh);
  sh.getRange(reg[hearingId], 流れ_列).setValue(値);
  return true;
}

// ============================================================
// 新の流れ ①: 【開始】ファイル → 登録（【入力】を経由しない）
// ============================================================
function processStartOrders_() {
  const cfg = getConfig_();
  const known = registry_(cfg);
  const folders = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID).getFolders();
  const P = PropertiesService.getScriptProperties();

  while (folders.hasNext()) {
    const folder = folders.next();
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (!開始ファイルか_(f.getName())) continue;

      let 開始;
      try { 開始 = 開始を読む_(f.getBlob().getDataAsString('UTF-8')); }
      catch (e) { Logger.log('【開始】を読めません（次の回にもう一度読みます）: ' + f.getName() + ' / ' + e); continue; }

      const hid = 開始.ヒアリングシートID;
      const 済の名前 = 開始_済の頭 + caseCompany_(開始.案件名) + '.json';
      if (known[hid]) {   // 押し直し・二重置き。登録し直さない
        try { f.setName(済の名前); } catch (e) { /* 次の回にまた名前を変える */ }
        continue;
      }

      const pname = 新の流れの案件名_(開始.案件名);
      const 失敗の鍵 = 'START_FAIL_' + hid;
      try {
        const doc = createPromptDoc_(pname, hid);                    // 30_prompts.js（【指示】LP_・【指示】資料_ も置く）
        const ss = createKouseiScaffold_(pname, hid, doc.getUrl());  // 40_kousei_sheet.js
        try { createUrlSheet_(cfg, pname); } catch (e) { Logger.log('URL管理シートの作成に失敗: ' + e); }
        upsertRow_(cfg, hid, pname, 'AI出力待ち', ss.getUrl(), '',
          '新しい流れ：自動で作っています（LP・サムネイル・構成と台本・図解・スライド・音声）。進みは操作パネルで見られます');
        流れを書く_(cfg, hid, 流れ_新);
        known[hid] = true;
        try { 新の流れの設定タブを整える_(ss.getSheetByName(TABS.SETTINGS), folder.getId()); } catch (e) { /* 体裁 */ }
        P.deleteProperty(失敗の鍵);
        進み_記録を変える_(hid, function (記) { 動きを足す_(記, '始めました（③診断OK）'); });
        f.setName(済の名前);
      } catch (e) {
        const n = Number(P.getProperty(失敗の鍵) || 0) + 1;
        P.setProperty(失敗の鍵, String(n));
        Logger.log('★新しい流れの登録に失敗（' + n + '回目）: ' + pname + ' / ' + e);
        if (n >= 開始_失敗の上限) {
          upsertRow_(cfg, hid, pname, 'エラー', '', '', '新しい流れの登録に失敗: ' + String(e));
          流れを書く_(cfg, hid, 流れ_新);
          known[hid] = true;
          P.deleteProperty(失敗の鍵);
          try { f.setName(済の名前); } catch (e2) { /* 次の回 */ }
          進み_記録を変える_(hid, function (記) { 知らせを積む_(記, pname, '始められませんでした（' + 開始_失敗の上限 + '回やり直しても登録できません）。' + String(e)); });
        }
      }
    }
  }
}

// ============================================================
// 新の流れ ②: 1件ずつ進める（5分ごと）
// ============================================================
function processNewFlow_() {
  const cfg = getConfig_();
  const rows = mgmtSheet_(cfg).getDataRange().getValues();
  const 一回 = { 組んだ: false };   // スライドを組むのは1回の実行で1件まで（6分の上限を守る）
  for (let i = 1; i < rows.length; i++) {
    if (!新の流れか_(rows[i])) continue;   // ★旧の流れの行には触らない
    const hearingId = String(rows[i][0] || '');
    const projectName = String(rows[i][1] || '');
    if (!hearingId || !projectName) continue;
    // ★Mac の見張りの心拍。新の流れで「まだ納品していない」案件があるときだけ、1回の実行で1度だけ読む
    //   （納品済みの古い案件やテスト案件だけのときに、Mac が眠るたび知らせが出ないように。2026-09-27）
    if (一回.見張り === undefined) 一回.見張り = 見張りの様子_(見張りの心拍を読む_(cfg), new Date());   // 画面に出すのは全部の案件
    if (!一回.知らせた && !テスト案件か_(projectName)) {
      const 記0 = 進み_記録を読む_(hearingId);
      if (!(記0.納品 && 記0.納品.済)) { 一回.知らせた = true; 見張りが止まったら知らせる_(一回.見張り); }
    }
    try { 新の流れを1件進める_(cfg, rows[i], 一回); }
    catch (e) { Logger.log('★新しい流れの処理で失敗: ' + projectName + ' / ' + e); }
  }
}

function 新の流れを1件進める_(cfg, row, 一回) {
  const hearingId = String(row[0]);
  const projectName = String(row[1]);
  let status = String(row[2] || '');
  const kouseiUrl = String(row[3] || '');
  const company = caseCompany_(projectName);

  const 記 = 進み_記録を読む_(hearingId);
  // ★「スライド生成中」のまま20分を超えたら、6分の上限で途中で止まったとみなしてエラーに戻す（止まったと出し、組み直せるようにする。2026-09-27）
  if (status === 'スライド生成中' && row[6] && Date.now() - new Date(row[6]).getTime() > 20 * 60000) {
    upsertRow_(cfg, hearingId, projectName, 'エラー', kouseiUrl, '', 'スライドを組むのが途中で止まりました（時間の上限）。次の回に組み直します');
    status = 'エラー';
    // 印を変えて、次の回に組み直させる（初めて組む回で止まったときも）。
    // ★2回続けて止まったら自動では組み直さない（20分待って止まるのをくり返さない）。止まったと出るので、PM相談室へ頼むと組み直す
    記.途中で止まった回 = (記.途中で止まった回 || 0) + 1;
    記.組んだ図解 = 記.途中で止まった回 >= 2 ? '' : '（途中で止まったので組み直す）';
    if (記.途中で止まった回 >= 2) 動きを足す_(記, 'スライドを組むのが2回続けて時間の上限で止まりました。PM相談室へ頼むか、宮津さんへ相談してください');
  }
  // 納品して14日たった案件は見に行かない（毎回すべての構成シートを開かないため）
  if (記.納品 && 記.納品.済 && 記.納品.日時 && Date.now() - Date.parse(記.納品.日時) > 14 * 86400000) return;

  let ss = null, st = null;
  if (kouseiUrl) {
    try { ss = SpreadsheetApp.openById(idFromUrl_(kouseiUrl)); st = ss.getSheetByName(TABS.SETTINGS); } catch (e) { ss = null; st = null; }
  }
  const folder = caseFolderIfExists_(cfg, projectName);
  if (!folder) return;
  if (st) 新の流れの設定タブを整える_(st, folder.getId());

  let 観 = 観測する_(cfg, row, folder);
  観.見張り = 一回.見張り || null;

  // (1) タイトルは1位を自動採用 → サムネイルへ（旧の流れの B19 を人が押す代わり）
  // ★LPの仕事が終わってから（【指示】LP_ が消え、【作業中】LP も無い）。見張りはLPの仕事の終わりに
  //   「【保留】サムネ指示_ ができたか」で成否を見るので、その前に名前を変えると失敗扱いになりLPを作り直してしまう
  const LPの仕事中 = 観.ファイル.指示LP || (観.作業中 || []).some(function (w) { return w['仕事'] === 'LP'; });
  if (観.ファイル.タイトル案 && 観.保留サムネ && !LPの仕事中) {
    try {
      観.保留サムネ.setName('【指示】サムネ_' + company + '.txt');
      動きを足す_(記, 'タイトルは1位を採用して、サムネイルを作り始めました');
      if (st) st.getRange('B3').setValue('サムネイル: タイトルは1位を採用して作り始めました（10分ほど）');
    } catch (e) { Logger.log('サムネ指示の名前を変えられません: ' + e); }
  }

  // (2) 取り込みは processImports_ が新の流れの文で行う（10_main.js）。取り込めたら覚えておく
  if (status === 'レビュー待ち' || status === 'スライド生成中' || status === '完了') 記.取り込み済み = true;

  // ★前の記録は (2b) より前に取る。(2b) で足した直しも「新しく受けた」に数えて、承認済みなら外すため（2026-09-27）
  const 前の直し = JSON.stringify(記.直し || []);

  // (2b) Mac の【直すセル】（スライド・台本の直し）: 構成シートへ書いて、スライドだけ組み直す（台本なら音声も）
  if (st && !一回.組んだ && 観.直すセル.length && 記.取り込み済み && (status === '完了' || status === 'エラー')) {
    // ★1回の実行で1件（組み直しは1回まで）。受けた順（日時の早いもの）から。同じ No を続けて直すと、
    //   あとの直しの「前」は先の直しの「後」なので、順を守らないと「違う」で止まる（2026-09-26）
    const x = 観.直すセル.slice().sort(function (p, q) {
      return String((p.中身 || {})['日時'] || '').localeCompare(String((q.中身 || {})['日時'] || ''));
    })[0];
    const o = x.中身 || {};
    const id = String(o['ID']);
    const r = 直すセルを書く_(ss, o['直す']);
    記.セル済 = 記.セル済 || {};
    if (r.違う.length) {
      記.セル済[id] = { 結果: '直せない', 文: r.違う.join('・') };
      動きを足す_(記, String(o['対象'] || '') + 'の直しを入れられませんでした：' + r.違う.join('・'));
      知らせを積む_(記, projectName, '直し（' + id + '）を構成シートへ入れられませんでした。' + r.違う.join('・'));
    } else {
      一回.組んだ = true;
      // ★台本（E列）が変われば音声も作り直す。スライドの文字・図解・タイトルの直しでも、Mac が台本を合わせたら E列が入る
      const 音声も = String(o['対象'] || '') === '台本' || r.書いた.some(function (w) { return w.列 === 5; });
      const ok = スライドを組む_(cfg, row, 音声も);
      if (ok === null) {
        一回.あと回し = true;   // ★【直すセル】は残す（次の回の最初に、書いたセルを当て直して組む）
      } else {
        記.セル済[id] = ok ? { 結果: '直した', 文: '' } : { 結果: '直せない', 文: 'セルは直しましたが、スライドを組めませんでした' };
        動きを足す_(記, '直し（' + r.書いた.length + 'か所）を構成シートへ入れ、スライドを組み直しました' + (音声も ? '（台本が変わったので音声も作り直します）' : '（音声はそのまま）'));
        // ★いまの図解・サムネイルで組んだので、(4) で同じものをもう一度組まない
        if (ok && 観.ファイル.図解済み) 記.組んだ図解 = 図の印_(観);
      }
    }
    // ★書けなかったときは【止まった】直すセル_ にする。ボットはこれを見て「直せませんでした」と返す（2026-09-27）
    if (!一回.あと回し) 済にする_(folder, x.実物, nfc_(x.名).replace('【直すセル】', r.違う.length ? '【止まった】直すセル_' : '【済】直すセル_'));
    // 直しの記録にまだ無い（【指示】直し_ を見ないうちに終わった）ときも、記録に載せる
    if (!(記.直し || []).some(function (d) { return d.ID === id; })) {
      記.直し = (記.直し || []).concat([{ ID: id, 受付: String(o['日時'] || ''), 対象: String(o['対象'] || ''), 依頼: '', 状態: '作り直し中', 頼んだ人: '', 元: 'PM相談室' }]);
    }
  }

  // (3) 直しの受付・直し終わりを記録へ
  記.直し = 直しをまとめる_(記.直し, 観.直し指示, 観.直し済, 日時の文字_(new Date()), 記.セル済);
  const 新しく受けた = 直しの差分_(JSON.parse(前の直し), 記.直し);
  新しく受けた.受けた.forEach(function (d) {
    動きを足す_(記, (d.対象 || '') + 'の直し（' + (d.元 || 'PM相談室') + (d.依頼 ? '：' + String(d.依頼).slice(0, 30) : '') + '）');
    if (記.承認 && 記.承認.済) {   // 承認のあとに直しが来たら、承認は外す（もう一度見てもらう）
      記.承認 = { 済: false, 人: '', 日時: '' };
      if (st) st.getRange('B2').setValue(false);
      動きを足す_(記, '直しが入ったので、承認を外しました');
    }
  });
  新しく受けた.終わった.forEach(function (d) {
    動きを足す_(記, d.対象 + 'の直し：' + (d.結果 === '直した' ? '直しました' : '直せませんでした（' + (d.文 || '') + '）'));
  });
  // ★スライド・台本の直しは (2b) の【直すセル】で、図解の直しは【図解済み】の日時が変わったのを見て (4) で組み直す

  // (4) スライド: 取り込みと図解がそろったら組む／図解の差し替え・直しなら音声を残して組み直す
  const 図 = 観.ファイル.図解済み;
  const 図の印 = 図の印_(観);
  // ★図の印の形を変えた（画像の更新を足した。2026-09-27）。前の形（項目3つ）で覚えている案件は、組み直さずに印だけ今の形へ移す
  if (記.組んだ図解 && 図の印 && String(記.組んだ図解).split('|').length === 3 && 図の印.split('|').length === 4 &&
      図の印.indexOf(String(記.組んだ図解)) === 0) 記.組んだ図解 = 図の印;
  if (st && !一回.組んだ) {
    if (status === 'レビュー待ち' && 図) {
      一回.組んだ = true;
      const ok = スライドを組む_(cfg, row, false);
      if (ok) { 記.組んだ図解 = 図の印; 記.途中で止まった回 = 0; 動きを足す_(記, 'スライドを組みました（音声を作り始めます）'); }
      else if (ok === false) 動きを足す_(記, 'スライドを組めませんでした');
    } else if ((status === '完了' || status === 'エラー') && 記.取り込み済み && 図 && 記.組んだ図解 && 記.組んだ図解 !== 図の印 && 記.組めなかった印 !== 図の印) {
      一回.組んだ = true;
      const ok2 = スライドを組む_(cfg, row, false);
      // ★組めなかったときは、同じ絵のままでは組み直さない（5分ごとに失敗し続けないように。絵が変われば、また組む）
      if (ok2 === false) { 記.組めなかった印 = 図の印; 動きを足す_(記, '画像の差し替えのあと、スライドを組めませんでした'); }
      if (ok2) {
        記.組んだ図解 = 図の印; 記.途中で止まった回 = 0;
        動きを足す_(記, '画像（図解・サムネイルなど）が差し替わったので、スライドだけ組み直しました（音声はそのまま）');
        if (記.承認 && 記.承認.済) {   // ★承認のあとにスライドが変わったら、もう一度見てもらう（PM相談室の直しと同じ）
          記.承認 = { 済: false, 人: '', 日時: '' };
          if (st) st.getRange('B2').setValue(false);
          動きを足す_(記, 'スライドが変わったので、承認を外しました');
        }
      }
    }
  }
  if (一回.組んだ) {   // 状態が変わったので、管理シートを読み直して見直す
    const reg = registry_(cfg);
    if (reg[hearingId]) {
      const 新 = mgmtSheet_(cfg).getRange(reg[hearingId], 1, 1, 流れ_列).getValues()[0];
      row = 新; status = String(新[2] || '');
      観 = 観測する_(cfg, row, folder);
      観.見張り = 一回.見張り || null;
    }
  }

  // (5) 承認: 操作パネルの【承認】◯◯社.json か B2（新の流れでは「確認して承認」）。全部できてからだけ受け付ける
  let 進 = 進みを計算_(観, 記, new Date());
  if (st && 観.承認ファイル) {
    const 頼み = 観.承認ファイル.中身 || {};
    const 人 = String(頼み['人'] || 頼み['メール'] || '操作パネル');
    if (進.押せる.承認) {
      記.承認 = { 済: true, 人: 人, 日時: 日時の文字_(new Date()) };
      st.getRange('B2').setValue(true);
      動きを足す_(記, '確認して承認しました（' + 人 + '）');
      st.getRange('B3').setValue('承認しました（' + 人 + '）。操作パネルの「納品フォルダへ入れる」で納品できます');
    } else if (!(記.承認 && 記.承認.済)) {
      動きを足す_(記, '承認できませんでした（' + 人 + '）：' + 進.いまここ.文);
    }
    済にする_(folder, 観.承認ファイル.実物, '【済】承認_' + company + '.json');
    観.承認ファイル = null;
  }
  if (st) {
    const b2 = st.getRange('B2').getValue() === true;
    if (b2 && !(記.承認 && 記.承認.済)) {
      if (進.押せる.承認) {
        記.承認 = { 済: true, 人: '構成シートのB2', 日時: 日時の文字_(new Date()) };
        動きを足す_(記, '確認して承認しました（構成シートのB2）');
        st.getRange('B3').setValue('承認しました。操作パネルの「納品フォルダへ入れる」で納品できます');
      } else {
        st.getRange('B2').setValue(false);
        st.getRange('B3').setValue('承認はまだできません。' + 進.いまここ.文 + '。全部できてから押してください');
      }
    } else if (!b2 && 記.承認 && 記.承認.済) {
      記.承認 = { 済: false, 人: '', 日時: '' };
      動きを足す_(記, '承認が外されました（構成シートのB2）');
    }
  }

  // (6) 納品: 操作パネルの【納品指示】◯◯社.json。承認のあとだけ（B22 は processDelivery_ が同じ決まりで受ける）
  if (st && 観.納品ファイル) {
    const 頼み = 観.納品ファイル.中身 || {};
    const 人 = String(頼み['人'] || 頼み['メール'] || '操作パネル');
    if (新の流れの承認済みか_(hearingId, st, 記)) {
      const r = 納品する_(cfg, projectName, st);   // 70_delivery.js（B22 と同じ中身）
      if (r.入れた) { 記.納品 = { 済: true, 日時: 日時の文字_(new Date()) }; 動きを足す_(記, '納品フォルダへ入れました（' + 人 + '）'); }
      else 動きを足す_(記, '納品フォルダへ入れられませんでした：' + String(r.文).slice(0, 80));
    } else {
      動きを足す_(記, '承認の前なので、納品フォルダへは入れませんでした（' + 人 + '）');
    }
    済にする_(folder, 観.納品ファイル.実物, '【済】納品指示_' + company + '.json');
    観.納品ファイル = null;
  }

  // (7) 進みを書き、止まったものを知らせる
  進 = 進みを計算_(観, 記, new Date());
  if (進.全部できた && !記.全部できた) { 記.全部できた = true; 動きを足す_(記, '全部できた'); 進 = 進みを計算_(観, 記, new Date()); }
  止まったを知らせる_(projectName, 進, 記);
  進.知らせ = (記.知らせ || []).slice();   // いま積んだ知らせも 進み.json に出す
  進.動き = (記.動き || []).slice(-動き_最大);
  進み_記録を書く_(hearingId, 記);
  進みを書く_(folder, 進);
}

/** 【直すセル】を構成シートの4タブへ書く。「前」と今のセルが1つでも違えば何も書かない */
function 直すセルを書く_(ss, 直す) {
  const タブたち = {};
  [TABS.INTRO, TABS.PROBLEM, TABS.STORY, TABS.OFFER].forEach(function (t) {
    const sh = ss.getSheetByName(t);
    if (sh) タブたち[t] = sh.getDataRange().getValues();
  });
  const r = 直すセルを当てる_(タブたち, 直す);
  r.書く.forEach(function (w) { ss.getSheetByName(w.タブ).getRange(w.行, w.列).setValue(w.後); });
  return { 書いた: r.書く, 違う: r.違う };
}

/** Mac の見張りが止まったら、止まった1回につき1通だけ社内（宮津さん）へ置く */
function 見張りが止まったら知らせる_(見張り) {
  if (!見張り || !見張り.止まった) return;
  const P = PropertiesService.getScriptProperties();
  if (P.getProperty('MIHARI_TOMATTA') === 見張り.最後) return;
  P.setProperty('MIHARI_TOMATTA', 見張り.最後);
  try {
    const いま = new Date();
    const o = { 宛先: '社内', 文: '宮津さんへ：Mac の資料生成の見張りが ' + 見張り.分 + '分止まっています（最後に動いたのは ' + 見張り.最後 + '）。新しい流れの案件の生成が進みません', 案件名: '', 置いた日時: 日時の文字_(いま) };
    知らせ待ちフォルダ_().createFile('【知らせ】見張り_' + 日時の文字_(いま).slice(0, 19).replace(/[-:]/g, '').replace('T', '-') + '.json', JSON.stringify(o), 'application/json');
  } catch (e) { Logger.log('★見張りの止まりを知らせられません: ' + e); }
}

/** パネルから来たファイルを「済」の名前に変える（同じ名前の古い【済】はゴミ箱へ） */
function 済にする_(folder, f, 新しい名前) {
  if (!f) return;
  try { const old = filesByName_(folder, 新しい名前); while (old.hasNext()) old.next().setTrashed(true); }
  catch (e) { /* 古い【済】が残っても困らない */ }
  try { f.setName(新しい名前); } catch (e) { try { f.setTrashed(true); } catch (e2) { /* 次の回にもう一度 */ } }
}

/**
 * 新の流れの案件が承認済みか。★B2 だけでは見ない（B2 は誰でも押せ、全部できる前に押されたものは
 * processNewFlow_ があとで外すため）。覚え（承認.済）と B2 の両方がそろっているときだけ承認済み。
 */
function 新の流れの承認済みか_(hearingId, st, 記) {
  const r = 記 || 進み_記録を読む_(String(hearingId));
  return !!(r.承認 && r.承認.済) && st.getRange('B2').getValue() === true;
}

/**
 * スライドを組む（組み直す）。★音声は消さない（音声も=true のときだけ消して作り直させる）。
 * 旧の流れの B16（processSlideRegen_）は音声を必ず消すので、それとは分けてある。
 * generateSlides_ は台本txtを書き直すが、音声mp3があるかぎり見張りは音声を作り直さない。
 */
// ★この実行が始まった時刻（ファイルが読まれた時＝実行の始まり）。スライドを組むのに5分ほどかかり、
//   実行の上限（6分）を越えると途中で止まるので、始まってから時間がたっていたら次の回の最初に回す（2026-09-26）
const 新_実行の始まり_ = Date.now();
const 新_組む待ちのキー = '新_組む待ち';
const 新_組む前の余裕_ms = 45 * 1000;
// ★main が「先に組む」回として呼んだときは true（時間の確かめをしない。しないと、ずっと先送りになり旧の流れが回らない）
var 新_先に組む回_ = false;

/** ★図解とサムネイル（スライドの1枚目）の印。変わったらスライドを組み直す（2026-09-26）。印にサムネイルの更新日時も入れる */
function 図の印_(観) {
  const 図 = 観.ファイル.図解済み, サ = 観.ファイル.サムネイル;
  // ★素材フォルダの画像を手で差し替えたとき（ChatGPT での手直し）も変わるように、画像のいちばん新しい更新日時も入れる
  return 図 ? String(図.日時 || '') + '|' + String(図.更新 || '') + '|' + String((サ && サ.日時) || '') + '|' + String(観.ファイル.画像の更新 || '') : '';
}

/** @return true=組めた / false=失敗 / null=時間が足りないので次の回へ回した */
function スライドを組む_(cfg, row, 音声も) {
  if (!新_先に組む回_ && Date.now() - 新_実行の始まり_ > 新_組む前の余裕_ms) {
    PropertiesService.getScriptProperties().setProperty(新_組む待ちのキー, String(row[1] || '1'));
    Logger.log('スライドを組むのは次の回の最初に回します（この回は時間が足りない）: ' + row[1]);
    return null;
  }
  const hearingId = String(row[0]), projectName = String(row[1]), kouseiUrl = String(row[3]);
  const kouseiId = idFromUrl_(kouseiUrl);
  upsertRow_(cfg, hearingId, projectName, 'スライド生成中', kouseiUrl, '', 音声も ? '組み直し中（音声も作り直す）' : '組み直し中（音声は残す）');
  setKouseiStatus_(kouseiId, 'スライドを組んでいます');
  try {
    const result = generateSlides_(kouseiId, { 安全: true });   // 50_slides.js（組み終わってから旧と入れ替える）
    // ★音声は、新しい【台本】txt を書いたあとで消す（2026-09-26）。先に消していたので、Mac が古い台本で音声を作り直していた
    if (音声も) {
      try {
        const folder = caseFolderIfExists_(cfg, projectName);
        if (folder) {
          const mp3 = filesByName_(folder, '【音声】' + caseCompany_(projectName) + '.mp3');
          while (mp3.hasNext()) mp3.next().setTrashed(true);
        }
      } catch (e) { /* 音声が消せなくても組み直しは済んでいる */ }
    }
    upsertRow_(cfg, hearingId, projectName, '完了', kouseiUrl, result.pres.getUrl(), result.warnings.join(' / '));
    setKouseiStatus_(kouseiId, 'スライドができました: ' + result.pres.getUrl());
    return true;
  } catch (e) {
    upsertRow_(cfg, hearingId, projectName, 'エラー', kouseiUrl, '', String(e));
    setKouseiStatus_(kouseiId, 失敗の文_('スライドを組む', e, '図解を差し替えるか、宮津さんへ相談してください'));
    return false;   // 知らせは、止まった行として 止まったを知らせる_ が積む
  }
}

/**
 * ★手で回す用: 構成シートのURLかヒアリングシートIDを入れて実行すると、音声を残してスライドだけ組み直す。
 *   図解を手で差し替えたとき・セルを手で直したときに使う（旧の流れの B16 と違い、音声は消さない）。
 */
function スライドだけ組み直す() {
  const 指定 = 'ここに構成シートのURLかヒアリングシートIDを入れる';
  const cfg = getConfig_();
  const rows = mgmtSheet_(cfg).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === 指定 || String(rows[i][3]) === 指定) {
      Logger.log(スライドを組む_(cfg, rows[i], false) ? '組み直しました（音声はそのまま）' : '組み直せませんでした（管理シートF列を見てください）');
      return;
    }
  }
  Logger.log('管理シートに見つかりません: ' + 指定);
}

// ============================================================
// 構成シート（設定タブ）の新の流れ用の体裁
// ============================================================
const 新_A2 = '確認して承認（全部できてから。操作パネルからも押せます）';
const パネルのリンクのセル = 'C1';   // ★設定タブで使っていない所（A・B列だけを使っている）

/** 操作パネル（別のGASプロジェクト ~/webinar-panel-gas のウェブアプリ）のURL。スクリプトプロパティ PANEL_URL */
function パネルのURL_() {
  try { return PropertiesService.getScriptProperties().getProperty('PANEL_URL') || ''; } catch (e) { return ''; }
}

/** 「操作パネルを開く」の式。?f= は案件フォルダ（【出力】生成物 → ◯◯社）のID */
function パネルのリンクの式_(url, folderId) {
  return '=HYPERLINK("' + String(url).replace(/"/g, '') + '?f=' + encodeURIComponent(folderId) + '","▶ 操作パネルを開く（別のタブ）")';
}

/** 新の流れの案件だけ: A2 の見出しを「確認して承認」に、C1 に「操作パネルを開く」リンク */
function 新の流れの設定タブを整える_(st, folderId) {
  if (!st) return;
  try {
    if (String(st.getRange('A2').getValue()) !== 新_A2) st.getRange('A2').setValue(新_A2);
    const url = パネルのURL_();
    if (!url || !folderId) return;   // 操作パネルを公開して PANEL_URL を入れるまでは出さない（入れた次の回に出る）
    const 式 = パネルのリンクの式_(url, folderId);
    const c = st.getRange(パネルのリンクのセル);
    if (String(c.getFormula()) !== 式) c.setFormula(式).setFontWeight('bold').setFontColor('#1a73e8');
    if (st.getColumnWidth(3) < 240) st.setColumnWidth(3, 240);   // ★リンクの文字が切れないように（2026-09-26）
    // ★新の流れでは、4行目から下の旧のボタン（LPを直す・スライドの不具合・タイトル・サムネ・納品のチェック）は使わない。
    //   直しは PM相談室、進みと承認・納品は操作パネル。残すと迷うので隠す（消さない。2026-09-26）
    if (st.getMaxRows() > 3 && !st.isRowHiddenByUser(4)) st.hideRows(4, st.getMaxRows() - 3);
    const b3 = st.getRange('B3');
    if (String(b3.getValue()).trim() === 'AI出力待ち') b3.setValue('資料を作っています。進みは右の「操作パネルを開く」で見られます');
  } catch (e) { /* 体裁の失敗で止めない */ }
}

// ============================================================
// 知らせ（★テスト案件は Google Chat の「テスト」スペースへ。本物の社内チャットには出さない）
// ============================================================
/**
 * テスト用の案件か。★operations の 定型連絡.gs 共通_テスト案件か_ と同じ決め方（案件名だけで見る版）。
 *   あちらを直したら、ここもそろえる。
 */
const 共通_テストの目印_ = ['テスト', 'リハーサル'];
const 共通_テストの案件名_ = ['4章_畠', '3-2_福田', '2-4_桑野', '稼働A', '稼働B'];
function テスト案件か_(案件名) {
  const 名 = caseCompany_(String(案件名 || '').trim());
  if (!名) return false;
  if (共通_テストの案件名_.indexOf(名) >= 0) return true;
  return 共通_テストの目印_.some(function (m) { return 名.indexOf(m) >= 0; });
}

/**
 * 新の流れの知らせ。★ここからは外へ送らない（このスクリプトは外へ送る権限を持たず、Webhook も持たない）。
 * 共有ドライブ【固定】ウェビナー自動生成/【知らせ待ち】/ に1通1ファイル置くだけ。送るのは Mac（形は設計書 10.）。
 *   宛先: テスト案件 → 「テスト」（Google Chat「テスト_Peatix催促」）、それ以外 → 「社内」
 * 同じものは 進み.json の「知らせ」にも残す（パネルで見られる）。
 */
function 知らせを積む_(記, projectName, 文) {
  記.知らせ = 記.知らせ || [];
  記.知らせの番号 = (Number(記.知らせの番号) || 0) + 1;
  const いま = new Date();
  const 会社 = caseCompany_(projectName);
  const テスト = テスト案件か_(projectName);
  記.知らせ.push({ ID: String(記.知らせの番号), 日時: 日時の文字_(いま), 文: 会社 + ' … ' + 文, テスト: テスト });
  if (記.知らせ.length > 知らせ_最大) 記.知らせ = 記.知らせ.slice(-知らせ_最大);
  Logger.log('（知らせ）' + 会社 + ' … ' + 文);
  try {
    const 中身 = 知らせの中身_(会社, 文, テスト, いま);
    知らせ待ちフォルダ_().createFile(知らせのファイル名_(会社, いま, 記.知らせの番号), JSON.stringify(中身), 'application/json');
  } catch (e) { Logger.log('★【知らせ待ち】へ置けません: ' + e); }
}

/** 【知らせ待ち】の1通（★Driveを触らない・単体テストあり） */
function 知らせの中身_(会社, 文, テスト, いま) {
  return { 宛先: テスト ? 'テスト' : '社内', 文: 会社 + ' … ' + 文, 案件名: 会社, 置いた日時: 日時の文字_(いま) };
}
function 知らせのファイル名_(会社, いま, 番号) {
  return '【知らせ】' + 会社 + '_' + 日時の文字_(いま).slice(0, 19).replace(/[-:]/g, '').replace('T', '-') + '_' + 番号 + '.json';
}

/** 【固定】ウェビナー自動生成/【知らせ待ち】（【出力】生成物の親の中。無ければ作る） */
const 知らせ待ち_名前 = '【知らせ待ち】';
function 知らせ待ちフォルダ_() {
  const 固定 = DriveApp.getFolderById(getConfig_().OUTPUT_FOLDER_ID).getParents().next();
  const it = foldersByName_(固定, 知らせ待ち_名前);
  return it.hasNext() ? it.next() : 固定.createFolder(知らせ待ち_名前);
}
const 知らせ_最大 = 20;

/** 止まった行を、同じ内容では1回だけ知らせる */
function 止まったを知らせる_(projectName, 進, 記) {
  記.知らせた = 記.知らせた || {};
  const いま = {};
  ['LP', '資料'].forEach(function (k) {
    (進.系統[k] || []).forEach(function (it) {
      if (it.状態 !== '止まった') return;
      const 鍵 = it.名 + '｜' + String(it.文 || '').slice(0, 60);
      いま[鍵] = true;
      if (記.知らせた[鍵]) return;
      知らせを積む_(記, projectName, '「' + it.名 + '」が止まりました。' + (it.文 || ''));
      動きを足す_(記, '「' + it.名 + '」が止まりました');
    });
  });
  記.知らせた = いま;   // 直ったものは忘れる（また止まったら知らせる）
}
