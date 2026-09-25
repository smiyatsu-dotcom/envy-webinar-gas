/**
 * 17_cleanup.gs（テスト案件の片付け・撮り直し用・2026-09-26）
 *
 * 動画の撮り直しのたびに、テスト案件（例：動画撮影テスト社）の webinar-gas 側を消す。
 * operations 側（台帳・ヒアリングシート・【納品】・【素材】のフォルダなど）は、案件の下地側（資料の流れ.gs）が消す。
 *
 * ★安全装置: 案件名が テスト案件か_（名前に「テスト」「リハーサル」など。15_new_flow.js）のときだけ動く。
 *   それ以外の名前では何もせずに止まる。本番の案件（ホスピタソン社など）は、どう呼んでも消さない。
 * ★消すのではなく「ゴミ箱へ移す」（30日以内なら戻せる）。管理シートの行だけは行を消す。
 * ★先に 下見（何を消すかの一覧だけ）を必ず見る。
 *
 * 片付けるもの（webinar-gas が作った・読むもの）
 *   ・管理シートの行（その案件名の行すべて）
 *   ・構成シート
 *   ・【出力】生成物 → ◯◯社 の中のファイルすべて（プロンプトDoc・【指示】・【作業中】・【状態】進み.json・
 *     【LP】・【スライド】・【台本】・【音声】・【URL管理】・【済】…）。★フォルダそのものは残す（案件の下地のもの）
 *   ・【素材】投影資料 → ◯◯社_ヒアリングシート の【図解済み】.json（図解の画像は残す。撮り直しで使う）
 *   ・スクリプトプロパティ（SHINKO_・START_FAIL_・AUTOIMPORT_）
 */

/** ★手で実行するときは、ここに案件名（◯◯社）を入れる */
const 片付ける案件名 = '動画撮影テスト社';

function テスト案件の片付けを下見する() { Logger.log(テスト案件を片付ける_(片付ける案件名, true).join('\n')); }
function テスト案件を片付ける() { Logger.log(テスト案件を片付ける_(片付ける案件名, false).join('\n')); }

/**
 * 下見=true なら一覧を返すだけで何も変えない。戻り値は行の配列（ログ・画面に出す文）。
 */
function テスト案件を片付ける_(案件名, 下見) {
  const company = caseCompany_(String(案件名 || '').trim());
  if (!company) throw new Error('案件名が空です');
  if (!テスト案件か_(company)) {
    throw new Error('「' + company + '」はテスト案件の名前ではないので、片付けません（名前に「テスト」「リハーサル」が入っている案件だけ）');
  }
  const cfg = getConfig_();
  const P = PropertiesService.getScriptProperties();
  const 出 = [(下見 ? '［下見］' : '［片付け］') + company + ' の webinar-gas 側'];
  const する = function (文, fn) { 出.push((下見 ? '・消す予定：' : '・消した：') + 文); if (!下見) { try { fn(); } catch (e) { 出.push('　★できませんでした：' + e); } } };

  // 管理シートの行（下から消す。上から消すと行番号がずれる）
  const sh = mgmtSheet_(cfg);
  const rows = sh.getDataRange().getValues();
  const 行たち = [];
  for (let i = 1; i < rows.length; i++) if (caseCompany_(String(rows[i][1] || '')) === company) 行たち.push({ 行: i + 1, v: rows[i] });
  行たち.slice().reverse().forEach(function (r) {
    const hid = String(r.v[0] || '');
    const kurl = String(r.v[3] || '');
    if (kurl) {
      let kid = '';
      try { kid = idFromUrl_(kurl); } catch (e) { kid = ''; }
      if (kid) {
        する('構成シート ' + kurl, function () { DriveApp.getFileById(kid).setTrashed(true); });
        if (P.getProperty('AUTOIMPORT_' + kid) !== null) する('プロパティ AUTOIMPORT_' + kid, function () { P.deleteProperty('AUTOIMPORT_' + kid); });
      }
    }
    ['SHINKO_', 'START_FAIL_'].forEach(function (k) {
      if (hid && P.getProperty(k + hid) !== null) する('プロパティ ' + k + hid, function () { P.deleteProperty(k + hid); });
    });
    する('管理シートの' + r.行 + '行目（' + r.v[1] + '・' + r.v[2] + '）', function () { sh.deleteRow(r.行); });
  });
  if (!行たち.length) 出.push('・管理シートに行はありません');

  // 【出力】生成物 → ◯◯社 の中身（フォルダは残す）
  const folder = caseFolderIfExists_(cfg, company);
  if (folder) {
    const it = folder.getFiles();
    let n = 0;
    while (it.hasNext()) { const f = it.next(); n++; する('【出力】生成物/' + company + '/' + f.getName(), function () { f.setTrashed(true); }); }
    if (!n) 出.push('・【出力】生成物/' + company + ' は空です');
  } else 出.push('・【出力】生成物/' + company + ' はありません');

  // 【素材】の【図解済み】.json
  try {
    const sub = foldersByName_(DriveApp.getFolderById(cfg.ASSETS_FOLDER_ID), 新の流れの案件名_(company));
    if (sub.hasNext()) {
      const z = filesByName_(sub.next(), 図解済み_ファイル名);
      while (z.hasNext()) { const f = z.next(); する('【素材】投影資料/' + 新の流れの案件名_(company) + '/' + 図解済み_ファイル名, function () { f.setTrashed(true); }); }
    }
  } catch (e) { 出.push('　★素材フォルダを見られませんでした：' + e); }

  return 出;
}
