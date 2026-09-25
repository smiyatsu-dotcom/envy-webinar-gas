/**
 * 16_progress.gs（進みの記録・2026-09-26）
 * 【出力】生成物 → ◯◯社 / 【状態】進み.json を書く。★書くのはこのスクリプトだけ。
 * Mac とボットは【作業中】【図解済み】【指示】直し_【済】直し_ のファイルを、操作パネルは【承認】【納品指示】を
 * 置くだけ（形は設計書 10.）。パネル・案件ボード・サイトは、同じ 進み.json を読む。形の正本は設計書 3.。
 *
 * つくり:
 *   観測する_（Driveを読む）→ 進みを計算_（Driveを触らない・単体テストあり）→ 進みを書く_
 *   案件ごとの覚え（承認・納品・直し・動き）はスクリプトプロパティ SHINKO_<ヒアリングシートID>
 */

const 進み_ファイル名 = '【状態】進み.json';
const 図解済み_ファイル名 = '【図解済み】.json';
const 心拍_切れ分 = 20;   // 【作業中】の心拍がこれより古いと「止まった」
const 動き_最大 = 30;
const 直し_最大 = 20;

/** 作るものの順番と、ふだんかかる分（目安。【作業中】の見込み分があればそちらを使う） */
const 進み_系統 = {
  LP: ['タイトル', 'サムネイル', 'LP'],
  資料: ['構成・台本', '図解・挿絵', 'スライド', '音声']
};
const 進み_目安分 = { 'タイトル': 12, 'サムネイル': 10, 'LP': 2, '構成・台本': 30, '図解・挿絵': 25, 'スライド': 5, '音声': 15 };
/** 【作業中】の「仕事」→ パネルの行 */
const 仕事の行 = { 'LP': ['タイトル', 'LP'], 'サムネ': ['サムネイル'], '資料': ['構成・台本'], '図解': ['図解・挿絵'], '音声': ['音声'] };
/** 直しの「対象」→ パネルの行 */
const 直しの行 = { 'LP': 'LP', 'タイトル': 'タイトル', 'サムネイル': 'サムネイル', 'スライド': 'スライド', '台本': '構成・台本', '図解': '図解・挿絵' };

// ============================================================
// 日時
// ============================================================
/** 2026-09-26T10:15:00+09:00 の形（日本時間） */
function 日時の文字_(d) {
  const t = new Date((d instanceof Date ? d.getTime() : Number(d)) + 9 * 3600000);
  return t.toISOString().slice(0, 19) + '+09:00';
}
function 分たった_(文字, いま) {
  const t = Date.parse(String(文字 || ''));
  if (isNaN(t)) return null;
  return Math.floor(((いま instanceof Date ? いま.getTime() : いま) - t) / 60000);
}

// ============================================================
// 覚え（スクリプトプロパティ）
// ============================================================
function 進み_記録の既定_() {
  return { 承認: { 済: false, 人: '', 日時: '' }, 納品: { 済: false, 日時: '' }, 直し: [], 動き: [], 組んだ図解: '', 取り込み済み: false, 全部できた: false, 知らせた: {}, 知らせ: [], 知らせの番号: 0 };
}
function 進み_記録を読む_(hearingId) {
  const 既 = 進み_記録の既定_();
  try {
    const s = PropertiesService.getScriptProperties().getProperty('SHINKO_' + hearingId);
    if (!s) return 既;
    const o = JSON.parse(s);
    Object.keys(既).forEach(function (k) { if (o[k] === undefined) o[k] = 既[k]; });
    return o;
  } catch (e) { return 既; }
}
function 進み_記録を書く_(hearingId, 記) {
  記.動き = (記.動き || []).slice(-動き_最大);
  記.直し = (記.直し || []).slice(-直し_最大);
  PropertiesService.getScriptProperties().setProperty('SHINKO_' + hearingId, JSON.stringify(記));
}
function 進み_記録を変える_(hearingId, fn) {
  const 記 = 進み_記録を読む_(hearingId);
  fn(記);
  進み_記録を書く_(hearingId, 記);
  return 記;
}
function 動きを足す_(記, 文, いま) {
  記.動き = 記.動き || [];
  記.動き.push({ 日時: 日時の文字_(いま || new Date()), 文: 文 });
  if (記.動き.length > 動き_最大) 記.動き = 記.動き.slice(-動き_最大);
}

// ============================================================
// 観測（Driveを読む）
// ============================================================
function JSONを読む_(f) {
  try { return JSON.parse(f.getBlob().getDataAsString('UTF-8')); } catch (e) { return null; }   // 書きかけは次の回
}

/**
 * 案件フォルダのファイル名を仕分ける。★Driveを触らない（単体テストあり）。
 * files: [{名, mime, url, 日時, 読む()}]（読む はJSONの中身を返す。読めなければ null）
 */
function ファイルを仕分ける_(company, files) {
  const 出 = {
    ファイル: { タイトル案: null, LP: null, LP待ち: false, 指示LP: false, 指示資料: false, 指示サムネ: false, 保留サムネ: false, 要確認サムネ: false, 台本txt: null, 音声: null, サムネイル: null, 図解済み: null },
    作業中: [], 直し指示: [], 直し済: [], 保留サムネの名前: '',
    承認ファイル: null, 納品ファイル: null   // 操作パネルが置く頼み（{名, 中身}）
  };
  const F = 出.ファイル;
  const 頭 = function (n, s) { return n.indexOf(s) === 0; };
  files.forEach(function (f) {
    const n = nfc_(f.名);
    const 印 = { url: f.url || '', 日時: f.日時 || '' };
    if (n === nfc_('【タイトル案】' + company + '.txt')) F.タイトル案 = 印;
    else if (頭(n, '【LP】') && f.mime === 'application/vnd.google-apps.document') F.LP = 印;
    else if (頭(n, '【LP】') && /\.(txt|md)$/i.test(n)) F.LP待ち = true;
    else if (頭(n, '【指示】LP_')) F.指示LP = true;
    else if (頭(n, '【指示】資料_')) F.指示資料 = true;
    else if (頭(n, '【指示】サムネ_')) F.指示サムネ = true;
    else if (頭(n, '【保留】サムネ指示_')) { F.保留サムネ = true; 出.保留サムネの名前 = n; }
    else if (頭(n, '【要確認】サムネ指示_')) F.要確認サムネ = true;
    else if (頭(n, '【台本】') && /\.txt$/i.test(n)) F.台本txt = 印;
    else if (頭(n, '【音声】') && /\.mp3$/i.test(n)) F.音声 = 印;
    else if (頭(n, '【サムネイル】')) F.サムネイル = 印;
    else if (頭(n, '【作業中】') && /\.json$/i.test(n)) { const o = f.読む(); if (o && o['仕事']) 出.作業中.push(o); }
    else if (頭(n, '【指示】直し_') && /\.json$/i.test(n)) { const o = f.読む(); if (o && o['ID']) 出.直し指示.push(o); }
    else if (頭(n, '【済】直し_') && /\.json$/i.test(n)) { const o = f.読む(); if (o && o['ID']) 出.直し済.push(o); }
    else if (n === nfc_('【承認】' + company + '.json')) 出.承認ファイル = { 名: n, 中身: f.読む() || {} };
    else if (n === nfc_('【納品指示】' + company + '.json')) 出.納品ファイル = { 名: n, 中身: f.読む() || {} };
  });
  return 出;
}

/** 管理シートの1行と案件フォルダ・素材フォルダを読んで、進みを計算する材料を集める */
function 観測する_(cfg, row, folder) {
  const projectName = String(row[1]);
  const company = caseCompany_(projectName);
  const list = [];
  const 実物 = {};
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const 名 = f.getName();
    実物[nfc_(名)] = f;
    list.push({
      名: 名, mime: f.getMimeType(), url: f.getUrl(),
      日時: (function () { try { return 日時の文字_(f.getLastUpdated()); } catch (e) { return ''; } })(),
      読む: function () { return JSONを読む_(f); }
    });
  }
  const 観 = ファイルを仕分ける_(company, list);
  観.保留サムネ = 観.保留サムネの名前 ? 実物[観.保留サムネの名前] : null;
  if (観.承認ファイル) 観.承認ファイル.実物 = 実物[観.承認ファイル.名];
  if (観.納品ファイル) 観.納品ファイル.実物 = 実物[観.納品ファイル.名];

  // 素材フォルダ（【素材】投影資料 → ◯◯社_ヒアリングシート）: 図解済み・サムネイル
  try {
    const sub = foldersByName_(DriveApp.getFolderById(cfg.ASSETS_FOLDER_ID), projectName);
    if (sub.hasNext()) {
      const af = sub.next();
      const z = filesByName_(af, 図解済み_ファイル名);
      if (z.hasNext()) {
        const zf = z.next();
        const o = JSONを読む_(zf);
        if (o) { o.更新 = 日時の文字_(zf.getLastUpdated()); 観.ファイル.図解済み = o; }
      }
      if (!観.ファイル.サムネイル) {
        const t = filesByName_(af, 'サムネイル.png');
        if (t.hasNext()) { const tf = t.next(); 観.ファイル.サムネイル = { url: tf.getUrl(), 日時: 日時の文字_(tf.getLastUpdated()) }; }
      }
    }
  } catch (e) { /* 素材フォルダが読めなくても進みは出す */ }

  観.案件名 = company;
  観.状態 = String(row[2] || '');
  観.メッセージ = String(row[5] || '');
  観.スライドURL = String(row[4] || '');
  観.構成シートURL = String(row[3] || '');
  return 観;
}

// ============================================================
// 直しの記録（★Driveを触らない・単体テストあり）
// ============================================================
/**
 * これまでの直しの記録に、いまある【指示】直し_（受付中）と【済】直し_（終わり）を重ねる。
 * ボットが【済】を消しても、ここに残る。
 */
function 直しをまとめる_(前, 指示たち, 済たち, いまの文字) {
  const 出 = (前 || []).map(function (d) { return Object.assign({}, d); });
  const 引く = {};
  出.forEach(function (d) { 引く[d.ID] = d; });
  (指示たち || []).forEach(function (o) {
    const id = String(o['ID']);
    if (引く[id]) return;
    const d = {
      ID: id, 受付: String(o['受付'] || いまの文字 || ''), 対象: String(o['対象'] || ''), 依頼: String(o['依頼'] || ''),
      状態: '作り直し中', 頼んだ人: String(o['頼んだ人'] || ''), 元: String(o['元'] || 'PM相談室')
    };
    出.push(d); 引く[id] = d;
  });
  (済たち || []).forEach(function (o) {
    const id = String(o['ID']);
    let d = 引く[id];
    if (!d) { d = { ID: id, 受付: '', 対象: '', 依頼: '', 頼んだ人: '', 元: 'PM相談室' }; 出.push(d); 引く[id] = d; }
    if (d.状態 === '済' || d.状態 === '止まった') return;
    const 直した = String(o['結果'] || '') === '直した';
    d.状態 = 直した ? '済' : '止まった';
    d.結果 = 直した ? '直した' : '直せない';
    d.文 = String(o['文'] || '');
    d.前 = String(o['前'] || '');
    d.後 = String(o['後'] || '');
    d.終わり = いまの文字 || '';
  });
  return 出;
}

/** 前と後の直しの記録を比べ、新しく受けたもの・新しく終わったものを返す */
function 直しの差分_(前, 後) {
  const 前の = {};
  (前 || []).forEach(function (d) { 前の[d.ID] = d; });
  const 受けた = [], 終わった = [];
  (後 || []).forEach(function (d) {
    const p = 前の[d.ID];
    if (!p && d.状態 === '作り直し中') 受けた.push(d);
    if ((d.状態 === '済' || d.状態 === '止まった') && (!p || (p.状態 !== '済' && p.状態 !== '止まった'))) {
      if (!p && d.状態) 受けた.push(d);   // 受付を見ないうちに終わっていた直しも、受けたことにする
      終わった.push(d);
    }
  });
  return { 受けた: 受けた, 終わった: 終わった };
}

// ============================================================
// 進みの計算（★Driveを触らない・単体テストあり）
// ============================================================
/**
 * 観: 観測する_ の戻り値（ファイル・作業中・状態…）
 * 記: 覚え（承認・納品・直し・動き・取り込み済み）
 * 戻り値は設計書 3. の形に「押せる」「全部できた」を足したもの
 */
function 進みを計算_(観, 記, いま) {
  いま = いま || new Date();
  const F = 観.ファイル || {};
  const st = String(観.状態 || '');
  const 取り込み済み = !!記.取り込み済み || st === 'レビュー待ち' || st === 'スライド生成中' || st === '完了';
  const 作 = {};   // 行の名 → 【作業中】
  const 止まった作業 = {};
  (観.作業中 || []).forEach(function (w) {
    const 仕事 = String(w['仕事'] || '');
    let 行たち = 仕事の行[仕事] || [];
    if (仕事 === '直し') {
      const d = (記.直し || []).filter(function (x) { return x.ID === String(w['ID'] || ''); })[0];
      行たち = d && 直しの行[d.対象] ? [直しの行[d.対象]] : [];
    }
    const 心拍 = 分たった_(w['心拍'] || w['始'], いま);
    行たち.forEach(function (名) {
      作[名] = w;
      if (心拍 !== null && 心拍 > 心拍_切れ分) 止まった作業[名] = 'Macからの知らせが' + 心拍 + '分途絶えています（' + 仕事 + '）';
    });
  });
  const 残り分 = function (名) {
    const w = 作[名];
    if (w && w['見込み分'] !== undefined) {
      const 経 = 分たった_(w['始'], いま);
      return Math.max(1, Number(w['見込み分']) - (経 === null ? 0 : 経));
    }
    return 進み_目安分[名] || 5;
  };

  const 行 = function (名, 状態, 足す) {
    const o = { 名: 名, 状態: 状態 };
    Object.keys(足す || {}).forEach(function (k) { if (足す[k] !== undefined && 足す[k] !== '') o[k] = 足す[k]; });
    if (状態 === '作成中' || 状態 === '作り直し中') o.あと分 = 残り分(名);
    return o;
  };
  const 済か = function (印) { return 印 ? { 日時: 印.日時, url: 印.url } : {}; };

  // ---- LPの系統 ----
  const lp作成中 = F.指示LP || !!作['LP'];
  const タイトル = F.タイトル案 ? 行('タイトル', '済', 済か(F.タイトル案)) : 行('タイトル', lp作成中 ? '作成中' : '待ち');
  let サムネ;
  if (F.サムネイル) サムネ = 行('サムネイル', '済', 済か(F.サムネイル));
  else if (F.要確認サムネ) サムネ = 行('サムネイル', '止まった', { 文: 'サムネイルの参考画像が見つかりません（【要確認】サムネ指示）' });
  else if (F.指示サムネ || F.保留サムネ || 作['サムネイル']) サムネ = 行('サムネイル', '作成中');
  else サムネ = 行('サムネイル', '待ち');
  const LP = F.LP ? 行('LP', '済', 済か(F.LP)) : 行('LP', (lp作成中 || F.LP待ち) ? '作成中' : '待ち');

  // ---- 資料の系統 ----
  let 構成;
  if (取り込み済み) 構成 = 行('構成・台本', '済');
  else if (st === 'AI出力待ち' && /取り込み失敗/.test(観.メッセージ || '')) 構成 = 行('構成・台本', '止まった', { 文: String(観.メッセージ).slice(0, 200) });
  else if (st === 'エラー') 構成 = 行('構成・台本', '止まった', { 文: String(観.メッセージ || '').slice(0, 200) });
  else if (st === 'AI出力待ち' || F.指示資料 || 作['構成・台本']) 構成 = 行('構成・台本', '作成中');
  else 構成 = 行('構成・台本', '待ち');

  const 図 = F.図解済み;
  let 図解;
  if (図) {
    const 止 = (図['止まった'] || []).filter(function (x) { return x; });
    図解 = 行('図解・挿絵', '済', { 日時: String(図['日時'] || 図.更新 || ''), 枚: 図['枚数'] !== undefined ? String(図['枚数']) + '枚' : '',
      文: 止.length ? '作れなかった図解：' + 止.join('・') + '（制作担当が直します）' : '' });
  } else if (作['図解・挿絵']) 図解 = 行('図解・挿絵', '作成中', { 枚: String(作['図解・挿絵']['枚'] || '') });
  else 図解 = 行('図解・挿絵', '待ち');

  let スライド;
  if (st === '完了') スライド = 行('スライド', '済', { url: 観.スライドURL || '' });
  else if (st === 'スライド生成中') スライド = 行('スライド', '作成中');
  else if (st === 'エラー' && 取り込み済み) スライド = 行('スライド', '止まった', { 文: String(観.メッセージ || '').slice(0, 200) });
  else スライド = 行('スライド', '待ち');

  let 音声;
  if (F.音声) 音声 = 行('音声', '済', 済か(F.音声));
  else if ((F.台本txt && st === '完了') || 作['音声']) 音声 = 行('音声', '作成中');
  else 音声 = 行('音声', '待ち');

  const 系統 = { LP: [タイトル, サムネ, LP], 資料: [構成, 図解, スライド, 音声] };

  // ---- 【作業中】の心拍切れ → 止まった ----
  ['LP', '資料'].forEach(function (k) {
    系統[k] = 系統[k].map(function (it) {
      if (止まった作業[it.名] && it.状態 !== '済') return { 名: it.名, 状態: '止まった', 文: 止まった作業[it.名] };
      return it;
    });
  });

  // ---- 直し（受付中は該当する行を「作り直し中」に） ----
  const 直し = (記.直し || []).map(function (d) { return Object.assign({}, d); });
  const 受付中 = 直し.filter(function (d) { return d.状態 === '作り直し中'; });
  受付中.forEach(function (d) {
    const 名 = 直しの行[d.対象];
    if (!名) return;
    ['LP', '資料'].forEach(function (k) {
      系統[k] = 系統[k].map(function (it) {
        if (it.名 !== 名) return it;
        const o = { 名: 名, 状態: '作り直し中', あと分: 作[名] ? 残り分(名) : 10 };
        if (it.url) o.url = it.url;
        return o;
      });
    });
  });

  // ---- まとめ ----
  const 全部 = 系統.LP.concat(系統.資料);
  const 済の数 = 全部.filter(function (x) { return x.状態 === '済'; }).length;
  const 止まった = 全部.filter(function (x) { return x.状態 === '止まった'; });
  const 全部できた = 済の数 === 全部.length;
  const 系統の残り = function (k) {
    return 系統[k].reduce(function (s, it) {
      if (it.状態 === '作成中' || it.状態 === '作り直し中') return s + (it.あと分 || 0);
      if (it.状態 === '待ち') return s + (進み_目安分[it.名] || 5);
      return s;
    }, 0);
  };
  const あと分 = Math.max(系統の残り('LP'), 系統の残り('資料'));
  const 承認済 = !!(記.承認 && 記.承認.済);
  const 納品済 = !!(記.納品 && 記.納品.済);

  let いまここ;
  if (止まった.length) {
    いまここ = { 文: '「' + 止まった.map(function (x) { return x.名; }).join('」「') + '」が止まっています', あと分: 0,
      次: '止まった理由を見て、直らなければ宮津さんへ（押すものはありません）' };
  } else if (受付中.length) {
    いまここ = { 文: '直しています（' + 受付中.map(function (d) { return d.対象; }).join('・') + '）', あと分: あと分,
      次: '直し終わったら「確認して承認」' };
  } else if (!全部できた) {
    いまここ = { 文: '一気に作っています（' + 全部.length + 'つ中 ' + 済の数 + 'つできた）', あと分: あと分, 次: '全部できたら「確認して承認」。いま押すものはありません' };
  } else if (!承認済) {
    いまここ = { 文: '全部できました', あと分: 0, 次: 'LP・サムネイル・スライド・台本・音声を見て、よければ「確認して承認する」' };
  } else if (!納品済) {
    いまここ = { 文: '承認しました', あと分: 0, 次: '「納品フォルダへ入れる」' };
  } else {
    いまここ = { 文: '納品しました', あと分: 0, 次: '終わりです（直しが来たら、もう一度承認と納品）' };
  }

  return {
    案件名: 観.案件名 || '',
    流れ: 流れ_新,
    更新: 日時の文字_(いま),
    いまここ: いまここ,
    診断: 記.診断 || null,
    系統: 系統,
    足りない素材: 記.足りない素材 || [],
    直し: 直し,
    承認: 記.承認 || { 済: false, 人: '', 日時: '' },
    納品: 記.納品 || { 済: false, 日時: '' },
    動き: (記.動き || []).slice(-動き_最大),
    知らせ: (記.知らせ || []).slice(),   // 操作パネルのプロジェクトが、テスト案件のものだけ Chat の「テスト」へ出す
    構成シート: 観.構成シートURL || '',
    受付中: { 承認: !!観.承認ファイル, 納品: !!観.納品ファイル },   // パネルで押して、まだこのスクリプトが拾っていないもの
    全部できた: 全部できた,
    押せる: { 承認: 全部できた && !受付中.length && !止まった.length && !承認済, 納品: 承認済 && !受付中.length }
  };
}

// ============================================================
// 書く
// ============================================================
function 進みを書く_(folder, 進) {
  const text = JSON.stringify(進, null, 2);
  const it = filesByName_(folder, 進み_ファイル名);
  if (it.hasNext()) { it.next().setContent(text); return; }
  folder.createFile(進み_ファイル名, text, 'application/json');
}


/** 納品できたら覚える（B22 からもパネルからも） */
function 進み_納品を記録_(hearingId, 文) {
  進み_記録を変える_(hearingId, function (記) {
    記.納品 = { 済: true, 日時: 日時の文字_(new Date()) };
    動きを足す_(記, '納品フォルダへ入れました（' + String(文 || '').slice(0, 60) + '）');
  });
}
