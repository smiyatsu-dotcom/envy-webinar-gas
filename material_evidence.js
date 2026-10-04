/** 納品writerと同じ記録・探索関数を使うread-only証拠。Driveを変更しない。 */
function materialDeliveryEvidence_() {
  const cfg = getConfig_();
  const rows = mgmtSheet_(cfg).getDataRange().getValues();
  const out = [];
  rows.slice(1).forEach(function (row) {
    if (!新の流れか_(row) || !row[1]) return;
    const project = String(row[1]), company = caseCompany_(project);
    let srcFolder = null, dest = null;
    try { srcFolder = caseFolderIfExists_(cfg, project); } catch (e) {}
    try { dest = deliveryFolderOf_(company); } catch (e) {}
    const progress = 進み_記録を読む_(String(row[0]));
    const delivered = !!(progress && progress.納品 && progress.納品.済);
    const artifacts = [];
    if (delivered && srcFolder && dest) {
      [['スライド','slides'],['音声','audio'],['LP','lp'],['サムネイル','thumbnail']].forEach(function (pair) {
        const rec = deliveredRecord_(dest, pair[0]);
        let dst = null, src = null;
        try { dst = rec && rec.id ? DriveApp.getFileById(rec.id) : null; } catch (e) {}
        try {
          if (pair[0] === 'サムネイル') {
            const ai = foldersByName_(DriveApp.getFolderById(cfg.ASSETS_FOLDER_ID), project);
            if (ai.hasNext()) { const fi = filesByName_(ai.next(), 'サムネイル.png'); if (fi.hasNext()) src = fi.next(); }
          } else {
            const name = pair[0] === 'スライド' ? '【スライド】' + project : '【' + pair[0] + '】' + company + (pair[0] === '音声' ? '.mp3' : '');
            const fi = filesByName_(srcFolder, name); if (fi.hasNext()) src = fi.next();
          }
        } catch (e) {}
        artifacts.push(materialFileEvidence_(pair[1], src, dst));
      });
    }
    out.push({ caseName: company, deliveryRecorded: delivered, outputFolderId: srcFolder ? srcFolder.getId() : null, deliveryFolderId: dest ? dest.getId() : null, artifacts: artifacts });
  });
  return { schema: 1, checkedAt: new Date().toISOString(), enumerationSucceeded: true, targetCount: out.length, cases: out };
}

function materialFileEvidence_(kind, src, dst) {
  function meta(file) {
    if (!file) return null;
    let digest = null, readable = false;
    try { digest = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, file.getBlob().getBytes())); readable = true; } catch (e) {}
    return { id: file.getId(), mime: file.getMimeType(), digest: digest, readable: readable };
  }
  const s = meta(src), d = meta(dst);
  return { kind: kind, sourceFileId: s && s.id, destinationFileId: d && d.id, sourceMimeType: s && s.mime, destinationMimeType: d && d.mime,
    sourceDigest: s && s.digest, destinationDigest: d && d.digest, destinationParentId: dst && dst.getParents().hasNext() ? dst.getParents().next().getId() : null,
    nativeContentMatch: s && d && s.readable && d.readable ? s.digest === d.digest : null };
}

/** Bridgeが読む共有Drive上の証拠JSONを更新する。納品物・ACLは変更しない。 */
function writeMaterialDeliveryEvidence_() {
  const name = '【状態】資料実物_webinar-gas.json';
  const evidence = materialDeliveryEvidence_();
  // 0件も「管理シート列挙成功・対象0」という監視証拠。ファイル不在（unknown）と区別する。
  const text = JSON.stringify(evidence);
  const cfg = getConfig_();
  const root = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID).getParents().next();
  const it = root.getFilesByName(name);
  if (it.hasNext()) it.next().setContent(text);
  else root.createFile(name, text, MimeType.PLAIN_TEXT);
}
