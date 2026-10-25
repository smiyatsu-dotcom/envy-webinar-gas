/*
 * Normal consultation-room geometry boundary.
 *
 * This file is deliberately not part of the deployed GAS checkout.  It is the
 * reviewed source to add to webinar-gas before enabling the runner mapping.
 * The public names do not overlap with the fixed isolation entrypoints.
 */

var NORMAL_GEOMETRY_SCHEMA_ = 'consultation-slide-geometry/v1';
var NORMAL_GEOMETRY_MAP_SCHEMA_ = 'consultation-slide-map/v2';

function normalGeometryStable_(value) {
  if (Array.isArray(value)) return '[' + value.map(normalGeometryStable_).join(',') + ']';
  if (value && typeof value === 'object') return '{' + Object.keys(value).sort().map(function (key) { return JSON.stringify(key) + ':' + normalGeometryStable_(value[key]); }).join(',') + '}';
  return JSON.stringify(value);
}

function normalGeometryHash_(value) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normalGeometryStable_(value), Utilities.Charset.UTF_8);
  return bytes.map(function (byte) { var number = byte < 0 ? byte + 256 : byte; return ('0' + number.toString(16)).slice(-2); }).join('');
}

function normalGeometryEqual_(left, right) {
  return !!left && !!right && ['left', 'top', 'width', 'height'].every(function (key) { return Number(left[key]) === Number(right[key]); });
}

function normalGeometryValid_(value) {
  return !!value && ['left', 'top', 'width', 'height'].every(function (key) { return typeof value[key] === 'number' && isFinite(value[key]) && value[key] >= 0; }) && value.width > 0 && value.height > 0;
}

function normalGeometryCopy_(value) { return JSON.parse(JSON.stringify(value)); }

function normalGeometryTargetRequest_(value) {
  // The bridge may nominate only a case and a human-facing slide number.  Deck,
  // page, and object IDs come solely from the management-ledger-owned map.
  if (!value || typeof value !== 'object' || typeof value.caseName !== 'string' || !value.caseName || value.caseName.length > 120 || !Number.isInteger(value.slideNo) || value.slideNo < 1 || ['presentationId', 'pageId', 'objectId'].some(function (key) { return Object.prototype.hasOwnProperty.call(value, key); })) throw new Error('案件台帳で解決した対象スライドだけを読めます');
  return { caseName:value.caseName, slideNo:value.slideNo };
}

function normalGeometrySelector_(element) {
  var selector = { type:String(element.type || ''), title:String(element.title || ''), description:String(element.description || ''), templateGeometry:normalGeometryCopy_(element.geometry) };
  // Editable Shapes are distinguished by their actual text.  IMAGE elements
  // intentionally have no text selector: visible characters baked into pixels
  // cannot be mistaken for a separately editable text element.
  if (selector.type === 'SHAPE') selector.text = String(element.text || '');
  return selector;
}

function normalGeometrySelectorEqual_(left, right) { return normalGeometryStable_(left) === normalGeometryStable_(right); }

function normalGeometryMapSelectorValid_(selector) {
  if (!selector || typeof selector.title !== 'string' || typeof selector.description !== 'string' || !normalGeometryValid_(selector.templateGeometry)) return false;
  if (selector.type === 'IMAGE') return selector.role === 'primary-image' && !Object.prototype.hasOwnProperty.call(selector, 'text');
  // A shape is eligible only when the generator recorded a concrete editable
  // role and its text.  This is deliberately not a generic SHAPE allowance.
  return selector.type === 'SHAPE' && selector.role === 'heading' && typeof selector.text === 'string' && selector.text.trim().length > 0 && selector.text.length <= 2000;
}

function normalGeometryMapSelectorMatches_(element, selector) {
  if (!normalGeometryMapSelectorValid_(selector) || !element) return false;
  var actual = normalGeometrySelector_(element);
  if (actual.type !== selector.type || actual.title !== selector.title || actual.description !== selector.description || !normalGeometryEqual_(actual.templateGeometry, selector.templateGeometry)) return false;
  return selector.type !== 'SHAPE' || actual.text === selector.text;
}

function normalGeometryPageHashes_(pages) {
  return pages.map(function (page) { return { id:page.id, contentHash:normalGeometryHash_({ id:page.id, elements:page.elements.map(function (element) { return { id:element.id, contentHash:element.contentHash }; }) }) }; });
}

function normalGeometryUntouched_(pages, pageId, objectId) {
  var items = [];
  pages.forEach(function (page) { page.elements.forEach(function (element) {
    if (page.id !== pageId || element.id !== objectId) items.push({ pageId:page.id, id:element.id, contentHash:element.contentHash, geometry:normalGeometryCopy_(element.geometry) });
  }); });
  return items;
}

function normalGeometryPlanHash_(plan) {
  var copy = {};
  Object.keys(plan || {}).forEach(function (key) { if (key !== 'planHash') copy[key] = plan[key]; });
  return normalGeometryHash_(copy);
}

// The production implementation is intentionally based on the same management
// row resolver used for the approval-time original check.  It never creates an
// output folder and it rejects missing or duplicate map files.
function normalGeometryRuntime_(cfg) {
  return {
    resolve: function (caseName) {
      var ledger = 相談室_管理行から原本を解決_(cfg, caseName);
      var company = caseCompany_(ledger.projectName);
      var folder = caseFolderIfExists_(cfg, company);
      if (!folder) throw new Error('案件台帳に対応する生成物フォルダがありません');
      var name = '【スライド対応】' + company + '.json';
      var files = filesByName_(folder, name);
      if (!files.hasNext()) throw new Error('案件台帳に対応するスライド対応表がありません');
      var file = files.next();
      if (files.hasNext()) throw new Error('案件台帳に対応するスライド対応表が複数あります');
      var map;
      try { map = JSON.parse(file.getBlob().getDataAsString('UTF-8')); }
      catch (error) { throw new Error('スライド対応表を読めません'); }
      if (!map || map.schema !== NORMAL_GEOMETRY_MAP_SCHEMA_ || map.caseName !== company || typeof map.presentationId !== 'string' || !map.presentationId || typeof map.presentationRevision !== 'string' || !map.presentationRevision || !Array.isArray(map.pages)) throw new Error('案件台帳に対応するスライド対応表の形式が不正です');
      return { caseName:company, hearingId:ledger.hearingId, presentationId:map.presentationId, presentationRevision:map.presentationRevision, pages:map.pages };
    },
    open: function (presentationId) { return SlidesApp.openById(presentationId); },
    presentationId: function (deck) { return deck.getId(); },
    revision: function (presentationId) { return DriveApp.getFileById(presentationId).getLastUpdated().toISOString(); },
    slides: function (deck) { return deck.getSlides(); },
    slideId: function (slide) { return slide.getObjectId(); },
    pageSize: function (deck) { return { width:deck.getPageWidth(), height:deck.getPageHeight() }; },
    elements: function (slide) { return slide.getPageElements(); },
    element: function (element) {
      var label = { type:String(element.getPageElementType()), title:String(element.getTitle() || ''), description:String(element.getDescription() || '') };
      var text = '';
      try { text = String(element.asShape().getText().asString() || ''); } catch (ignore) { /* images and non-shapes have no text */ }
      return { id:element.getObjectId(), type:label.type, title:label.title, description:label.description, text:text,
        geometry:{ left:element.getLeft(), top:element.getTop(), width:element.getWidth(), height:element.getHeight() },
        contentHash:normalGeometryHash_({ label:label, text:text }), raw:element };
    },
    setGeometry: function (element, geometry) { element.setLeft(geometry.left).setTop(geometry.top).setWidth(geometry.width).setHeight(geometry.height); }
  };
}

function normalGeometryResolvedTarget_(request, resolved, runtime) {
  if (!resolved || request.caseName !== resolved.caseName) throw new Error('案件台帳と提案の案件が一致しません');
  var hits = resolved.pages.filter(function (page) { return page && Number(page.slideNo) === Number(request.slideNo); });
  if (hits.length !== 1) throw new Error(hits.length ? '案件台帳のスライド対応が一意ではありません' : '案件台帳に対象スライドがありません');
  var mapping = hits[0];
  if (typeof mapping.pageId !== 'string' || !mapping.pageId || !normalGeometryMapSelectorValid_(mapping.targetSelector)) throw new Error('案件台帳のスライド対応に許可された対象要素がありません');
  if (typeof resolved.presentationRevision !== 'string' || !resolved.presentationRevision || runtime.revision(resolved.presentationId) !== resolved.presentationRevision) throw new Error('案件台帳のスライド対応表が資料の最新revisionと一致しません');
  var deck = runtime.open(resolved.presentationId);
  if (runtime.presentationId(deck) !== resolved.presentationId) throw new Error('開いた資料が案件台帳の資料と一致しません');
  var pages = runtime.slides(deck).filter(function (slide) { return runtime.slideId(slide) === mapping.pageId; });
  if (pages.length !== 1) throw new Error('案件台帳の対象ページを資料から一意に確認できません');
  var elements = runtime.elements(pages[0]).map(function (raw) { return runtime.element(raw); }).filter(function (element) { return normalGeometryMapSelectorMatches_(element, mapping.targetSelector); });
  if (elements.length !== 1) throw new Error(elements.length ? '案件台帳の対象selectorが一意ではありません' : '案件台帳の対象selectorを資料から確認できません');
  return { caseName:resolved.caseName, hearingId:resolved.hearingId, presentationId:resolved.presentationId, pageId:mapping.pageId, objectId:elements[0].id, slideNo:Number(request.slideNo), resolvedBy:'management-ledger' };
}

function normalGeometrySnapshot_(target, resolved, runtime, options) {
  var allowTargetGeometryChange = !!(options && options.allowTargetGeometryChange);
  if (!allowTargetGeometryChange && runtime.revision(target.presentationId) !== resolved.presentationRevision) throw new Error('案件台帳のスライド対応表が資料の最新revisionと一致しません');
  var deck = runtime.open(target.presentationId);
  if (runtime.presentationId(deck) !== target.presentationId) throw new Error('開いた資料が案件台帳の資料と一致しません');
  var size = runtime.pageSize(deck);
  var pages = runtime.slides(deck).map(function (slide) {
    var pageId = runtime.slideId(slide);
    var mapping = resolved.pages.filter(function (item) { return item && item.pageId === pageId; });
    if (mapping.length > 1) throw new Error('案件台帳のページ対応が重複しています');
    var slideNo = mapping.length ? Number(mapping[0].slideNo) : null;
    var elements = runtime.elements(slide).map(function (raw) {
      var element = runtime.element(raw);
      element.persistentSelector = { slideNo:slideNo, selector:normalGeometrySelector_(element) };
      delete element.raw;
      return element;
    });
    return { id:pageId, width:size.width, height:size.height, elements:elements };
  });
  var targetPage = pages.filter(function (page) { return page.id === target.pageId; });
  if (targetPage.length !== 1) throw new Error('案件台帳の対象ページを資料から確認できません');
  var targetElement = targetPage[0].elements.filter(function (element) { return element.id === target.objectId; });
  if (targetElement.length !== 1) throw new Error('案件台帳の対象要素を資料から確認できません');
  if (targetElement[0].persistentSelector.slideNo !== target.slideNo) throw new Error('案件台帳の対象ページ番号が一致しません');
  var targetMapping = resolved.pages.filter(function (item) { return item && item.pageId === target.pageId; });
  if (targetMapping.length !== 1 || !normalGeometryMapSelectorValid_(targetMapping[0].targetSelector) || (!allowTargetGeometryChange && !normalGeometryMapSelectorMatches_(targetElement[0], targetMapping[0].targetSelector))) throw new Error('案件台帳の対象selectorを資料から確認できません');
  return { presentationId:target.presentationId, revision:runtime.revision(target.presentationId), pageOrder:pages.map(function (page) { return page.id; }), pages:pages };
}

// Testable read core.  runtime is a small adapter; only normalGeometryRuntime_
// may touch Drive or Slides in production.
function normalGeometryRead_(request, runtime) {
  var requested = normalGeometryTargetRequest_(request);
  var resolved = runtime.resolve(requested.caseName);
  var target = normalGeometryResolvedTarget_(requested, resolved, runtime);
  var snapshot = normalGeometrySnapshot_(target, resolved, runtime);
  return { ok:true, caseName:target.caseName, target:target, snapshot:snapshot };
}

function normalGeometryReadbackMatches_(snapshot, target, plan, runtime, expectedGeometry) {
  var after = normalGeometrySnapshot_(target, runtime.resolve(target.caseName), runtime, { allowTargetGeometryChange:true });
  var page = after.pages.filter(function (item) { return item.id === target.pageId; })[0];
  var changed = page && page.elements.filter(function (item) { return item.id === target.objectId; })[0];
  return !!changed && normalGeometryEqual_(changed.geometry, expectedGeometry) && changed.contentHash === plan.contentHash && normalGeometryStable_(after.pageOrder) === normalGeometryStable_(snapshot.pageOrder) && normalGeometryStable_(normalGeometryPageHashes_(after.pages)) === normalGeometryStable_(normalGeometryPageHashes_(snapshot.pages)) && normalGeometryStable_(normalGeometryUntouched_(after.pages, target.pageId, target.objectId)) === normalGeometryStable_(plan.untouched);
}

function normalGeometryApply_(plan, runtime) {
  if (!plan || plan.schema !== NORMAL_GEOMETRY_SCHEMA_ || plan.status !== 'approved' || !plan.approvedBy || !Number.isInteger(plan.revision) || plan.planHash !== normalGeometryPlanHash_(plan)) throw new Error('承認済みの配置計画を確認できません');
  var read = normalGeometryRead_({ caseName:plan.caseName, slideNo:plan.persistentSelector && plan.persistentSelector.slideNo }, runtime);
  var snapshot = read.snapshot, target = read.target;
  if (plan.caseName !== target.caseName || plan.presentationId !== target.presentationId || plan.pageId !== target.pageId || plan.objectId !== target.objectId || plan.presentationRevision !== snapshot.revision || normalGeometryStable_(plan.pageOrder) !== normalGeometryStable_(snapshot.pageOrder) || normalGeometryStable_(plan.pageHashes) !== normalGeometryStable_(normalGeometryPageHashes_(snapshot.pages))) throw new Error('案件台帳または資料の版が提案時から変わりました');
  var page = snapshot.pages.filter(function (item) { return item.id === target.pageId; })[0];
  var element = page && page.elements.filter(function (item) { return item.id === target.objectId; })[0];
  if (!page || !element || !normalGeometryEqual_(element.geometry, plan.before) || element.contentHash !== plan.contentHash || !normalGeometryValid_(plan.before) || !normalGeometryValid_(plan.after) || plan.after.left + plan.after.width > page.width || plan.after.top + plan.after.height > page.height) throw new Error('対象要素または配置前後の値を確認できません');
  if (!plan.persistentSelector || Number(plan.persistentSelector.slideNo) !== target.slideNo || !normalGeometrySelectorEqual_(element.persistentSelector.selector, plan.persistentSelector.selector)) throw new Error('対象要素の再生成対応を確認できません');
  var matches = page.elements.filter(function (item) { return normalGeometrySelectorEqual_(item.persistentSelector.selector, plan.persistentSelector.selector); });
  if (matches.length !== 1 || matches[0].id !== element.id) throw new Error('対象要素のselectorが一意ではありません');
  if (normalGeometryStable_(plan.untouched) !== normalGeometryStable_(normalGeometryUntouched_(snapshot.pages, target.pageId, target.objectId))) throw new Error('対象外要素が提案時から変わりました');
  var deck = runtime.open(target.presentationId);
  var rawSlide = runtime.slides(deck).filter(function (slide) { return runtime.slideId(slide) === target.pageId; })[0];
  var rawElement = rawSlide && runtime.elements(rawSlide).map(function (raw) { return runtime.element(raw); }).filter(function (item) { return item.id === target.objectId; })[0];
  if (!rawElement) throw new Error('更新対象の実要素を確認できません');
  var writeStarted = false;
  try {
    writeStarted = true;
    runtime.setGeometry(rawElement.raw, plan.after);
    if (!normalGeometryReadbackMatches_(snapshot, target, plan, runtime, plan.after)) throw new Error('配置後の読戻しが一致しません');
    return { ok:true, proposalId:plan.proposalId, revision:plan.revision, planHash:plan.planHash, readback:true };
  } catch (error) {
    if (!writeStarted) throw error;
    var restoreError = null;
    try {
      runtime.setGeometry(rawElement.raw, plan.before);
      if (!normalGeometryReadbackMatches_(snapshot, target, plan, runtime, plan.before)) restoreError = new Error('元の配置の読戻しが一致しません');
    } catch (rollbackError) {
      restoreError = rollbackError;
    }
    if (restoreError) throw new Error('配置後の読戻しに失敗し、元の配置の復元または読戻し確認にも失敗しました。現物照合が必要です');
    throw new Error('配置後の読戻しが一致しません。元の配置へ復元し、読戻しを確認しました');
  }
}

// The deployed normal endpoints normally resolve only the production ledger.
// The one literal isolation case is allowed through the *same* resolver with
// its independently verified ledger configuration, so an end-to-end executable
// test never needs a customer case or a client-provided Drive/Slides id.  This
// is intentionally an exact-name exception: every other case retains the
// normal production configuration.
function normalGeometryConfigForCase_(caseName) {
  if (caseName !== ISOLATION_GEOMETRY_CASE) return getConfig_();
  var isolation = 隔離通し_生成設定を作る_();
  if (!isolation || isolation.ISOLATION_TEST !== true || isolation.MGMT_SHEET_ID !== 隔離通し_生成_台帳ID || isolation.OUTPUT_FOLDER_ID !== 隔離通し_生成_生成物rootID) {
    throw new Error('固定隔離案件の設定を確認できません');
  }
  return isolation;
}

function スライド_通常案件承認用geometryを読む(request) {
  return normalGeometryRead_(request, normalGeometryRuntime_(normalGeometryConfigForCase_(request && request.caseName)));
}

function スライド_通常案件承認済み配置を適用(plan) {
  return normalGeometryApply_(plan, normalGeometryRuntime_(normalGeometryConfigForCase_(plan && plan.caseName)));
}
