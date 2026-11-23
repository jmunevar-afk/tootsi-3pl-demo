/**
 * Code.gs — web app entry point.
 *
 * Two deployments of this same script:
 *  1) INTERNAL hub  → Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone within Tootsi
 *  2) 3PL portal    → Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 *     3PLs open the portal link with their private code: <PORTAL_URL>?k=<token>
 *
 * The script always runs as you. The Sheet and the Drive folder are never shared with the 3PLs:
 * they only ever see what portalLoad() returns (see portalRow() in Logic.gs).
 */
function doGet(e) {
  var k = e && e.parameter && e.parameter.k;
  var t;
  if (k) {
    try { var u = tplUser_(k); } catch (err) { return page_('Message', { msg: err.message }); }
    t = HtmlService.createTemplateFromFile('Portal');
    t.token = k; t.wh = u.wh; t.name = u.name; t.lang = String(config_()['LANG_' + u.wh] || 'fr');
    return t.evaluate().setTitle('Tootsi · 3PL Portal').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }
  try { hubUser_(); } catch (err) { return page_('Message', { msg: err.message }); }
  t = HtmlService.createTemplateFromFile('Hub');
  return t.evaluate().setTitle('Tootsi · PO Hub').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function page_(name, vars) {
  var t = HtmlService.createTemplateFromFile(name);
  Object.keys(vars || {}).forEach(function (k) { t[k] = vars[k]; });
  return t.evaluate().setTitle('Tootsi PO Hub').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Used inside HTML templates: <?!= include('Styles') ?> */
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }
