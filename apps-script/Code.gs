/**
 * Code.gs — web app entry point.
 *
 * Two deployments of this same script:
 *  1) INTERNAL hub  → Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone within Tootsi
 *  2) 3PL portal    → Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 *     3PLs open the portal link and sign in with the username and password from addThreePL().
 *     Tootsi staff opening any link go straight to the hub (Google account check). Add ?view=3pl to test the 3PL screen.
 *
 * The script always runs as you. The Sheet and the Drive folder are never shared with the 3PLs:
 * they only ever see what portalLoad() returns (see portalRow() in Logic.gs).
 */
function doGet(e) {
  var view = e && e.parameter && e.parameter.view;
  // Tootsi staff: Google has already signed them in. The domain check is the login.
  if (view !== '3pl') {
    try {
      hubUser_();
      return tpl_('Hub').evaluate().setTitle('Tootsi · PO Hub').addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) { /* not a Tootsi Google account: show the 3PL sign-in page */ }
  }
  // 3PL partners: the page opens on the sign-in screen (username + password from the Users tab).
  var t = tpl_('Portal');
  t.lang = 'fr'; t.hubUrl = String(config_().HUB_URL || '');
  return t.evaluate().setTitle('Tootsi · 3PL Portal').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function page_(name, vars) {
  var t = tpl_(name);
  Object.keys(vars || {}).forEach(function (k) { t[k] = vars[k]; });
  return t.evaluate().setTitle('Tootsi PO Hub').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Used inside HTML templates: <?!= include('Styles') ?> */
/** The receiving form code, sent as data (fallback if the page builder damaged the inline copy). Public page code only. */
function clientCode(name) {
  if (['Receive'].indexOf(name) < 0) throw new Error('Not available.');
  var h = include(name), m = h.match(/<script[^>]*>([\s\S]*)<\/script>/);
  return m ? m[1] : '';
}
function include(name) {
  return typeof HTML_FILES !== 'undefined' ? HTML_FILES[name] : HtmlService.createHtmlOutputFromFile(name).getContent();
}
/** The one-file install (PO_Hub.gs) carries the pages inside HTML_FILES. The multi-file install uses the .html files. */
function tpl_(name) {
  return typeof HTML_FILES !== 'undefined' ? HtmlService.createTemplate(HTML_FILES[name]) : HtmlService.createTemplateFromFile(name);
}
