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
      return HtmlService.createTemplateFromFile('Hub').evaluate().setTitle('Tootsi · PO Hub').addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) { /* not a Tootsi Google account: show the 3PL sign-in page */ }
  }
  // 3PL partners: the page opens on the sign-in screen (username + password from the Users tab).
  var t = HtmlService.createTemplateFromFile('Portal');
  t.lang = 'fr'; t.hubUrl = String(config_().HUB_URL || '');
  return t.evaluate().setTitle('Tootsi · 3PL Portal').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function page_(name, vars) {
  var t = HtmlService.createTemplateFromFile(name);
  Object.keys(vars || {}).forEach(function (k) { t[k] = vars[k]; });
  return t.evaluate().setTitle('Tootsi PO Hub').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Used inside HTML templates: <?!= include('Styles') ?> */
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }
