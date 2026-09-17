/**
 * ALPINO CRM — MAIN ENTRY POINT
 * ------------------------------
 * doGet() runs every time someone opens the Web App URL. It hands back
 * the Index.html page. include() is a small helper that lets Index.html
 * pull in Stylesheet.html and JavaScript.html — Apps Script can't link
 * to external .css/.js files the way a normal website does, so every
 * piece has to be its own HTML file stitched together at request time.
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Alpino CRM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
