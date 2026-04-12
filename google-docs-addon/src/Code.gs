/**
 * WordToWP — Google Docs add-on
 * Opens a sidebar for settings, formatting, and WordPress publishing (coming in later steps).
 */

/**
 * Runs when the document is opened. Adds the add-on menu under Extensions.
 */
function onOpen() {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem("Open WordToWP", "showSidebar")
    .addToUi();
}

/**
 * Required for add-on homepage / first install (can stay minimal).
 */
function onInstall() {
  onOpen();
}

/**
 * Shows the HTML sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("WordToWP")
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Sample server function: returns document name and a short text preview (proves Docs access from the sidebar).
 * Called via google.script.run from Sidebar.html.
 */
function getActiveDocumentInfo() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const full = body.getText();
  const preview = full.length > 400 ? full.substring(0, 400) + "…" : full;
  return {
    name: doc.getName(),
    preview: preview || "(empty document)",
  };
}
