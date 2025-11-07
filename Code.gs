/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of execution to the
 * current document only.
 */

/**
 * Adds a custom menu to the active document, containing a single menu item
 * for showing the sidebar.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createMenu('Content Assistant')
      .addItem('Show Sidebar', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (Unfortunately, e.authMode is not
 *     actually available for simple onInstall triggers.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document.
 */
function showSidebar() {
  console.log('🎨 Etapa 3: Criação da Sidebar');
  const html = HtmlService.createTemplateFromFile('index');
  const sidebarData = getJsonData(); // Fetch the data

  html.preloadedData = sidebarData; // Pass the already obtained data

  const sidebar = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME) // Using IFRAME for better security and compatibility
    .setTitle('Content Assistant')
    .setWidth(300);

  console.log('✅ Sidebar criada');
  DocumentApp.getUi().showSidebar(sidebar);
}

/**
 * Gets the content of the JSON file from Google Drive.
 *
 * @return {string} The content of the JSON file.
 */
function getJsonData() {
  var fileId = '1HM4CdqFY40hl7r1gdwrxkJqfhUEqrL4t';
  var file = DriveApp.getFileById(fileId);
  var content = file.getBlob().getDataAsString();
  return content;
}

/**
 * Inserts text into the active document.
 *
 * @param {string} text The text to insert.
 */
function insertText(text) {
  var cursor = DocumentApp.getActiveDocument().getCursor();
  if (cursor) {
    // The insertText method automatically handles splitting the text element
    // if the cursor is in the middle of it.
    cursor.insertText(text);
  } else {
    DocumentApp.getUi().alert('Could not find a cursor in the document. Please click in the document to place the cursor.');
  }
}

/**
 * A placeholder function to handle patient updates.
 * In a real application, this would update data based on the selected patient.
 *
 * @param {string} patientName The name of the selected patient.
 */
function updatePatient(patientName) {
  // For now, we'll just log the patient's name to the console.
  console.log('Selected patient:', patientName);
}

/**
 * Includes the content of another HTML file.
 * This is a common utility function for templated HTML in Google Apps Script.
 *
 * @param {string} filename The name of the file to include.
 * @return {string} The content of the file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
