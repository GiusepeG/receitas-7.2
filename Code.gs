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
      .addSeparator()
      .addItem('Update Cache from Sheet', 'updateJsonFromSheet')
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

  html.preloadedData = JSON.stringify(sidebarData); // Pass the already obtained data

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
  const fileId = '1HM4CdqFY40hl7r1gdwrxkJqfhUEqrL4t';
  try {
    const file = DriveApp.getFileById(fileId);
    const jsonString = file.getBlob().getDataAsString();
    const data = JSON.parse(jsonString);
    return data;
  } catch (e) {
    Logger.log('Error reading or parsing JSON file: ' + e);
    return { error: 'Failed to load data: ' + e.message };
  }
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

/**
 * Reads data from a Google Sheet, processes it, and saves it as a JSON file in Google Drive.
 * This function acts as a "cache generator" for the sidebar.
 */
function updateJsonFromSheet() {
  const sheetId = 'YOUR_GOOGLE_SHEET_ID_HERE'; // Replace with your actual Google Sheet ID
  const jsonFileId = '1HM4CdqFY40hl7r1gdwrxkJqfhUEqrL4t';

  try {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheets = spreadsheet.getSheets();
    let allItems = [];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith('_')) {
        return; // Skip tabs starting with an underscore
      }

      const data = sheet.getDataRange().getValues();
      // Start from the second row to skip the header
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const item = {
          tab: sheetName,
          color: row[0],
          title: row[1],
          tags: row[2] ? row[2].toString().split(' ').filter(tag => tag) : [],
          requirement: row[3],
          content: row[4]
        };
        allItems.push(item);
      }
    });

    const jsonString = JSON.stringify(allItems, null, 2);
    const jsonFile = DriveApp.getFileById(jsonFileId);
    jsonFile.setContent(jsonString);

    DocumentApp.getUi().alert('Successfully updated the JSON cache from the Google Sheet.');

  } catch (e) {
    Logger.log('Error updating JSON from Sheet: ' + e);
    DocumentApp.getUi().alert('Failed to update the JSON cache. Please check the logs for details.');
  }
}
