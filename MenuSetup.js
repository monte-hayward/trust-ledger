/**
 * @fileoverview Sets up the custom menu for the Trust Ledger application in the spreadsheet.
 * This script is container-bound and relies on SpreadsheetApp.getActiveSpreadsheet().
 */

const TrustUtils = TrustUtilsLib;

// --- CONFIGURATION CONSTANTS (Local only: Column indices) ---
// Column index (1-indexed) where the action button ("add proof...") is located (Column H)
const PROOF_ACTION_COL_INDEX = 8; 
// Column index (1-indexed) where the final Proof Link URL is stored (Column G)
const PROOF_LINK_COL_INDEX = 7; 
const LEDGER_SHEET_NAME = 'LEDGER'; // Defined locally as the sheet name is only needed here and should not be a library constant

/**
 * Runs automatically when the spreadsheet is opened. Creates the custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('⚙️ Trust Automation')
      .addItem('Run Discover Card Import', 'runDiscoverImport')
      .addItem('---', 'noop')
      .addItem('Undo Last Import', 'runUndoLastImport')
      .addItem('---', 'noop')
      .addItem('Authorize Drive Access', 'authorizeDriveAccess')
      .addToUi();
}

/**
 * Custom menu function that calls the Discover Card processing logic 
 * from the external TrustUtils library.
 */
function runDiscoverImport() {
  try {
    const ledger = SpreadsheetApp.getActiveSpreadsheet();
    // Pass the ledger object explicitly to the utility library for execution.
    TrustUtils.processRawDiscoverCardData_Mapped(ledger); 
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Import Failed: ${e.message}`);
    Logger.log(e);
  }
}

/**
 * Custom menu function that calls the Undo logic from the external TrustUtils library.
 */
function runUndoLastImport() {
  try {
    const ledger = SpreadsheetApp.getActiveSpreadsheet();
    TrustUtils.undoLastImport(ledger);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Undo Failed: ${e.message}`);
    Logger.log(e);
  }
}

/**
 * Forces the authorization dialog to appear for the DriveApp scope.
 */
function authorizeDriveAccess() {
  // A simple call to the DriveApp service forces the authorization prompt.
  try {
    DriveApp.getFolders(); 
    SpreadsheetApp.getUi().alert("Drive access authorized or already enabled. Try the import again.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Please check your browser pop-up permissions to ensure the authorization dialog is visible.");
    Logger.log(`Failed to run DriveApp.getFolders(): ${e.message}`);
  }
}

/**
 * Runs automatically when a cell is manually edited.
 * If the edit occurs in the Proof Action column, this launches the file picker dialog or attempts auto-match.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onEdit(e) {
  // onEdit uses the event object which is intrinsically bound to the active spreadsheet.
  const sheet = e.range.getSheet();
  if (sheet.getName() !== LEDGER_SHEET_NAME) return;

  const editedCol = e.range.getColumn();
  const editedRow = e.range.getRow();
  
  // Check if the edit occurred in the designated Proof Action column (Column H / index 8)
  if (editedCol === PROOF_ACTION_COL_INDEX) {
    const ui = SpreadsheetApp.getUi();
    const cell = e.range;
    
    // Get transaction details for use in the matching function
    const transactionRange = sheet.getRange(editedRow, 1, 1, 4); 
    const [, , vendorStr, amount] = transactionRange.getValues()[0]; 

    // Restore the button formula immediately so the user can click again
    const buttonFormula = `=HYPERLINK("#", "add proof...")`;
    cell.setFormula(buttonFormula); 
    
    // --- 1. Attempt Automatic Receipt Match (File Name Metadata) ---
    const autoLink = TrustUtils.findReceiptByMetadata(vendorStr, amount); 
    
    if (autoLink) {
      // If a match is found, apply the hyperlink formula directly to the Proof Link column (G)
      const linkRange = sheet.getRange(editedRow, PROOF_LINK_COL_INDEX);
      linkRange.setFormula(`=HYPERLINK("${autoLink}", "Receipt (Auto-Match)")`);
      ui.alert(`Auto-match successful! Link added to cell ${linkRange.getA1Notation()}.`);
    } else {
      // --- 2. If no auto-match, launch Manual File Picker Dialog ---
      // Pass the current row and column to ensure the result goes back to the correct cell.
      TrustUtils.showReceiptDialog(editedRow, PROOF_LINK_COL_INDEX);
    }
  }
}

/**
 * Dummy function for the separator menu item.
 */
function noop() {
  // Does nothing. Used for menu separators.
}
