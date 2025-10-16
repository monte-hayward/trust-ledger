/**
 * MenuSetup.js (BOUND TO LEDGER SPREADSHEET)
 * Handles custom menu creation and the onEdit trigger for interactive receipt linking.
 * This file is responsible for determining the execution environment.
 * NOTE: This script REQUIRES the TrustUtils library to be linked with the identifier 'TrustUtilsLib'.
 */

// --- ENVIRONMENT CONFIGURATION ---
// PROD_LEDGER_SCRIPT_ID is used to determine if the script is running in the production environment.
const PROD_LEDGER_SCRIPT_ID = '1XZ9JSz7GlYkTfqY06pXK4iLzitORsktkvMKS-xDC6V-U8eahxgXwBKI2'; 

// --- CONFIGURATION CONSTANTS (Local only) ---
const LEDGER_SHEET_NAME = 'LEDGER'; 
const PROOF_ACTION_COL_INDEX = 8; 
const PROOF_LINK_COL_INDEX = 7; 

/**
 * Calculates whether the current execution context is the production environment.
 * @returns {boolean} True if the current script ID matches the PROD_LEDGER_SCRIPT_ID.
 */
function isProductionEnvironment() {
  const currentScriptId = ScriptApp.getScriptId();
  return currentScriptId === PROD_LEDGER_SCRIPT_ID;
}

/**
 * Runs automatically when the spreadsheet is opened. Creates the custom menu.
 */
function onOpen() {
  const isProd = isProductionEnvironment();
  const envLabel = isProd ? 'PRODUCTION' : 'STAGING';
  
  SpreadsheetApp.getUi()
      .createMenu(`⚙️ Trust Automation (${envLabel})`) // Adds the environment label to the menu
      .addItem('Run Discover Card Import', 'runDiscoverImport')
      .addItem('---', 'noop')
      .addItem('Undo Last Import', 'runUndoLastImport')
      .addItem('---', 'noop')
      .addItem('Authorize Drive Access', 'authorizeDriveAccess') 
      .addToUi();
}

/**
 * Custom menu function that calls the Discover Card processing logic 
 * from the external TrustUtils library, passing the environment flag.
 */
function runDiscoverImport() {
  const ledger = SpreadsheetApp.getActiveSpreadsheet();
  const isProd = isProductionEnvironment();
  try {
    // Pass the ledger object and the environment flag to the utility library.
    TrustUtilsLib.processRawDiscoverCardData_Mapped(ledger, isProd);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Import Failed: ${e.message}`);
    Logger.log(e);
  }
}

/**
 * Custom menu function that calls the Undo logic from the external TrustUtils library.
 * This also needs to pass the isProd flag if the undo logic might involve environment-specific actions.
 */
function runUndoLastImport() {
  const ledger = SpreadsheetApp.getActiveSpreadsheet();
  const isProd = isProductionEnvironment();
  try {
    // Assuming undoLastImport might also need the ledger and environment status
    TrustUtilsLib.undoLastImport(ledger, isProd); 
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Undo Failed: ${e.message}`);
    Logger.log(e);
  }
}

/**
 * Forces the authorization dialog to appear for the DriveApp scope.
 */
function authorizeDriveAccess() {
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
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== LEDGER_SHEET_NAME) return;

  const editedCol = e.range.getColumn();
  const editedRow = e.range.getRow();
  
  // Check if the edit occurred in the designated Proof Action column
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
    // Note: This function doesn't need isProd, as finding a receipt is read-only.
    const autoLink = TrustUtilsLib.findReceiptByMetadata(vendorStr, amount); 
    
    if (autoLink) {
      const linkRange = sheet.getRange(editedRow, PROOF_LINK_COL_INDEX);
      linkRange.setFormula(`=HYPERLINK("${autoLink}", "Receipt (Auto-Match)")`);
      ui.alert(`Auto-match successful! Link added to cell ${linkRange.getA1Notation()}.`);
    } else {
      // --- 2. If no auto-match, launch Manual File Picker Dialog ---
      TrustUtilsLib.showReceiptDialog(editedRow, PROOF_LINK_COL_INDEX);
    }
  }
}

/**
 * Dummy function for the separator menu item.
 */
function noop() {
  // Does nothing. Used for menu separators.
}
