/**
 * MenuSetup.gs (BOUND TO LEDGER SPREADSHEET)
 * Handles custom menu creation and the onEdit trigger for interactive receipt linking.
 * * NOTE: This script REQUIRES the TrustUtils library to be linked with the identifier 'TrustUtils' 
 * * and the library version must be updated every time core logic changes.
 */

// --- CONFIGURATION CONSTANTS (Local only: Column indices) ---
// Column index (1-indexed) where the action button ("add proof...") is located (Column H)
const PROOF_ACTION_COL_INDEX = 8; 
// Column index (1-indexed) where the final Proof Link URL is stored (Column G)
const PROOF_LINK_COL_INDEX = 7; 
const LEDGER_SHEET_NAME = 'LEDGER'; // Defined locally as the sheet name is only needed here and should not be a library constant

/**
 * Runs automatically when the spreadsheet is opened. Creates the custom menu.
 * MAY need to run authorizeDriveAccess from the custom menu before manually running onOpen,
 * to fix Failed to access drive folder: Unexpected error while getting the method or property getFolderById on object DriveApp
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('⚙️ Trust Automation')
      .addItem('Run Discover Card Import', 'runDiscoverImport')
      .addItem('---', 'noop')
      .addItem('Undo Last Import', 'runUndoLastImport')
      .addItem('---', 'noop')
      .addItem('Authorize Drive Access', 'authorizeDriveAccess') // <-- NEW MENU ITEM
      .addToUi();
}

/**
 * Custom menu function that calls the Discover Card processing logic 
 * from the external TrustUtils library.
 */
function runDiscoverImport() {
  try {
    TrustUtils.processRawDiscoverCardData_Mapped();
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
    TrustUtils.undoLastImport();
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Undo Failed: ${e.message}`);
    Logger.log(e);
  }
}

/**
 * NEW FUNCTION: Forces the authorization dialog to appear for the DriveApp scope.
 * The user must run this manually from the menu if an access error occurs.
 * Fixes error: Failed to access drive folder: Unexpected error while getting the method or property getFolderById on object DriveApp
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
 * Runs automatically when a cell is manually edited (e.g., user clicks the "add proof..." formula and presses Enter).
 * If the edit occurs in the Proof Action column, this launches the file picker dialog or attempts auto-match.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== LEDGER_SHEET_NAME) return;

  const editedCol = e.range.getColumn();
  const editedRow = e.range.getRow();
  
  // Check if the edit occurred in the designated Proof Action column (Column H / index 8)
  if (editedCol === PROOF_ACTION_COL_INDEX) {
    const ui = SpreadsheetApp.getUi();
    const cell = e.range;
    
    // Get transaction details for use in the matching function
    // Assuming Date (Col A), Source (Col B), Payee/Vendor (Col C), Amount (Col D)
    const transactionRange = sheet.getRange(editedRow, 1, 1, 4); 
    const [, , vendorStr, amount] = transactionRange.getValues()[0]; 

    // Restore the button formula immediately so the user can click again
    const buttonFormula = `=HYPERLINK("#", "add proof...")`;
    cell.setFormula(buttonFormula); 
    
    // --- 1. Attempt Automatic Receipt Match (File Name Metadata) ---
    // Calls the library function to search by metadata
    const autoLink = TrustUtils.findReceiptByMetadata(vendorStr, amount); 
    
    if (autoLink) {
      // If a match is found, apply the hyperlink formula directly to the Proof Link column (G)
      const linkRange = sheet.getRange(editedRow, PROOF_LINK_COL_INDEX);
      linkRange.setFormula(`=HYPERLINK("${autoLink}", "Receipt (Auto-Match)")`);
      ui.alert(`Auto-match successful! Link added to cell ${linkRange.getA1Notation()}.`);
    } else {
      // --- 2. If no auto-match, launch Manual File Picker Dialog ---
      // Call the library function to display the file picker UI
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
