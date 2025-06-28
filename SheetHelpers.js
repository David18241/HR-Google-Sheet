/**
 * @fileoverview Helper functions for interacting with Google Sheets.
 */

/**
 * Gets sheet data, headers, and a map of header names to column indices.
 * Handles potential errors gracefully.
 *
 * @param {string} sheetName The name of the sheet to retrieve data from.
 * @param {Spreadsheet} ss Optional. The specific spreadsheet object. Defaults to the active spreadsheet.
 * @returns {{sheet: Sheet, data: Object[][], headers: string[], headerMap: Object}|null} An object containing the sheet, data, headers, and header map, or null if sheet not found.
 */
function getSheetDataWithHeaders(sheetName, ss = SpreadsheetApp.getActiveSpreadsheet()) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Error: Sheet "${sheetName}" not found.`);
    SpreadsheetApp.getUi().alert(`Error: Sheet "${sheetName}" not found. Please check the sheet name.`);
    return null;
  }
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  if (data.length === 0) {
    Logger.log(`Error: Sheet "${sheetName}" is empty.`);
    SpreadsheetApp.getUi().alert(`Error: Sheet "${sheetName}" appears to be empty.`);
    return null;
  }

  const headers = data[0].map(String); // Ensure headers are strings
  const headerMap = {};
  headers.forEach((header, index) => {
    if (header) { // Only map non-empty headers
      headerMap[header] = index;
    }
  });

  return { sheet, data, headers, headerMap };
}

/**
 * Gets the data for the currently selected single row in a given sheet.
 * Validates that only one row is selected.
 *
 * @param {string} sheetName The name of the sheet where the row is selected.
 * @param {Object} headerMap An object mapping header names to column indices (from getSheetDataWithHeaders).
 * @param {Spreadsheet} ss Optional. The specific spreadsheet object. Defaults to the active spreadsheet.
 * @returns {{rowData: Object[], rowIndex: number, range: Range}|null} An object containing the row data array, its 1-based index, and the range, or null if selection is invalid.
 */
function getActiveRowData(sheetName, headerMap, ss = SpreadsheetApp.getActiveSpreadsheet()) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet "${sheetName}" not found.`);
      Logger.log(`Sheet "${sheetName}" not found.`);
      return null;
  }
  const activeRange = sheet.getActiveRange();
  const ui = SpreadsheetApp.getUi();

  if (activeRange.getNumRows() !== 1) {
    ui.alert("Please select exactly one row to perform this action.");
    Logger.log("Invalid row selection: Multiple rows or no rows selected.");
    return null;
  }

  const rowIndex = activeRange.getRow();
  // Check if selected row is the header row
  if (rowIndex === 1) {
      ui.alert("Please select an employee row, not the header row.");
      Logger.log("Invalid row selection: Header row selected.");
      return null;
  }

  const rowValues = activeRange.getValues()[0];

  // Basic validation: Check if essential data (e.g., name, email) might be missing
  if (!rowValues[headerMap[COL_FIRST_NAME]] || !rowValues[headerMap[COL_LAST_NAME]]) {
       ui.alert("The selected row appears to be missing essential data (like First or Last Name). Please check the row.");
       Logger.log(`Selected row ${rowIndex} seems incomplete.`);
       // Decide if you want to proceed or return null. Returning null is safer.
       return null;
  }

  return { rowData: rowValues, rowIndex: rowIndex, range: activeRange };
}

/**
 * Parses an employee name string typically formatted as "Last Name, First Name".
 *
 * @param {string} employeeNameString The name string.
 * @returns {{firstName: string, lastName: string, fullName: string}} An object with first, last, and full names. Returns empty strings if parsing fails.
 */
function parseEmployeeName(employeeNameString) {
    const nameParts = employeeNameString ? employeeNameString.split(', ') : [];
    const lastName = nameParts[0] ? nameParts[0].trim() : '';
    const firstName = nameParts[1] ? nameParts[1].trim() : '';
    const fullName = firstName && lastName ? `${firstName} ${lastName}` : (firstName || lastName); // Handle cases where only one name part exists
    return { firstName, lastName, fullName };
}

/**
 * Formats a JavaScript Date object into "Month Day, Year" format (e.g., "January 1, 2024").
 * Returns an empty string if the date is invalid.
 *
 * @param {Date} dateObject The Date object to format.
 * @returns {string} The formatted date string or "".
 */
function formatDate(dateObject) {
  if (!dateObject || !(dateObject instanceof Date) || isNaN(dateObject.getTime())) {
    Logger.log("Invalid date passed to formatDate.");
    return ""; // Return empty string for invalid dates
  }
  return dateObject.toLocaleDateString('en-US', {
    month: 'long',
    day: 'numeric',
    year: 'numeric'
  });
}