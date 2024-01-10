function splitProspectsList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // Prompt for the new sheet's name
  var sheetNameResponse = ui.prompt('Enter the name for the new sheet');
  var newSheetName = sheetNameResponse.getResponseText();

  if (ss.getSheetByName(newSheetName)) {
    ui.alert('Sheet name already exists. Please choose a different name.');
    return;
  }

  // Prompt for the number of rows to copy
  var numRowsResponse = ui.prompt('How many rows do you want to copy (excluding header)?', ui.ButtonSet.OK_CANCEL);
  var numRows = parseInt(numRowsResponse.getResponseText(), 10);

  if (isNaN(numRows) || numRows < 1) {
    ui.alert('Invalid number of rows. Please enter a valid number.');
    return;
  }

  // Create a new sheet with the specified name
  var newSheet = ss.insertSheet(newSheetName);

  // Calculate the range to copy
  var headerRange = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn());
  var dataRange = sourceSheet.getRange(2, 1, numRows, sourceSheet.getLastColumn());

  // Copy the header
  headerRange.copyTo(newSheet.getRange(1, 1));

  // Copy the data
  dataRange.copyTo(newSheet.getRange(2, 1));

  // Delete the copied rows from the original sheet
  sourceSheet.deleteRows(2, numRows);

  // Remove empty rows and columns from the new sheet
  removeEmptyRowsAndColumns(newSheet);

  // Resize columns based on the width of headers
  resizeColumnsBasedOnHeaders(sourceSheet, newSheet);

  // Freeze the first row in the new sheet
  newSheet.setFrozenRows(1);

  ui.alert('Rows have been successfully copied and removed from the original sheet. Empty rows and columns have been cleaned up, columns resized, and the first row frozen in the new sheet. Created by EcomMedia.co');
}

// Function to remove empty rows and columns
function removeEmptyRowsAndColumns(sheet) {
  var maxRows = sheet.getMaxRows(); 
  var lastRow = sheet.getLastRow();
  if (maxRows - lastRow > 0) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }

  var maxColumns = sheet.getMaxColumns();
  var lastColumn = sheet.getLastColumn();
  if (maxColumns - lastColumn > 0) {
    sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
  }
}

// Function to resize columns based on the width of headers
function resizeColumnsBasedOnHeaders(sourceSheet, newSheet) {
  var numColumns = sourceSheet.getLastColumn();
  for (var i = 1; i <= numColumns; i++) {
    var columnWidth = sourceSheet.getColumnWidth(i);
    newSheet.setColumnWidth(i, columnWidth);
  }
}

// Adding a custom menu to run the script
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Prospect Tools')
    .addItem('Split Prospect List', 'splitProspectsList')
    .addToUi();
}
