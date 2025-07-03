function highlightCells() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Data");
  if (!sheet) {
    Logger.log("Sheet not found!");
    return;
  }

  var range = sheet.getDataRange();
  var values = range.getValues();
  var numRows = values.length;
  var numCols = values[0].length;

  var statusColIndex = 5; // Column F (zero-based index)
  var statusRows = [];
  var logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");

  if (!logsSheet) {
    logsSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Logs");
  }

  // Clear previous logs
  logsSheet.getRange("A3:A").clear();
  logsSheet.appendRow(["Highlighted Cell Addresses"]); // Header

  // Step 1: Clear all background colors from F2 to the last column
  var clearRange = sheet.getRange(2, statusColIndex + 1, numRows - 1, numCols - statusColIndex);
  clearRange.setBackground(null); // Removes any existing highlights

  // Find all occurrences of "Status" in column F
  for (var row = 0; row < numRows; row++) {
    if (values[row][statusColIndex] === "Status") {
      statusRows.push(row);
    }
  }

  var cellsToHighlight = [];
  var logEntries = [];

  // Loop through each "Status" row
  statusRows.forEach(function (statusRow) {
    var rowIndex = statusRow; // Zero-based index

    // Loop through columns in this row to find "P"
    for (var col = 0; col < numCols; col++) {
      if (values[rowIndex][col] === "P") {
        // Check 1 row and 2 rows below
        var below1 = rowIndex + 1 < numRows ? values[rowIndex + 1][col] : null;
        var below2 = rowIndex + 2 < numRows ? values[rowIndex + 2][col] : null;

        // If any one of the below cells is blank, mark for highlighting
        if (below1 === "" || below1 === null || below2 === "" || below2 === null) {
          var cell = sheet.getRange(rowIndex + 1, col + 1); // Convert to 1-based index
          cellsToHighlight.push(cell);

          var cellAddress = cell.getA1Notation();
          logEntries.push([cellAddress]); // Store cell address in log
        }
      }
    }
  });

  // Apply red background to all identified cells
  if (cellsToHighlight.length > 0) {
    cellsToHighlight.forEach(cell => cell.setBackground("red"));
  }

  // Log addresses in "Logs" sheet
  if (logEntries.length > 0) {
    logsSheet.getRange(3, 1, logEntries.length, 1).setValues(logEntries);
  }

  Logger.log(logEntries.length + " cells highlighted and logged.");
}
