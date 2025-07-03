function goToNextHighlightedCell() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ss.getSheetByName("Logs");
  var inputSheet = ss.getSheetByName("Input Data");

  var range = logsSheet.getRange("A2:A" + logsSheet.getLastRow());
  var values = range.getValues().flat().filter(String); // Get non-empty values
  if (values.length === 0) return;

  // Store index in C1 of Input Data
  var indexCell = inputSheet.getRange("C1");
  var currentIndex = parseInt(indexCell.getValue()) || 0;

  if (currentIndex >= values.length - 1) {
    currentIndex = 0; // Reset to first if at the end
  } else {
    currentIndex++; // Move to next
  }

  indexCell.setValue(currentIndex); // Store updated index
  jumpToCell(inputSheet, values[currentIndex]);
}

function goToPreviousHighlightedCell() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ss.getSheetByName("Logs");
  var inputSheet = ss.getSheetByName("Input Data");

  var range = logsSheet.getRange("A2:A" + logsSheet.getLastRow());
  var values = range.getValues().flat().filter(String); // Get non-empty values
  if (values.length === 0) return;

  // Store index in C1 of Input Data
  var indexCell = inputSheet.getRange("C1");
  var currentIndex = parseInt(indexCell.getValue()) || 0;

  if (currentIndex <= 0) {
    currentIndex = values.length - 1; // Loop back to last item
  } else {
    currentIndex--; // Move to previous
  }

  indexCell.setValue(currentIndex); // Store updated index
  jumpToCell(inputSheet, values[currentIndex]);
}

function jumpToCell(sheet, cellAddress) {
  try {
    var cell = sheet.getRange(cellAddress);
    sheet.setActiveRange(cell);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Invalid cell reference: " + cellAddress);
  }
}
