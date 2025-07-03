function multi_run_optimized_preserveData() {
  applyFormulaBasedOnConditionOptimizedPreserveData();
  updateAttendanceStatus();
  updateFormulaBasedOnStatus();
  applyFormulaToDurationRows();
  applyFormulaToOTRows();
  applyFormulaForWOP();
}

function applyFormulaBasedOnConditionOptimizedPreserveData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Input Data');
  if (!sheet) {
    Logger.log("Input Data Sheet not found");
    return;
  }

  var lastRow = sheet.getLastRow(); // Get the last row with data
  var lastCol = sheet.getLastColumn(); // Get the last column with data
  var rangeF = sheet.getRange(3, 6, lastRow - 2, 1).getValues(); // Get F3:F
  var row2Data = sheet.getRange(2, 7, 1, lastCol - 6).getValues()[0]; // Get values in row 2, columns G:lastCol
  var existingData = sheet.getRange(3, 7, lastRow - 2, lastCol - 6).getValues(); // Get existing data in the target range

  var outputToApply = [];

  for (var i = 0; i < rangeF.length; i++) {
    var valueF = rangeF[i][0]; // Get the value in column F for this row
    var rowIndex = i + 3; // Adjust for row index (from row 3)
    var rowData = [];

    // Apply formulas only if the value in column F is "Late By" or "Early By"
    if (valueF === "Late By" || valueF === "Early By") {
      for (var colIndex = 7; colIndex <= lastCol; colIndex++) {
        if (row2Data[colIndex - 7] !== "") { // If corresponding row 2 column is not blank
          var cellRef = sheet.getRange(rowIndex, colIndex).getA1Notation(); // Get A1 notation of the current cell
          var cellCol = cellRef.replace(/[0-9]/g, ''); // Extract the column letter (e.g., "H")
          var cellRow = rowIndex - 3; // Adjust row number by subtracting 3

          // Construct the formula using the adjusted row number
          var formula = (valueF === "Late By") 
            ? `IF(${cellCol}${cellRow}<TIME(9,21,0), 0, ${cellCol}${cellRow}-C3)` // Formula for "Late By"
            : `IF(ISBLANK(${cellCol}${cellRow}),0,IF(${cellCol}${cellRow}>TIME(17, 0, 0), 0, D3-${cellCol}${cellRow}))`; // Formula for "Early By"
          
          rowData.push(`=${formula}`); // Apply the formula
        } else {
          rowData.push(existingData[i][colIndex - 7]); // Keep the existing value
        }
      }
    } else {
      rowData = existingData[i]; // Keep the entire row's existing data
    }

    outputToApply.push(rowData); // Add the row of formulas or existing values
  }
  
  // Apply all formulas and preserve existing data where no formulas are needed
  sheet.getRange(3, 7, outputToApply.length, lastCol - 6).setValues(outputToApply); // Use setValues to avoid adding "=" sign to existing data
}


function applyFormulaToDurationRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Input Data');
  if (!sheet) {
    Logger.log("Input Data Sheet not found");
    return;
  }

  var lastRow = sheet.getLastRow(); // Get the last row with data
  var lastCol = sheet.getLastColumn(); // Get the last column with data
  var rangeF = sheet.getRange(3, 6, lastRow - 2, 1).getValues(); // Get values from F3:F (Column F)
  var row2Data = sheet.getRange(2, 7, 1, lastCol - 6).getValues()[0]; // Get values in row 2 from G to the last column
  var existingData = sheet.getRange(3, 7, lastRow - 2, lastCol - 6).getValues(); // Get existing data in the target range (G3:lastCol)

  var outputToApply = [];

  for (var i = 0; i < rangeF.length; i++) {
    var valueF = rangeF[i][0]; // Get the value in column F for this row
    var rowIndex = i + 3; // Adjust for row index (starts from row 3)
    var rowData = [];

    // Apply formulas only if the value in column F is "Duration"
    if (valueF === "Duration") {
      for (var colIndex = 7; colIndex <= lastCol; colIndex++) {
        if (row2Data[colIndex - 7] !== "") { // If the corresponding column in row 2 is not blank
          var formula = `=INDIRECT(ADDRESS(ROW()-1, COLUMN())) - INDIRECT(ADDRESS(ROW()-2, COLUMN()))`;
          rowData.push(formula); // Apply the formula in this cell
        } else {
          rowData.push(existingData[i][colIndex - 7]); // Keep the existing value
        }
      }
    } else {
      rowData = existingData[i]; // Keep the entire row's existing data
    }

    outputToApply.push(rowData); // Add the row of formulas or existing values to output
  }

  // Apply all formulas and preserve existing data where no formulas are needed
  sheet.getRange(3, 7, outputToApply.length, lastCol - 6).setValues(outputToApply); // Apply to range starting from G3
}


function applyFormulaToOTRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Input Data');
  if (!sheet) {
    Logger.log("Input Data Sheet not found");
    return;
  }

  var lastRow = sheet.getLastRow(); // Get the last row with data
  var lastCol = sheet.getLastColumn(); // Get the last column with data
  var rangeF = sheet.getRange(3, 6, lastRow - 2, 1).getValues(); // Get values from F3:F (Column F)
  var row2Data = sheet.getRange(2, 7, 1, lastCol - 6).getValues()[0]; // Get values in row 2 from G to the last column
  var existingData = sheet.getRange(3, 7, lastRow - 2, lastCol - 6).getValues(); // Get existing data in the target range (G3:lastCol)

  var outputToApply = [];

  for (var i = 0; i < rangeF.length; i++) {
    var valueF = rangeF[i][0]; // Get the value in column F for this row
    var rowIndex = i + 3; // Adjust for row index (starts from row 3)
    var rowData = [];

    // Apply formulas only if the value in column F is "Duration"
    if (valueF === "OT") {
      for (var colIndex = 7; colIndex <= lastCol; colIndex++) {
        if (row2Data[colIndex - 7] !== "") { // If the corresponding column in row 2 is not blank
          var formula = [
            "=LET(",
              "  checkTime, INDIRECT(ADDRESS(ROW()-5, COLUMN())),",
              "  endTime, INDIRECT(ADDRESS(ROW()-4, COLUMN())),",
              "  isCheckTimeValid, ISNUMBER(checkTime),",
              "  isEndTimeValid, ISNUMBER(endTime),",
              "  otHour1,",
              "    IF(",
              "      NOT(isCheckTimeValid),",
              "      0,",
              "      IF(checkTime=TIME(0,0,0),",
              "        0,",
              "        IF(checkTime<TIME(8,10,0),",
              "          TIME(9,0,0) - checkTime,",
              "          0",
              "        )",
              "      )",
              "    ),",
              "  otHour2,",
              "    IF(",
              "      AND(isEndTimeValid, endTime>TIME(17,59,0)),",
              "      endTime - TIME(17,30,0),",
              "      0",
              "    ),",
              "  totalOT, otHour1 + otHour2,",
              "  IF(totalOT<0, 0, totalOT)",
              ")"
            ].join("\n");

          rowData.push(formula); // Apply the formula in this cell
        } else {
          rowData.push(existingData[i][colIndex - 7]); // Keep the existing value
        }
      }
    } else {
      rowData = existingData[i]; // Keep the entire row's existing data
    }

    outputToApply.push(rowData); // Add the row of formulas or existing values to output
  }

  // Apply all formulas and preserve existing data where no formulas are needed
  sheet.getRange(3, 7, outputToApply.length, lastCol - 6).setValues(outputToApply); // Apply to range starting from G3
}

function updateAttendanceStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Data");
  const data = sheet.getDataRange().getValues();
  const color = '#FFFF00'; // Set the desired background color (yellow)

  // Loop through each row in the sheet
  for (let row = 0; row < data.length; row++) {
    const rowData = data[row];

    // Check if column F contains 'Status'
    if (rowData[5] === 'Status') {

      // Loop through the attendance data after column F
      for (let col = 6; col < rowData.length; col++) {

        // Check if the current cell contains 'WO'
        if (rowData[col] === 'WO') {

          // Check previous and next columns
          const beforeWO = rowData[col - 1];
          const afterWO = (col + 1 < rowData.length) ? rowData[col + 1] : null;

          // Determine if we need to check one step before and after if current is blank
          const beforeBeforeWO = (col - 2 >= 0) ? rowData[col - 2] : null;
          const afterAfterWO = (col + 2 < rowData.length) ? rowData[col + 2] : null;

          // If 'A' is found before 'WO' and after 'WO', overwrite 'WO' with 'WOA'
          const beforeCondition = (beforeWO === 'A' || (beforeWO === '' && beforeBeforeWO === 'A'));
          const afterCondition = (afterWO === 'A' || (afterWO === '' && afterAfterWO === 'A'));

          if (beforeCondition && afterCondition) {
            // Update the cell and set the background color
            sheet.getRange(row + 1, col + 1).setValue('WOA');
            sheet.getRange(row + 1, col + 1).setBackground(color);
          }
        }
      }
    }
  }
}


function updateFormulaBasedOnStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Data");
  const data = sheet.getDataRange().getValues();

  // Loop through each row in the sheet
  for (let row = 0; row < data.length; row++) {
    const rowData = data[row];

    // Check if column F contains 'Status'
    if (rowData[5] === 'Status') {
      // Find the first occurrence of 'P' in the same row after column F
      for (let col = 6; col < rowData.length; col++) {
        if (rowData[col] === 'P') {
          // Convert the column index to letter
          const columnLetter = getColumnLetter(col + 1); // col is 0-indexed, so add 1
          
          // Create the formula using the entire row
          const formula = `=COUNTA(${columnLetter}${row + 1}:${columnLetter}${row + 1}) - COUNTIF(${columnLetter}${row + 1}:${columnLetter}${row + 1}, "A") - COUNTIF(${columnLetter}${row + 1}:${columnLetter}${row + 1}, "WOA")`;
          
          // Set the formula in column C of the same row
          sheet.getRange(row + 1, 3).setFormula(`=COUNTA(${columnLetter}${row + 1}:${row + 1}) - COUNTIF(${columnLetter}${row + 1}:${row + 1}, "A") - COUNTIF(${columnLetter}${row + 1}:${row + 1}, "WOA")`); // Column C is index 3
          break; // Exit the loop once 'P' is found
        }
      }
    }
  }
}

// Helper function to convert column index to letter
function getColumnLetter(column) {
  let temp;
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp) / 26);
  }
  return letter;
}


// Clear the data in the "Input Data" sheet
function clearDataInInputData() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Input Data");
    if (sheet) {
      // Define the range to clear
      var range = sheet.getRange("F2:AQ");
      range.clearContent();
      range.setBackground(null);
    }
}




function applyFormulaForWOP() {
  const sheetName = "Input Data";
  const searchRange = "F:AQ";
  const wopEarly = 5;
  const wopOT = 6;

  // Get the target sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }

  // Get the range and values to search for "WOP"
  const range = sheet.getRange(searchRange);
  const values = range.getValues();

  // Loop through the range to find "WOP"
  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      if (values[row][col] === "WOP") {
        // Calculate the target row (current row + offset)
        const targetRow = row + 1 + wopOT; // +1 because arrays are 0-indexed
        const secondTargetRow = row + 1 + wopEarly; // Early for WOP
        const targetCol = col + 6; // column F (col 6 in absolute terms)

        // Ensure targetRow does not exceed sheet bounds
        if (targetRow > sheet.getMaxRows()) {
          continue; // Skip if out of bounds
        }

        // Ensure secondTargetRow does not exceed sheet bounds
        if (secondTargetRow > sheet.getMaxRows()) {
          continue; // Skip if out of bounds
        }

        // Explicitly log the row and column to confirm behavior
        Logger.log(`Found WOP at row ${row + 1}, column ${col + 1}. Applying formulas at rows ${targetRow} and ${secondTargetRow}, column ${targetCol}.`);

        // First formula - applied 6 rows below
        const targetCell1 = sheet.getRange(targetRow, targetCol);
        const formula1 = `=IF(INDIRECT(ADDRESS(ROW()-3, COLUMN()))=TIME(0,0,0), 0, INDIRECT(ADDRESS(ROW()-3, COLUMN())) + TIME(0,30,0))`;
        targetCell1.setFormula(formula1);

        // Second formula - applied 5 rows below
        const targetCell2 = sheet.getRange(secondTargetRow, targetCol);
        const formula2 = `=IF(D4 > INDIRECT(ADDRESS(ROW()-3, COLUMN())), D4-INDIRECT(ADDRESS(ROW()-3, COLUMN())), 0)`;
        targetCell2.setFormula(formula2);
      }
    }
  }
}