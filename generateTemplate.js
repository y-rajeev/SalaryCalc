function generateNextMonth() {
  // Get the active spreadsheet and sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Loan & Advance");

  // Store original formulas from C:H (Header + Data)
  const rangeC1H3 = sheet.getRange("C1:H3"); 
  const headerFormulas = rangeC1H3.getFormulas(); // Store formulas for C1:H3
  
  const rangeC4H = sheet.getRange("C4:H" + sheet.getLastRow());
  const dataFormulas = rangeC4H.getFormulas(); // Store formulas for C4:H (DO NOT COPY VALUES)

  // Store column widths before insertion
  const columnWidths = [];
  for (let i = 3; i <= 8; i++) { // C=3 to H=8
    columnWidths.push(sheet.getColumnWidth(i));
  }

  // Store text values for D3:F3 (Headers)
  const headerTextRange = sheet.getRange("D3:F3");
  const headerTextValues = headerTextRange.getValues(); // Save headers separately

  // Get original month value for later update
  const originalMonthValue = sheet.getRange("C1").getValue();

  // Get formatting details (no values) for C:H
  const sourceRange = sheet.getRange("C:H");
  const sourceFormats = sourceRange.getBackgrounds();
  const sourceFontColors = sourceRange.getFontColors();
  const sourceFontWeights = sourceRange.getFontWeights();
  const sourceHorizontalAlignments = sourceRange.getHorizontalAlignments();
  const sourceVerticalAlignments = sourceRange.getVerticalAlignments();
  const numRows = sheet.getLastRow();

  // Insert new columns (Shift old C:H to I:N)
  sheet.insertColumnsAfter(2, 6);

  // Restore formatting (background, fonts, alignments) in new C:H
  const targetRange = sheet.getRange("C:H");
  targetRange.setBackgrounds(sourceFormats);
  targetRange.setFontColors(sourceFontColors);
  targetRange.setFontWeights(sourceFontWeights);
  targetRange.setHorizontalAlignments(sourceHorizontalAlignments);
  targetRange.setVerticalAlignments(sourceVerticalAlignments);

  // Restore column widths
  for (let i = 3; i <= 8; i++) {
    sheet.setColumnWidth(i, columnWidths[i - 3]);
  }

  // Copy borders using copyFormatToRange
  sheet.getRange("I1:N" + numRows).copyFormatToRange(sheet, 3, 8, 1, numRows);

  // Restore formulas to C1:H3 (Header)
  sheet.getRange("C1:H3").setFormulas(headerFormulas);

  // Restore headers (D3:F3)
  sheet.getRange("D3:F3").setValues(headerTextValues);

  // Restore formulas to C4:H (Data section)
  sheet.getRange("C4:H" + numRows).setFormulas(dataFormulas);

  // Calculate next month for C1
  const currentMonthStr = String(originalMonthValue).trim();
  let nextMonth;

  if (originalMonthValue instanceof Date) {
    const date = new Date(originalMonthValue);
    date.setMonth(date.getMonth() + 1);
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    nextMonth = monthNames[date.getMonth()] + "-" + date.getFullYear();
  } else {
    const parts = currentMonthStr.split("-");
    if (parts.length == 2) {
      const month = parts[0];
      const year = parseInt(parts[1]);
      const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const monthIndex = monthNames.indexOf(month);
      if (monthIndex !== -1) {
        let nextMonthIndex = monthIndex + 1;
        let nextYear = year;
        if (nextMonthIndex >= 12) {
          nextMonthIndex = 0;
          nextYear += 1;
        }
        nextMonth = monthNames[nextMonthIndex] + "-" + nextYear;
      } else {
        const today = new Date();
        const nextMonthDate = new Date(today);
        nextMonthDate.setMonth(today.getMonth() + 1);
        nextMonth = monthNames[nextMonthDate.getMonth()] + "-" + nextMonthDate.getFullYear();
      }
    } else {
      const today = new Date();
      const nextMonthDate = new Date(today);
      nextMonthDate.setMonth(today.getMonth() + 1);
      const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      nextMonth = monthNames[nextMonthDate.getMonth()] + "-" + nextMonthDate.getFullYear();
    }
  }

  // Set the new month value in C1
  sheet.getRange("C1").setValue(nextMonth);
}
