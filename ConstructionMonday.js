function generateConstructionMondaySheet() {
  // Get original sheet
  const originalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 5");
  // Set new sheet name e.g., "Construction April 2025"
  const constructionName = "Construction " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM yyyy");
  // Create new sheet with sheet name
  const constructionSheet = createNewMonthlySheet(constructionName);
  // Set up a template in the new sheet
  setupConstructionTemplate(constructionSheet);
  // Get filtered and sorted data from the original sheet
  const constructionData = getConstructionData(originalSheet);

  // Copy to new sheet starting at row 7
  // getRange(Row 7, Column 1 (i.e. A), Number of rows in your data, Number of columns in your data)
  if (constructionData.length > 0) {
    constructionSheet.getRange(7, 1, constructionData.length, 10).setValues(constructionData);
  }
  
  // Clean up availability column (J ➝ K)
  cleanConstructionAvailabilityColumn(constructionSheet);

  // Copy the headers from the form in the desired order
  copyConstructionHeaders(originalSheet, constructionSheet);
  
  // Insert a column to the left of H and shift everything to the right. H(8) ➝ I(9)
  // getRange(Row 7, Column 8 (H), Number of rows in your data, Number of columns in your data (this seems to not be important))
  constructionSheet.getRange(7, 8, constructionData.length, 9).moveTo(constructionSheet.getRange(7, 9));

  // Add Total Speedups in column G
  addConstructionTotalSpeedups(constructionSheet, constructionData.length);

  // Populates 30 minute time intervals
  populateConstructionTimeSlots(constructionSheet);
}

// Create a new sheet named after the current month
function createNewMonthlySheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let constructionSheet = ss.getSheetByName(sheetName);
  if (constructionSheet) ss.deleteSheet(constructionSheet); // optional: delete if it already exists
  constructionSheet = ss.insertSheet(sheetName);
  return constructionSheet;
}

// Set up a template in the new sheet (customize this as needed)
function setupConstructionTemplate(sheet) {
  sheet.getRange("A1").setValue(sheet.getName());
  sheet.getRange("C1").setValue("Monday");
  sheet.getRange("D1").setValue("Vice President");
  sheet.getRange("A1:D1").setFontWeight("bold");
  sheet.getRange("F1").setValue("DUPLICATE players are highlighted in grey").setBackground("grey");
  sheet.getRange("A2").setValue("UNSCHEDULED players based on FIRE CRYSTALS are highlighted in red").setBackground("#FFAAA5");
  sheet.getRange("A3").setValue("UNSCHEDULED players based on LIMITED AVAILABILITY are highlighted in yellow").setBackground("#FFFFBA");
  sheet.getRange("A4").setValue("UNSCHEDULED players based on TIME SLOTS are highlighted in green").setBackground("#BAFFC9");
  sheet.getRange("A5").setValue("UNSCHEDULED players based on TOTAL SPEEDUPS are highlighted in purple").setBackground("#E7AFED");
  sheet.getRange("N6").setBackground("#FFAAA5");
  sheet.getRange("O6").setBackground("#FFFFBA");
  sheet.getRange("P6").setBackground("#BAFFC9");
  sheet.getRange("Q6").setBackground("#E7AFED");
  var headerRange = sheet.getRange("L6:Q6");
  headerRange.setValues([["Cleaned availability", "Start Time (UTC)", "Schedule based on fire crystals", "Schedule based on limited availability", "Schedule based on hard-to-fill time slots", "Schedule based on total speedups"]]);
  headerRange.setWrap(true);
  headerRange.setVerticalAlignment("middle");
  headerRange.setHorizontalAlignment("center");
  sheet.setFrozenRows(6);
  sheet.getRange("K7:K107").setBorder(false, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
}

// Copy the headers from the form in the desired order
function copyConstructionHeaders(sourceSheet, targetSheet) {
  const headerMap = [
    { source: "B1", target: "A6" },// Alliance name
    { source: "C1", target: "B6" },// In-game name
    { source: "D1", target: "C6" },// Furnace level
    { source: "H1", target: "D6" },// Flexible - fire crystal
    { source: "I1", target: "E6" },// Sort - refined crystal
    { source: "P1", target: "F6" },// Specific speedup
    { source: "O1", target: "G6" },// General speedup
    { source: "S1", target: "I6" },// Enough RSS?
    { source: "U1", target: "J6" },// Comments
    { source: "T1", target: "K6" }// Availability
  ];

  headerMap.forEach(pair => {
    const value = sourceSheet.getRange(pair.source).getValue();
    targetSheet.getRange(pair.target).setValue(value);
  });
}

// Filter 'Yes' responses in column E ("Do you want Vice President on Monday?") and 
// Sort by column I ("How many refined crystals are you going to spend on Monday?") in descending order
// Copy the data under the desired headers and map them in desired order
function getConstructionData(sheet) {
  const constructionData = sheet.getDataRange().getValues();
  const rows = constructionData.slice(1); // exclude headers

  const colE = 4; // 0-based index, filter by column E (index 4) in response form
  const colH = 8; // sort by column I (index 8) in response form

  const filtered = rows.filter(row => row[colE] === "Yes");

  const sorted = filtered.sort((a, b) => {
    const valA = parseFloat(a[colH]) || 0;
    const valB = parseFloat(b[colH]) || 0;
    return valB - valA;
  });

  // Only copy the data under the headers identified with copyCustomHeaders
  // Extract data from each row under these columns[index] of form response: 
  // B[1], C[2], D[3], H[7], I[8], P[15], O[14], S[18], U[20], T[19] 
  const selectedColumns = sorted.map(row => [row[1], row[2], row[3], row[7], row[8], row[15], row[14], row[18], row[20], row[19]]);

  return selectedColumns;
}

// Strip away the Korean in column J (availability) and paste to column K
function cleanConstructionAvailabilityColumn(sheet) {
  const startRow = 7;
  const lastRow = sheet.getLastRow();
  const availabilityColumn = 10; // Column J
  const cleanedColumn = 11;     // Column K

  const availabilityData = sheet.getRange(startRow, availabilityColumn, lastRow - startRow + 1).getValues();

  const cleanedData = availabilityData.map(row => {
    const cell = row[0];
    if (!cell) return [''];
    
    // Extract only the UTC portions like "0-1 UTC"
    const utcRanges = cell
      .split(',')
      .map(part => part.trim().match(/^([\d]{1,2})-([\d]{1,2}) UTC/))
      .filter(match => match);

    // Convert each UTC range to 30-minute intervals
    const intervals = [];
    const padded = (n) => n.toString().padStart(2, '0');
    
    utcRanges.forEach(match => {
      const start = parseInt(match[1], 10);
      const end = parseInt(match[2], 10);
      for (let h = start; h < end; h++) {
        intervals.push(`${padded(h)}:00`);
        intervals.push(`${padded(h)}:30`);
      }
    });

    return [intervals.join(', ')];
  });

  sheet.getRange(startRow, cleanedColumn, cleanedData.length, 1).setValues(cleanedData);
}

// Populates 30 minute time intervals
function populateConstructionTimeSlots(sheet) {
  const startRow = 7;
  const intervals = 48; // 30-minute intervals from 00:00 to 23:30
  const timeColumn = 13; // Column M
  sheet.getRange("M7:M54").setNumberFormat("@");

  for (let i = 0; i < intervals; i++) {
    const startTime = Utilities.formatDate(new Date(0, 0, 0, 0, i * 30), Session.getScriptTimeZone(), "HH:mm");

    const timeCell = sheet.getRange(startRow + i, timeColumn);

    timeCell.setNumberFormat("@"); // Set format to plain text
    timeCell.setValue(startTime);  // Set time as text
  }
}

function addConstructionTotalSpeedups(sheet, dataLength) {
  const startRow = 7;
  const endRow = startRow + dataLength - 1;

  // Set header in H6
  sheet.getRange("H6").setValue("Total Speedups");

  // Get values from F and G columns
  const colE = sheet.getRange(startRow, 6, dataLength).getValues(); // Column F
  const colF = sheet.getRange(startRow, 7, dataLength).getValues(); // Column G

  // Calculate totals and prepare for column H
  const totals = colE.map((row, i) => {
    const valE = parseFloat(row[0]) || 0;
    const valF = parseFloat(colF[i][0]) || 0;
    return [valE + valF];
  });

  // Set totals in column H
  sheet.getRange(startRow, 8, dataLength).setValues(totals);
}