function generateTroopsThursdaySheet() {
  // Get original sheet
  const originalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 5");
  // Set new sheet name e.g., "Troops April 2025"
  const troopName = "Troops " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM yyyy");
  // Create new sheet with sheet name
  const troopSheet = createNewMonthlySheet(troopName);
  // Set up a template in the new sheet
  setupTroopTemplate(troopSheet);
  // Get filtered and sorted data from the original sheet
  const troopData = getTroopData(originalSheet);
  
  // Copy to new sheet
  // getRange(Row 7, Column 1 (i.e. A), Number of rows in your data, Number of columns in your data)
  if (troopData.length > 0) {
    troopSheet.getRange(7, 1, troopData.length, 10).setValues(troopData);
  }
  
  // Clean up availability column (J ➝ K)
  cleanTroopsAvailabilityColumn(troopSheet);
  
  // Copy the headers from the form in the desired order
  copyTroopHeaders(originalSheet, troopSheet);

  // Insert a column to the left of H and shift everything to the right. H(8) ➝ I(9)
  // getRange(Row 7, Column 8 (H), Number of rows in your data, Number of columns in your data (this seems to not be important))
  troopSheet.getRange(7, 8, 101, 9).moveTo(troopSheet.getRange(7, 9));
  
  // Add Total Speedups in column G
  addTroopsTotalSpeedups(troopSheet, troopData.length);

  // Populates 30 minute time intervals
  populateTroopsTimeSlots(troopSheet);
}

// Create a new sheet named after the current month
function createNewMonthlySheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let troopSheet = ss.getSheetByName(sheetName);
  if (troopSheet) ss.deleteSheet(troopSheet); // optional: delete if it already exists
  troopSheet = ss.insertSheet(sheetName);
  return troopSheet;
}

// Set up a template in the new sheet (customize this as needed)
function setupTroopTemplate(sheet) {
  sheet.getRange("A1").setValue(sheet.getName());
  sheet.getRange("C1").setValue("Thursday");
  sheet.getRange("D1").setValue("Minister of Education");
  sheet.getRange("A1:D1").setFontWeight("bold");
  sheet.getRange("F1").setValue("DUPLICATE players are highlighted in grey").setBackground("grey");
  sheet.getRange("A2").setValue("UNSCHEDULED players based on T9 TROOPS are highlighted in red").setBackground("#FFAAA5");
  sheet.getRange("A3").setValue("UNSCHEDULED players based on LIMITED AVAILABILITY are highlighted in yellow").setBackground("#FFFFBA");
  sheet.getRange("A4").setValue("UNSCHEDULED players based on TIME SLOTS are highlighted in blue").setBackground("#BAFFC9");
  sheet.getRange("A5").setValue("UNSCHEDULED players based on TOTAL SPEEDUPS are highlighted in purple").setBackground("#E7AFED");
  sheet.getRange("N6").setBackground("#FFAAA5");
  sheet.getRange("O6").setBackground("#FFFFBA");
  sheet.getRange("P6").setBackground("#BAFFC9");
  sheet.getRange("Q6").setBackground("#E7AFED");
  var headerRange = sheet.getRange("L6:Q6");
  headerRange.setValues([["Cleaned availability", "Start Time (UTC)", "Schedule based on crystal shards", "Schedule based on limited availability", "Schedule based on hard-to-fill time slot", "Schedule based on total speedups"]]);
  headerRange.setWrap(true);
  headerRange.setVerticalAlignment("middle");
  headerRange.setHorizontalAlignment("center");
  sheet.setFrozenRows(6);
  sheet.getRange("K7:K107").setBorder(false, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
}

// Copy the headers from the form in the desired order
function copyTroopHeaders(sourceSheet, targetSheet) {
  const headerMap = [
    { source: "B1", target: "A6" },// Alliance name
    { source: "C1", target: "B6" },// In-game name
    { source: "D1", target: "C6" },// Furnace level
    { source: "K1", target: "D6" },// Flexible - highest troop level
    { source: "N1", target: "E6" },// Sort - troop promotion
    { source: "R1", target: "F6" },// Specific speedup
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

// Filter 'Yes' responses in column G ("Do you want Minister of Education on Thursday?") and 
// Sort by column N ("How many total soldiers will you promote on Thursday?") in descending order
// Copy the data under the desired headers and map them in desired order
function getTroopData(sheet) {
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // exclude headers

  const colG = 6;   // 0-based index, filter by column G (index 6) in response form
  const colN = 13;  // sort by column N (index 13) in response form

  const filtered = rows.filter(row => row[colG] === "Yes");

  const sorted = filtered.sort((a, b) => {
    const valA = parseFloat(a[colN]) || 0;
    const valB = parseFloat(b[colN]) || 0;
    return valB - valA;
  });

  // Compose combined troop string from K (10), L (11), M (12)
  const selectedColumns = sorted.map(row => {
    const infantry = row[10] ? `${row[10]} Infantry` : '';
    const lancer = row[11] ? `${row[11]} Lancer` : '';
    const marksman = row[12] ? `${row[12]} Marksman` : '';

    const troopString = [infantry, lancer, marksman].filter(Boolean).join(", ");

    // Build the row: B, C, D, [Troop String], N, R, O, S, U, T
    return [
      row[1], // Alliance Name (B)
      row[2], // In-game Name (C)
      row[3], // Furnace Level (D)
      troopString, // Column D in new sheet
      row[13], // Troop Promotion Count (N)
      row[17], // Specific Speedup (R)
      row[14], // General Speedup (O)
      row[18], // Enough RSS? (S)
      row[20], // Comments (U)
      row[19]  // Availability (T)
    ];
  });

  return selectedColumns;
}

// Strip away the Korean in column J (availability) and paste to column K
function cleanTroopsAvailabilityColumn(sheet) {
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
function populateTroopsTimeSlots(sheet) {
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

function addTroopsTotalSpeedups(sheet, dataLength) {
  const startRow = 7;
  const endRow = startRow + dataLength - 1;

  // Set header in H6
  sheet.getRange("H6").setValue("Total Speedups");

  // Get values from E and F columns
  const colE = sheet.getRange(startRow, 6, dataLength).getValues(); // Column F
  const colF = sheet.getRange(startRow, 7, dataLength).getValues(); // Column G

  // Calculate totals and prepare for column G
  const totals = colE.map((row, i) => {
    const valE = parseFloat(row[0]) || 0;
    const valF = parseFloat(colF[i][0]) || 0;
    return [valE + valF];
  });

  // Set totals in column H
  sheet.getRange(startRow, 8, dataLength).setValues(totals);
}

function combineTroopLevels(sheet) {
  const startRow = 7;
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - startRow + 1;

  const infantryData = sheet.getRange(startRow, 11, numRows).getValues(); // Column K
  const lancerData = sheet.getRange(startRow, 12, numRows).getValues();   // Column L
  const marksmanData = sheet.getRange(startRow, 13, numRows).getValues(); // Column M

  const combined = infantryData.map((_, i) => {
    const parts = [];

    if (infantryData[i][0]) parts.push(`Infantry ${infantryData[i][0]}`);
    if (lancerData[i][0]) parts.push(`Lancer ${lancerData[i][0]}`);
    if (marksmanData[i][0]) parts.push(`Marksman ${marksmanData[i][0]}`);

    return [parts.join(", ")];
  });

  // Set combined results into column D
  sheet.getRange(startRow, 4, combined.length).setValues(combined);
}