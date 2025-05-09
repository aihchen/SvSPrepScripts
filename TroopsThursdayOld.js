function generateTroopsThursdaySheet() {
  // Get original sheet
  const originalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 4");
  // Set new sheet name e.g., "Troops April 2025"
  const troopName = "Troops " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM yyyy");
  // Create new sheet with sheet name
  const troopSheet = createNewMonthlySheet(troopName);
  // Set up a template in the new sheet
  setupTroopTemplate(troopSheet);
  // Get filtered and sorted data from the original sheet
  // Then copy data identified in copyCustomHeaders and paste them in desired order
  const troopData = getTroopData(originalSheet);
  
  // Copy to new sheet
  // getRange(Row 7, Column 1 (i.e. A), Number of rows in your data, Number of columns in your data)
  if (troopData.length > 0) {
    troopSheet.getRange(7, 1, troopData.length, 9).setValues(troopData);
  }
  
  // Clean up availability column (I ➝ J)
  cleanAvailabilityColumn(troopSheet);
  // Populates 30 minute time intervals with corresponding private message
  populateTimeSlots(troopSheet);
  
  // Copy the headers from the form in the desired order
  copyTroopHeaders(originalSheet, troopSheet);

  // Always shift columns G to O (columns 7–15) over to H to P (columns 8–16), rows 7–107
  troopSheet.getRange(7, 7, 101, 9).moveTo(troopSheet.getRange(7, 8));
  // Add Total Speedups in column G and shift other data to the right
  addTotalSpeedups(troopSheet, troopData.length);
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
  sheet.getRange("M6").setBackground("#FFAAA5");
  sheet.getRange("N6").setBackground("#FFFFBA");
  sheet.getRange("O6").setBackground("#BAFFC9");
  sheet.getRange("P6").setBackground("#E7AFED");
  var headerRange = sheet.getRange("K6:P6");
  headerRange.setValues([["Cleaned availability", "Start Time (UTC)", "Schedule based on crystal shards", "Schedule based on limited availability", "Schedule based on hard-to-fill time slot", "Schedule based on total speedups"]]);
  headerRange.setWrap(true);
  headerRange.setVerticalAlignment("middle");
  headerRange.setHorizontalAlignment("center");
  sheet.setFrozenRows(6);
  sheet.getRange("J7:J107").setBorder(false, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
}

// Copy the headers from the form in the desired order
function copyTroopHeaders(sourceSheet, targetSheet) {
  const headerMap = [
    { source: "B1", target: "A6" },
    { source: "C1", target: "B6" },
    { source: "D1", target: "C6" },
    { source: "N1", target: "D6" },
    { source: "M1", target: "E6" },
    { source: "J1", target: "F6" },
    { source: "P1", target: "H6" },
    { source: "Q1", target: "I6" },
    { source: "O1", target: "J6" }
  ];

  headerMap.forEach(pair => {
    const value = sourceSheet.getRange(pair.source).getValue();
    targetSheet.getRange(pair.target).setValue(value);
  });
}

// Filter 'Yes' responses in column G ("Do you want Minister of Education on Thursday?") and 
// Sort by column N ("How many T9 soldiers will you promote to T10 on Thursday?") in descending order
// Copy the data under the desired headers and map them in desired order
function getTroopData(sheet) {
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // exclude headers

  const colE = 6; // 0-based index, filter by column G (index 6)
  const colH = 13; // sort by column I (index 13)

  const filtered = rows.filter(row => row[colE] === "Yes");

  const sorted = filtered.sort((a, b) => {
    const valA = parseFloat(a[colH]) || 0;
    const valB = parseFloat(b[colH]) || 0;
    return valB - valA;
  });

  // Only copy the data under the headers identified with copyCustomHeaders
  // Extract data from each row under these Columns[Index] of Form Response: B[1], C[2], D[3], N[13], M[12], J[9], P[15], Q[16], O[14] 
  const selectedColumns = sorted.map(row => [row[1], row[2], row[3], row[13], row[12], row[9], row[15], row[16], row[14]]);

  return selectedColumns;
}

// Strip away the Korean in column I (availability) and paste to column J
function cleanAvailabilityColumn(sheet) {
  const startRow = 7;
  const lastRow = sheet.getLastRow();
  const availabilityColumn = 9; // Column I
  const cleanedColumn = 10;     // Column J

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

// Populates 30 minute time intervals with corresponding private message
function populateTimeSlots(sheet) {
  const startRow = 7;
  const intervals = 48; // 30-minute intervals from 00:00 to 23:30
  const timeColumn = 11; // Column K
  sheet.getRange("L7:L54").setNumberFormat("@");

  for (let i = 0; i < intervals; i++) {
    const startTime = Utilities.formatDate(new Date(0, 0, 0, 0, i * 30), Session.getScriptTimeZone(), "HH:mm");

    const timeCell = sheet.getRange(startRow + i, timeColumn);

    timeCell.setNumberFormat("@"); // Set format to plain text
    timeCell.setValue(startTime);  // Set time as text
  }
}

function addTotalSpeedups(sheet, dataLength) {
  const startRow = 7;
  const endRow = startRow + dataLength - 1;

  // Set header in G6
  sheet.getRange("G6").setValue("Total Speedups");

  // Get values from E and F columns
  const colE = sheet.getRange(startRow, 5, dataLength).getValues(); // Column E
  const colF = sheet.getRange(startRow, 6, dataLength).getValues(); // Column F

  // Calculate totals and prepare for column G
  const totals = colE.map((row, i) => {
    const valE = parseFloat(row[0]) || 0;
    const valF = parseFloat(colF[i][0]) || 0;
    return [valE + valF];
  });

  // Set totals in column G
  sheet.getRange(startRow, 7, dataLength).setValues(totals);
}