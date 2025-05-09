function createMonthlyLayoutSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");
  const year = now.getFullYear();
  const sheetName = `Layout ${monthName} ${year}`;

  // Check if sheet already exists
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear(); // If it exists, clear its contents
  }

  formatScheduleHeaders(sheet);
  populateTimeAndMessages(sheet);
  styleBordersAndBackgrounds(sheet);
  setColumnWidths(sheet);
}

// Set up a template in the new sheet (customize this as needed)
function formatScheduleHeaders(sheet) {
  // Header labels
  const groupHeaders = [
    { range: "B2:F2", label: "CONSTRUCTION MONDAY" },
    { range: "H2:L2", label: "RESEARCH TUESDAY" },
    { range: "N2:R2", label: "TROOPS THURSDAY" }
  ];

  const columnHeaders = [
    { range: "B3", label: "Start Time (UTC)" },
    { range: "D3", label: "Scheduled in game?" },
    { range: "E3", label: "Sent private message?" },
    { range: "F3", label: "Message Copy" },
    { range: "H3", label: "Start Time (UTC)" },
    { range: "J3", label: "Scheduled in game?" },
    { range: "K3", label: "Sent private message?" },
    { range: "L3", label: "Message Copy" },
    { range: "N3", label: "Start Time (UTC)" },
    { range: "P3", label: "Scheduled in game?" },
    { range: "Q3", label: "Sent private message?" },
    { range: "R3", label: "Message Copy" }
  ];

  // Merge and style group headers
  groupHeaders.forEach(header => {
    const range = sheet.getRange(header.range);
    range.merge();
    range.setValue(header.label);
    range.setFontWeight("bold");
    range.setHorizontalAlignment("center");
    range.setFontSize(14);
    range.setBackground("#FFDFBA");
  });

  // Set individual column headers
  columnHeaders.forEach(header => {
    const range = sheet.getRange(header.range);
    range.setValue(header.label);
    range.setFontWeight("bold");
    range.setWrap(true);
    range.setVerticalAlignment("middle");
    range.setHorizontalAlignment("center");
  });

  sheet.setFrozenRows(3);
}

function populateTimeAndMessages(sheet) {
  const startRow = 4;
  const intervals = 48; // 30-minute intervals from 00:00 to 23:30
  const timeColumns = [2, 8, 14]; // B, H, N
  const messageColumns = [6, 12, 18]; // F, L, R
  const roles = [
    "Vice President time for Construction on Monday",
    "Vice President time for Research on Tuesday",
    "Minister of Education time for Troop Training on Thursday"
  ];

  for (let i = 0; i < intervals; i++) {
    const startTime = Utilities.formatDate(new Date(0, 0, 0, 0, i * 30), Session.getScriptTimeZone(), "HH:mm");
    const endMinutes = (i + 1) * 30;
    const endHour = Math.floor(endMinutes / 60);
    const endMinute = endMinutes % 60;
    const endTime = Utilities.formatDate(new Date(0, 0, 0, endHour, endMinute), Session.getScriptTimeZone(), "HH:mm");

    for (let j = 0; j < timeColumns.length; j++) {
      const timeCol = timeColumns[j];
      const messageCol = messageColumns[j];
      const row = startRow + i;

      const timeCell = sheet.getRange(row, timeCol);
      const messageCell = sheet.getRange(row, messageCol);

      timeCell.setNumberFormat("@"); // Plain text format
      timeCell.setValue(startTime);

      const message = `Your ${roles[j]} will be UTC ${startTime}-${endTime}
I will register you so you just need to show up on time. I will put you in before reset so check your schedule. If you can't show up, just don't do anything. If you cancel your schedule, everyone else will be messed up.`;

      messageCell.setValue(message);
    }
  }

  sheet.getRange("D4:E51").insertCheckboxes();
  sheet.getRange("J4:K51").insertCheckboxes();
  sheet.getRange("P4:Q51").insertCheckboxes();
}

function styleBordersAndBackgrounds(sheet) {
  const gray = "#BCBCBC";
  const borderRanges = [
    "A1:A52",   // Left border column
    "G1:G52",   // Between Construction & Research
    "M1:M52",   // Between Research & Troops
    "S1:S52",   // Right border column
    "A1:S1",    // Top row
    "A52:S52"   // Bottom row
  ];

  borderRanges.forEach(rangeStr => {
    const range = sheet.getRange(rangeStr);
    range.setBackground(gray);
  });
}

function setColumnWidths(sheet) {
  const pixelWidth = 44; 

  const columnsToResize = [
    "A", "G", "M", "S"
  ];

  columnsToResize.forEach(col => {
    const colIndex = columnLetterToIndex(col);
    sheet.setColumnWidth(colIndex, pixelWidth);
  });
}

// Helper function to convert column letter to index
function columnLetterToIndex(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column *= 26;
    column += letter.charCodeAt(i) - 64;
  }
  return column;
}
