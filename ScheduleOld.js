// Normalize time string by converting it to a string, trimming whitespace,
// and removing a leading zero from the hour (e.g., "00:00" ➝ "0:00")
function normalizeTime(timeStr) {
  if (!timeStr) return '';
  return String(timeStr).trim().replace(/^0/, ''); // Convert 00:00 ➝ 0:00
}

function scheduleBasedOnSort(sheet) {
  const names = sheet.getRange("B7:B54").getValues().flat();           // Player names
  const availabilities = sheet.getRange("K7:K54").getValues().flat(); // Cleaned availability
  const startTimes = sheet.getRange("L7:L54").getValues().flat().map(normalizeTime); // Time intervals
  const contributions = sheet.getRange("F7:F54").getValues().flat();  // Fire Crystals (or contribution metric)
  const targetColumn = 13; // Column M — Sorted Item schedule
  const assignedPlayers = new Set();

  // Build unique player pool and deduplicate
  const seenNames = new Set();
  const playerPool = [];

  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (!name || seenNames.has(name)) continue;

    seenNames.add(name);
    const slots = availabilities[i] ? availabilities[i].split(", ").map(normalizeTime) : [];
    const contribution = Number(contributions[i]) || 0;
    playerPool.push({ name, index: i, slots, contribution });
  }

  // Sort players by:
  // 1. Fewest available time slots
  // 2. Highest contribution (tie-breaker)
  playerPool.sort((a, b) => {
    if (a.slots.length !== b.slots.length) return a.slots.length - b.slots.length;
    return b.contribution - a.contribution;
  });

  // Build availability mapping for each time slot
  const timeAvailability = startTimes.map((time, i) => {
    const availablePlayers = playerPool.filter(p => p.slots.includes(time) && !assignedPlayers.has(p.name));
    return { time, row: i + 7, availablePlayers };
  });

  // Sort time slots by fewest available players
  timeAvailability.sort((a, b) => a.availablePlayers.length - b.availablePlayers.length);

  // Assign players to time slots
  for (const slot of timeAvailability) {
    const player = slot.availablePlayers.find(p => !assignedPlayers.has(p.name));

    if (player) {
      sheet.getRange(slot.row, targetColumn).setValue(player.name);
      assignedPlayers.add(player.name);
    } else {
      sheet.getRange(slot.row, targetColumn).setBackground("#FFAAA5"); // Unfilled slot — mark red
    }
  }

  // Highlight unused players
  for (const player of playerPool) {
    if (!assignedPlayers.has(player.name)) {
      sheet.getRange(player.index + 7, 2).setBackground("#FFAAA5"); // Column B
    }
  }

  Logger.log(`Assigned ${assignedPlayers.size} of ${playerPool.length} players.`);
}

// Schedule players with the *least* availability first (based on fewest options)
function scheduleBasedOnLimitedAvailability(sheet) {
  const names = sheet.getRange("B7:B54").getValues().flat(); // Player names
  const availabilities = sheet.getRange("K7:K54").getValues().flat(); // Cleaned availability
  const startTimes = sheet.getRange("L7:L54").getValues().flat().map(normalizeTime); // Time intervals
  const targetColumn = 14; // Column N — Limited Availability schedule

  const assignedSlots = new Set();
  const assignedPlayers = new Set();

  // Build player pool with unique names
  const seenNames = new Set();
  const playerPool = [];

  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (!name || seenNames.has(name)) continue;

    seenNames.add(name);
    const slots = availabilities[i] ? availabilities[i].split(", ").map(normalizeTime) : [];
    playerPool.push({ name, index: i, slots });
  }

  // Sort players by fewest available time slots
  playerPool.sort((a, b) => a.slots.length - b.slots.length);

  // Build availability mapping for each time slot
  const timeAvailability = startTimes.map((time, i) => {
    const availablePlayers = playerPool.filter(p => p.slots.includes(time) && !assignedPlayers.has(p.name));
    return { time, row: i + 7, availablePlayers };
  });

  // Sort time slots by fewest available players
  timeAvailability.sort((a, b) => a.availablePlayers.length - b.availablePlayers.length);

  for (const slot of timeAvailability) {
    const player = slot.availablePlayers.find(p => !assignedPlayers.has(p.name));

    if (player) {
      sheet.getRange(slot.row, targetColumn).setValue(player.name);
      assignedPlayers.add(player.name);
    }
  }

  // Highlight unassigned players in column B
  for (const player of playerPool) {
    if (!assignedPlayers.has(player.name)) {
      sheet.getRange(player.index + 7, 2).setBackground("#FFFFBA");
    }
  }

  // Highlight blank slots in column N
  const targetRange = sheet.getRange(7, targetColumn, startTimes.length);
  const values = targetRange.getValues();

  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      sheet.getRange(7 + i, targetColumn).setBackground("#FFFFBA");
    }
  }

  Logger.log(`Assigned ${assignedPlayers.size} of ${playerPool.length} players.`);
}

// Assign slots to players rather than players to slots
function scheduleAllSlots(sheet) {
  const names = sheet.getRange("B7:B54").getValues().flat();
  const availabilities = sheet.getRange("K7:K54").getValues().flat();
  const startTimes = sheet.getRange("L7:L54").getValues().flat().map(normalizeTime);
  const targetColumn = 15; // Column O
  const assignedNames = new Set(); // Track unique player names

  // Build player pool (skip duplicates by only keeping first occurrence)
  const seenNames = new Set();
  const playerPool = [];

  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (!name || seenNames.has(name)) continue;

    seenNames.add(name);
    const slots = availabilities[i] ? availabilities[i].split(", ").map(normalizeTime) : [];
    playerPool.push({ name, index: i, slots });
  }

  // Build time slot availability
  const timeAvailability = startTimes.map((time, i) => {
    const availablePlayers = playerPool.filter(
      p => p.slots.includes(time) && !assignedNames.has(p.name)
    );
    return { time, row: i + 7, availablePlayers };
  });

  // Sort time slots by fewest available players
  timeAvailability.sort((a, b) => a.availablePlayers.length - b.availablePlayers.length);

  // Schedule players
  for (const slot of timeAvailability) {
    const player = slot.availablePlayers.find(p => !assignedNames.has(p.name));
    if (player) {
      sheet.getRange(slot.row, targetColumn).setValue(player.name);
      assignedNames.add(player.name);
    }
  }

  // Highlight unused (first instance only) in green in Column B
  for (const player of playerPool) {
    if (!assignedNames.has(player.name)) {
      sheet.getRange(player.index + 7, 2).setBackground("#BAFFC9");
    }
  }

  // Highlight blank slots in column N (target column)
  const targetRange = sheet.getRange(7, targetColumn, startTimes.length);
  const values = targetRange.getValues();

  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      sheet.getRange(7 + i, targetColumn).setBackground("#BAFFC9");
    }
  }
}

function scheduleBasedOnSpeedups(sheet) {
  const START_ROW = 7;
  const LAST_ROW = sheet.getLastRow(); // Dynamically find last row
  const RANGE_SIZE = LAST_ROW - START_ROW + 1;

  const names = sheet.getRange(START_ROW, 2, RANGE_SIZE).getValues().flat();           // Column B
  const availabilities = sheet.getRange(START_ROW, 11, RANGE_SIZE).getValues().flat(); // Column K
  const startTimes = sheet.getRange(START_ROW, 12, RANGE_SIZE).getValues().flat().map(normalizeTime); // Column L
  const speedups = sheet.getRange(START_ROW, 7, RANGE_SIZE).getValues().flat();        // Column G

  const targetColumn = 16; // Column P
  const assignedPlayers = new Set();
  const seenNames = new Set();
  const playerPool = [];

  // Deduplicate and collect full player list
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (!name || seenNames.has(name)) continue;

    seenNames.add(name);
    const slots = availabilities[i] ? availabilities[i].split(", ").map(normalizeTime) : [];
    const speedup = parseFloat(speedups[i]) || 0;
    playerPool.push({ name, index: i, slots, speedup });
  }

  // Sort by fewest available slots first, then highest speedup
  playerPool.sort((a, b) => {
    if (a.slots.length !== b.slots.length) return a.slots.length - b.slots.length;
    return b.speedup - a.speedup;
  });

  // Take only top 48 players
  const topPlayers = playerPool.slice(0, 48);

  // Build availability mapping per time slot
  const timeAvailability = startTimes.map((time, i) => {
    const availablePlayers = topPlayers.filter(p => p.slots.includes(time) && !assignedPlayers.has(p.name));
    return { time, row: i + START_ROW, availablePlayers };
  });

  // Sort time slots by fewest available players
  timeAvailability.sort((a, b) => a.availablePlayers.length - b.availablePlayers.length);

  // Assign players to slots
  for (const slot of timeAvailability) {
    const player = slot.availablePlayers.find(p => !assignedPlayers.has(p.name));
    if (player) {
      sheet.getRange(slot.row, targetColumn).setValue(player.name);
      assignedPlayers.add(player.name);
    }
  }

  // Highlight unfilled slots in column P
  const scheduledValues = sheet.getRange(START_ROW, targetColumn, startTimes.length).getValues().flat();
  scheduledValues.forEach((val, i) => {
    if (!val) {
      sheet.getRange(START_ROW + i, targetColumn).setBackground("#E7AFED"); // Purple
    }
  });

  // Highlight unused players from top 48
  for (const player of topPlayers) {
    if (!assignedPlayers.has(player.name)) {
      sheet.getRange(player.index + START_ROW, 2).setBackground("#E7AFED"); // Column B
    }
  }

  Logger.log(`Assigned ${assignedPlayers.size} of ${topPlayers.length} top speedup players.`);
}

function highlightDuplicateNames(sheet) {
  const nameRange = sheet.getRange("B7:B107");
  const names = nameRange.getValues().flat();
  const nameCounts = {};

  // Count occurrences of each name
  names.forEach(name => {
    if (!name) return;
    nameCounts[name] = (nameCounts[name] || 0) + 1;
  });

  // Highlight rows where name occurs more than once
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (name && nameCounts[name] > 1) {
      // Highlight cells A–G for the duplicate row
      sheet.getRange(i + 7, 1, 1, 10).setBackground("grey");
    }
  }
}

// Main function to clear old results and run all 3 strategies
function Schedule() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Clear previous assignments and highlights
  sheet.getRange("M7:P54").clearContent();      // Columns M, N, O, P
  sheet.getRange("B7:B54").setBackground(null); // Reset name highlights
  sheet.getRange("M7:P54").setBackground(null); // Reset start time highlights

  // Run all three scheduling strategies
  scheduleBasedOnSort(sheet);
  scheduleBasedOnLimitedAvailability(sheet);
  scheduleAllSlots(sheet);
  scheduleBasedOnSpeedups(sheet)
  highlightDuplicateNames(sheet);
}