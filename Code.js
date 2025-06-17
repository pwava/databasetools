// --- Constants for Sheet Names and Column Indices ---
// 'Tools' Sheet (This Spreadsheet)
const SETTINGS_SHEET_NAME = 'Settings';
const UPDATE_TRACKER_SHEET_NAME = 'Update Attendance Tracker';
const EVENT_ATTENDANCE_TAB_NAME = 'Event Attendance';

// 'Settings' Tab Columns (1-based index)
const SETTINGS_COL_COMMUNITY_ID = 2; // Column B
const SETTINGS_COL_ATTENDANCE_STATS_URL = 3; // Column C
const SETTINGS_COL_DIRECTORY_URL = 4; // Column D

// 'Update Attendance Tracker' Tab Columns (1-based index)
const UAT_COL_ID = 1; // Column A
const UAT_COL_FULL_NAME = 2; // Column B
const UAT_COL_LAST_NAME = 3; // Column C
const UAT_COL_FIRST_NAME = 4; // Column D
const UAT_COL_ACTIVITY_LEVEL = 5; // Column E
const UAT_HEADER_ROW = 6; // Row 6 for headers
const UAT_START_DATA_ROW = 7; // Data starts from Row 7
const UAT_CELL_COMMUNITY_ID = 'B4';

// Community Sheets: 'Directory' Tab Columns (1-based index)
const DIR_COL_PERSON_ID = 1; // Column A
const DIR_COL_LAST_NAME = 3; // Column C
const DIR_COL_FIRST_NAME = 4; // Column D
const DIRECTORY_TAB_NAME = 'Directory';

// Community Sheets: 'Attendance Stats' Tab Columns (1-based index)
const STATS_COL_PERSON_ID = 1; // Column A
const STATS_COL_FIRST_NAME = 3; // Column C
const STATS_COL_LAST_NAME = 4; // Column D
const STATS_COL_QUARTER_EVENTS = 5; // Column E for "Events This Quarter"
const STATS_COL_ACTIVITY_SCORE = 12; // Column L (for the numeric activity score)
const ATTENDANCE_STATS_TAB_NAME = 'Attendance Stats';
// --- End of Constants Section ---

// --- UI and Menu ---

/**
 * Runs when the spreadsheet is opened. Adds a custom menu with icons (emojis).
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Community Tools')
    .addItem('📥 Load Data', 'loadData')
    .addItem('📋 Get Names From Attendance Stats', 'getNamesFromAttendanceStats')
    .addItem('📂 Get Names From Directory', 'getNamesFromDirectory')
    .addSeparator()
    .addItem('✅ Update Activity Level', 'updateActivityLevels')
    .addSeparator()
    .addItem('🗑️ Clear Names', 'clearNamesAndCommunityID')
    .addItem('↩️ Reset Activity Level', 'resetActivityLevelValues')
    .addToUi();
}

// --- Helper Function to Get Community URLs ---

/**
 * Gets the Directory and Attendance Stats URLs for the selected community.
 * @param {string} communityId The ID of the community selected in UAT_CELL_COMMUNITY_ID.
 * @return {object|null} An object with { directoryUrl: '...', statsUrl: '...' } or null if not found.
 */
function getCommunityUrls(communityId) {
  if (!communityId) {
    SpreadsheetApp.getUi().alert('Error', 'Please select a Community ID from cell ' + UAT_CELL_COMMUNITY_ID + '.', SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    SpreadsheetApp.getUi().alert('Error', 'The "Settings" tab could not be found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }

  const settingsData = settingsSheet.getDataRange().getValues();
  for (let i = 0; i < settingsData.length; i++) {
    if (settingsData[i][SETTINGS_COL_COMMUNITY_ID - 1] == communityId) {
      return {
        directoryUrl: settingsData[i][SETTINGS_COL_DIRECTORY_URL - 1],
        statsUrl: settingsData[i][SETTINGS_COL_ATTENDANCE_STATS_URL - 1]
      };
    }
  }
  SpreadsheetApp.getUi().alert('Error', 'Community ID "' + communityId + '" not found in the "Settings" tab.', SpreadsheetApp.getUi().ButtonSet.OK);
  return null;
}


// --- Main Functions for Menu Items ---

/**
 * Loads Person IDs and Full Names from the community's 'Directory' sheet
 * based on Last Name and First Name entered in the 'Update Attendance Tracker' tab.
 */
function loadData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const uatSheet = ss.getSheetByName(UPDATE_TRACKER_SHEET_NAME);

  if (!uatSheet) {
    ui.alert('Error', 'Sheet "' + UPDATE_TRACKER_SHEET_NAME + '" not found.', ui.ButtonSet.OK);
    return;
  }

  const selectedCommunityId = uatSheet.getRange(UAT_CELL_COMMUNITY_ID).getValue();
  if (!selectedCommunityId) {
    ui.alert('Error', 'Please select a Community ID in cell ' + UAT_CELL_COMMUNITY_ID + '.', ui.ButtonSet.OK);
    return;
  }

  const urls = getCommunityUrls(selectedCommunityId);
  if (!urls || !urls.directoryUrl) {
    return;
  }

  let directorySpreadsheet;
  try {
    directorySpreadsheet = SpreadsheetApp.openByUrl(urls.directoryUrl);
  } catch (e) {
    ui.alert('Error', 'Could not open the Directory spreadsheet. Please check the URL in Settings for Community ID "' + selectedCommunityId + '". Error: ' + e.message, ui.ButtonSet.OK);
    return;
  }

  const directorySheet = directorySpreadsheet.getSheetByName(DIRECTORY_TAB_NAME);
  if (!directorySheet) {
    ui.alert('Error', 'Tab "' + DIRECTORY_TAB_NAME + '" not found in the Directory spreadsheet for ' + selectedCommunityId + '.', ui.ButtonSet.OK);
    return;
  }

  const directoryData = directorySheet.getDataRange().getValues();
  let lastRowWithInput = UAT_START_DATA_ROW - 1;
  const lastNameColRange = uatSheet.getRange(UAT_START_DATA_ROW, UAT_COL_LAST_NAME, uatSheet.getMaxRows() - UAT_START_DATA_ROW + 1, 1);
  const firstNameColRange = uatSheet.getRange(UAT_START_DATA_ROW, UAT_COL_FIRST_NAME, uatSheet.getMaxRows() - UAT_START_DATA_ROW + 1, 1);
  const lastNameColValues = lastNameColRange.getValues();
  const firstNameColValues = firstNameColRange.getValues();

  for (let i = lastNameColValues.length - 1; i >= 0; i--) {
    if ((lastNameColValues[i][0] && String(lastNameColValues[i][0]).trim() !== "") ||
      (firstNameColValues[i][0] && String(firstNameColValues[i][0]).trim() !== "")) {
      lastRowWithInput = UAT_START_DATA_ROW + i;
      break;
    }
  }

  if (lastRowWithInput < UAT_START_DATA_ROW) {
    ui.alert('Info', 'No names entered in columns C (Last Name) or D (First Name) of "Update Attendance Tracker" tab to load data for.', ui.ButtonSet.OK);
    return;
  }

  const numDataRowsToProcess = lastRowWithInput - UAT_START_DATA_ROW + 1;
  const uatRangeToProcess = uatSheet.getRange(UAT_START_DATA_ROW, UAT_COL_ID, numDataRowsToProcess, UAT_COL_FIRST_NAME);
  const uatValuesToUpdate = uatRangeToProcess.getValues();

  let namesProcessed = 0;
  let namesFound = 0;
  let namesNotFoundList = [];

  for (let i = 0; i < uatValuesToUpdate.length; i++) {
    const lastNameToSearch = uatValuesToUpdate[i][UAT_COL_LAST_NAME - UAT_COL_ID];
    const firstNameToSearch = uatValuesToUpdate[i][UAT_COL_FIRST_NAME - UAT_COL_ID];

    if (lastNameToSearch && String(lastNameToSearch).trim() !== "" &&
      firstNameToSearch && String(firstNameToSearch).trim() !== "") {
      namesProcessed++;
      let foundMatchInDirectory = false;
      for (let j = 0; j < directoryData.length; j++) {
        const dirLastName = directoryData[j][DIR_COL_LAST_NAME - 1];
        const dirFirstName = directoryData[j][DIR_COL_FIRST_NAME - 1];

        if (String(dirLastName).trim().toLowerCase() === String(lastNameToSearch).trim().toLowerCase() &&
          String(dirFirstName).trim().toLowerCase() === String(firstNameToSearch).trim().toLowerCase()) {

          const personId = directoryData[j][DIR_COL_PERSON_ID - 1];
          const fullName = String(dirFirstName).trim() + " " + String(dirLastName).trim();

          uatValuesToUpdate[i][UAT_COL_ID - UAT_COL_ID] = personId;
          uatValuesToUpdate[i][UAT_COL_FULL_NAME - UAT_COL_ID] = fullName;
          namesFound++;
          foundMatchInDirectory = true;
          break;
        }
      }
      if (!foundMatchInDirectory) {
        namesNotFoundList.push(String(firstNameToSearch).trim() + " " + String(lastNameToSearch).trim());
        uatValuesToUpdate[i][UAT_COL_ID - UAT_COL_ID] = null;
        uatValuesToUpdate[i][UAT_COL_FULL_NAME - UAT_COL_ID] = null;
        uatValuesToUpdate[i][UAT_COL_LAST_NAME - UAT_COL_ID] = null;
        uatValuesToUpdate[i][UAT_COL_FIRST_NAME - UAT_COL_ID] = null;
      }
    }
  }

  if (numDataRowsToProcess > 0) {
    uatRangeToProcess.setValues(uatValuesToUpdate);
  }

  let message = "Load Data Complete.\n";
  if (namesProcessed > 0) {
    message += "Attempted to process " + namesProcessed + " names (where both First and Last Name were provided).\n";
    message += "Found and data loaded for: " + namesFound + ".\n";
    if (namesNotFoundList.length > 0) {
      message += "Not Found in Directory (and cleared from sheet): " + namesNotFoundList.length + ".\n";
      message += "Details of not found (and cleared): " + namesNotFoundList.join(", ") + ".";
    } else if (namesFound === namesProcessed) {
      message += "All processed names were found in the directory.";
    }
  } else {
    message = 'No names with both First and Last Name were found to process in the "Update Attendance Tracker" tab.';
  }
  ui.alert('Load Data Results', message, ui.ButtonSet.OK);
}


/**
 * -- MODIFIED VERSION --
 * Calculates activity scores internally and logs corresponding "event attendances".
 * This version DOES NOT write or update the score in the 'Attendance Stats' sheet.
 */
function updateActivityLevels() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = ss.getSpreadsheetTimeZone();
  const uatSheet = ss.getSheetByName(UPDATE_TRACKER_SHEET_NAME);

  if (!uatSheet) {
    ui.alert('Error', `Sheet "${UPDATE_TRACKER_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }

  const selectedCommunityId = uatSheet.getRange(UAT_CELL_COMMUNITY_ID).getValue();
  if (!selectedCommunityId) {
    ui.alert('Error', `Please select a Community ID in cell ${UAT_CELL_COMMUNITY_ID}.`, ui.ButtonSet.OK);
    return;
  }

  const urls = getCommunityUrls(selectedCommunityId);
  if (!urls || !urls.statsUrl) {
    if (urls && !urls.statsUrl) {
      ui.alert('Error', `URL for 'Attendance Stats' of community "${selectedCommunityId}" is missing in Settings.`, ui.ButtonSet.OK);
    }
    return;
  }

  let statsSpreadsheets;
  try {
    statsSpreadsheets = SpreadsheetApp.openByUrl(urls.statsUrl);
  } catch (e) {
    ui.alert('Error', `Could not open the Attendance Stats spreadsheet for "${selectedCommunityId}". Error: ${e.message}`, ui.ButtonSet.OK);
    return;
  }

  const statsSheet = statsSpreadsheets.getSheetByName(ATTENDANCE_STATS_TAB_NAME);
  if (!statsSheet) {
    ui.alert('Error', `Tab "${ATTENDANCE_STATS_TAB_NAME}" not found in the 'Attendance Stats' sheet for ${selectedCommunityId}.`, ui.ButtonSet.OK);
    return;
  }

  let eventAttendanceSheet = statsSpreadsheets.getSheetByName(EVENT_ATTENDANCE_TAB_NAME);
  if (!eventAttendanceSheet) {
    eventAttendanceSheet = statsSpreadsheets.insertSheet(EVENT_ATTENDANCE_TAB_NAME);
    eventAttendanceSheet.appendRow([
      "Person ID", "Full Name", "Event", "First Name", "Last Name",
      null, null, null, null, null,
      "Event Date",
      null, null,
      "Update Timestamp"
    ]);
    SpreadsheetApp.flush();
  }

  const statsData = statsSheet.getDataRange().getValues();
  const uatLastRow = uatSheet.getLastRow();
  if (uatLastRow < UAT_START_DATA_ROW) {
    ui.alert("Info", "No data rows to process in 'Update Attendance Tracker'.", ui.ButtonSet.OK);
    return;
  }
  const uatData = uatSheet.getRange(UAT_START_DATA_ROW, 1, uatLastRow - UAT_START_DATA_ROW + 1, UAT_COL_ACTIVITY_LEVEL).getValues();

  const today = new Date();
  const formattedExecutionDate = Utilities.formatDate(today, timeZone, "M/d/yyyy");

  let recordsToProcessCount = 0;
  let recordsSkippedOrFailed = [];
  let recordsMissingDetailsUAT = 0;
  let eventAttendancesLogged = 0;

  const TARGET_SUM_ACTIVE_MIN = 3;
  const TARGET_SUM_ACTIVE_MAX = 11;
  const TARGET_SUM_CORE_MIN = 12;
  const CORE_K_MIN_VALUE = 12;

  for (let i = 0; i < uatData.length; i++) {
    const personIdUAT = uatData[i][UAT_COL_ID - 1];
    const fullNameUAT = uatData[i][UAT_COL_FULL_NAME - 1];
    const lastNameUAT = uatData[i][UAT_COL_LAST_NAME - 1];
    const firstNameUAT = uatData[i][UAT_COL_FIRST_NAME - 1];
    const activityLevelUAT = uatData[i][UAT_COL_ACTIVITY_LEVEL - 1];

    if (personIdUAT && String(personIdUAT).trim() !== "" &&
      lastNameUAT && String(lastNameUAT).trim() !== "" &&
      firstNameUAT && String(firstNameUAT).trim() !== "" &&
      activityLevelUAT && String(activityLevelUAT).trim() !== "") {

      recordsToProcessCount++;
      let newCalculatedScoreK;

      let current_E_val = 0;
      let current_K_val = 0;

      for (let j = 0; j < statsData.length; j++) {
        if (String(statsData[j][STATS_COL_PERSON_ID - 1]).trim() == String(personIdUAT).trim() &&
          String(statsData[j][STATS_COL_FIRST_NAME - 1]).trim().toLowerCase() == String(firstNameUAT).trim().toLowerCase() &&
          String(statsData[j][STATS_COL_LAST_NAME - 1]).trim().toLowerCase() == String(lastNameUAT).trim().toLowerCase()) {

          current_E_val = parseInt(statsData[j][STATS_COL_QUARTER_EVENTS - 1], 10) || 0;
          current_K_val = parseInt(statsData[j][STATS_COL_ACTIVITY_SCORE - 1], 10) || 0;
          break;
        }
      }

      const uatActivityLevelTrimmed = String(activityLevelUAT).trim().toLowerCase();

      switch (uatActivityLevelTrimmed) {
        case 'inactive':
          newCalculatedScoreK = 1;
          break;
        case 'active':
          let points_to_add_active = 3;
          let tentative_K_active = current_K_val + points_to_add_active;
          let target_sum_active = current_E_val + tentative_K_active;
          target_sum_active = Math.max(TARGET_SUM_ACTIVE_MIN, target_sum_active);
          target_sum_active = Math.min(TARGET_SUM_ACTIVE_MAX, target_sum_active);
          newCalculatedScoreK = target_sum_active - current_E_val;
          newCalculatedScoreK = Math.max(0, newCalculatedScoreK);
          break;
        case 'core':
          newCalculatedScoreK = CORE_K_MIN_VALUE;
          break;
        default:
          console.log(`Unknown activity level "${activityLevelUAT}" for ${firstNameUAT} ${lastNameUAT}.`);
          recordsSkippedOrFailed.push(`${firstNameUAT} ${lastNameUAT} (ID: ${personIdUAT}) - Unknown Level`);
          continue;
      }

      if (newCalculatedScoreK > 0) {
        const eventDates = generateRecentDistinctDates(newCalculatedScoreK, today, timeZone);
        const rowsForEventAttendance = [];
        let eventCounter = 1;
        for (const eventDate of eventDates) {
          const newEventAttendanceRow = new Array(14).fill(null);
          newEventAttendanceRow[0] = personIdUAT;
          newEventAttendanceRow[1] = fullNameUAT;
          newEventAttendanceRow[2] = `BASELINE ADJUSTMENT ${eventCounter}`;
          newEventAttendanceRow[4] = firstNameUAT;
          newEventAttendanceRow[5] = lastNameUAT;
          newEventAttendanceRow[10] = eventDate;
          newEventAttendanceRow[13] = formattedExecutionDate;

          rowsForEventAttendance.push(newEventAttendanceRow);
          eventCounter++;
        }
        if (rowsForEventAttendance.length > 0) {
          eventAttendanceSheet.getRange(
            eventAttendanceSheet.getLastRow() + 1, 1,
            rowsForEventAttendance.length,
            rowsForEventAttendance[0].length
          ).setValues(rowsForEventAttendance);
          eventAttendancesLogged += rowsForEventAttendance.length;
        }
      }
    } else if (activityLevelUAT && String(activityLevelUAT).trim() !== "" &&
      !(personIdUAT && String(personIdUAT).trim() !== "" &&
        lastNameUAT && String(lastNameUAT).trim() !== "" &&
        firstNameUAT && String(firstNameUAT).trim() !== "")) {
      recordsMissingDetailsUAT++;
    }
  }

  let message = "Log Event Attendance - Results:\n";
  if (recordsToProcessCount > 0) {
    message += `Attempted to process ${recordsToProcessCount} records from the Tools sheet.\n`;
    message += `- Logged to 'Event Attendance' tab: ${eventAttendancesLogged} entries.\n`;
    message += `(The 'Attendance Stats' sheet was not modified).\n`;
    if (recordsSkippedOrFailed.length > 0) {
      message += `Skipped or failed (e.g., unknown activity level): ${recordsSkippedOrFailed.length} records.\n   Details: ${recordsSkippedOrFailed.join("; ")}\n`;
    }
  } else {
    message += "No records in 'Update Attendance Tracker' had sufficient details (ID, Name, Activity Level) to process.\n";
  }
  if (recordsMissingDetailsUAT > 0) {
    message += `${recordsMissingDetailsUAT} record(s) had an activity level set but were missing ID, First Name, or Last Name, and were skipped.\n`;
  }

  ui.alert('Processing Complete', message, ui.ButtonSet.OK);
}


function clearNamesAndCommunityID() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const uatSheet = ss.getSheetByName(UPDATE_TRACKER_SHEET_NAME);

  if (!uatSheet) {
    ui.alert('Error', 'Sheet "' + UPDATE_TRACKER_SHEET_NAME + '" not found.', ui.ButtonSet.OK);
    return;
  }

  uatSheet.getRange(UAT_CELL_COMMUNITY_ID).clearContent();
  const lastRow = uatSheet.getLastRow();
  if (lastRow >= UAT_START_DATA_ROW) {
    uatSheet.getRange(UAT_START_DATA_ROW, UAT_COL_ID, lastRow - UAT_START_DATA_ROW + 1, UAT_COL_ACTIVITY_LEVEL).clearContent();
  }
  ui.alert('Form Cleared', 'Community ID and all names/data have been cleared from the "Update Attendance Tracker" tab.', ui.ButtonSet.OK);
}


function resetActivityLevelValues() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const uatSheet = ss.getSheetByName(UPDATE_TRACKER_SHEET_NAME);

  if (!uatSheet) {
    ui.alert('Error', `Sheet "${UPDATE_TRACKER_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }

  let uatRowsCleared = 0;
  const uatLastRow = uatSheet.getLastRow();
  if (uatLastRow >= UAT_START_DATA_ROW) {
    const activityLevelRangeUAT = uatSheet.getRange(UAT_START_DATA_ROW, UAT_COL_ACTIVITY_LEVEL, uatLastRow - UAT_START_DATA_ROW + 1, 1);
    activityLevelRangeUAT.getValues().forEach(row => {
      if (row[0] !== "") uatRowsCleared++;
    });
    activityLevelRangeUAT.clearContent();
  }

  const selectedCommunityId = uatSheet.getRange(UAT_CELL_COMMUNITY_ID).getValue();
  if (!selectedCommunityId) {
    if (uatRowsCleared > 0) {
      ui.alert('Activity Levels Reset in Tools Sheet',
        `The "Activity Level" column in the "Update Attendance Tracker" tab has been cleared for ${uatRowsCleared} row(s).\n\n` +
        `No Community ID was selected in B4, so no changes were made to any 'Attendance Stats' sheet.`,
        ui.ButtonSet.OK);
    } else {
      ui.alert('Info', 'No Community ID selected and no activity levels to clear in the Tools sheet.', ui.ButtonSet.OK);
    }
    return;
  }

  const urls = getCommunityUrls(selectedCommunityId);
  if (!urls || !urls.statsUrl) {
    ui.alert('Error', `Could not get URL for 'Attendance Stats' sheet of community "${selectedCommunityId}". Please check Settings.`, ui.ButtonSet.OK);
    return;
  }

  let statsSpreadsheets;
  try {
    statsSpreadsheets = SpreadsheetApp.openByUrl(urls.statsUrl);
  } catch (e) {
    ui.alert('Error', `Could not open the Attendance Stats spreadsheet for community "${selectedCommunityId}". Please check the URL. Error: ${e.message}`, ui.ButtonSet.OK);
    return;
  }

  const statsSheet = statsSpreadsheets.getSheetByName(ATTENDANCE_STATS_TAB_NAME);
  if (!statsSheet) {
    ui.alert('Error', `Tab "${ATTENDANCE_STATS_TAB_NAME}" not found in the Attendance Stats spreadsheet for ${selectedCommunityId}.`, ui.ButtonSet.OK);
    return;
  }

  const uatDataForMatching = uatSheet.getRange(UAT_START_DATA_ROW, UAT_COL_ID, uatLastRow - UAT_START_DATA_ROW + 1, UAT_COL_FIRST_NAME).getValues();
  const statsData = statsSheet.getDataRange().getValues();

  let statsScoresResetCount = 0;
  let uatPeopleProcessedForStatsReset = 0;
  let notFoundInStatsSheet = [];

  for (let i = 0; i < uatDataForMatching.length; i++) {
    const personIdUAT = uatDataForMatching[i][UAT_COL_ID - 1];
    const lastNameUAT = uatDataForMatching[i][UAT_COL_LAST_NAME - 1];
    const firstNameUAT = uatDataForMatching[i][UAT_COL_FIRST_NAME - 1];

    if (personIdUAT && String(personIdUAT).trim() !== "" &&
      lastNameUAT && String(lastNameUAT).trim() !== "" &&
      firstNameUAT && String(firstNameUAT).trim() !== "") {

      uatPeopleProcessedForStatsReset++;
      let foundMatchInStats = false;
      for (let j = 0; j < statsData.length; j++) {
        const statsSheetRowIndex = j + 1;
        const personIdStats = statsData[j][STATS_COL_PERSON_ID - 1];
        const firstNameStats = statsData[j][STATS_COL_FIRST_NAME - 1];
        const lastNameStats = statsData[j][STATS_COL_LAST_NAME - 1];

        if (String(personIdStats).trim() == String(personIdUAT).trim() &&
          String(firstNameStats).trim().toLowerCase() == String(firstNameUAT).trim().toLowerCase() &&
          String(lastNameStats).trim().toLowerCase() == String(lastNameUAT).trim().toLowerCase()) {

          const currentStatsScoreCell = statsSheet.getRange(statsSheetRowIndex, STATS_COL_ACTIVITY_SCORE);
          const currentStatsScoreValue = currentStatsScoreCell.getValue();
          if (currentStatsScoreValue !== 0 && currentStatsScoreValue !== "") {
            currentStatsScoreCell.setValue(0);
            statsScoresResetCount++;
          } else if (currentStatsScoreValue === 0 || currentStatsScoreValue === "") {
            statsScoresResetCount++;
          }
          foundMatchInStats = true;
          break;
        }
      }
      if (!foundMatchInStats) {
        notFoundInStatsSheet.push(`${firstNameUAT} ${lastNameUAT} (ID: ${personIdUAT})`);
      }
    }
  }

  let message = "";
  if (uatRowsCleared > 0) {
    message += `${uatRowsCleared} entr(y/ies) in "Activity Level" column (Tools sheet) have been cleared.\n\n`;
  } else {
    message += `No activity levels to clear in the "Tools" sheet's "Activity Level" column.\n\n`;
  }

  if (uatPeopleProcessedForStatsReset > 0) {
    message += `For community "${selectedCommunityId}":\n`;
    message += `- Attempted to reset scores in 'Attendance Stats' (Column K) for ${uatPeopleProcessedForStatsReset} people listed in the Tools sheet.\n`;
    message += `- Scores reset to 0 (or confirmed as already 0/blank) for: ${statsScoresResetCount} people.\n`;
    if (notFoundInStatsSheet.length > 0) {
      message += `- Not found in 'Attendance Stats' sheet (or details mismatched): ${notFoundInStatsSheet.length} people.\n   Details: ${notFoundInStatsSheet.join("; ")}\n`;
    }
  } else if (selectedCommunityId) {
    message += `No people with full details (ID, First & Last Name) were listed in the Tools sheet to process for resetting scores in 'Attendance Stats' for community "${selectedCommunityId}".\n`;
  }

  ui.alert('Reset Activity Level - Results', message, ui.ButtonSet.OK);
}


function fetchAndPopulateNames(sourceType) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const uatSheet = ss.getSheetByName(UPDATE_TRACKER_SHEET_NAME);

  if (!uatSheet) {
    ui.alert('Error', `Sheet "${UPDATE_TRACKER_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }

  const selectedCommunityId = uatSheet.getRange(UAT_CELL_COMMUNITY_ID).getValue();
  if (!selectedCommunityId) {
    ui.alert('Error', `Please select a Community ID in cell ${UAT_CELL_COMMUNITY_ID}.`, ui.ButtonSet.OK);
    return;
  }

  const communityUrls = getCommunityUrls(selectedCommunityId);
  if (!communityUrls) return;

  let targetUrl, targetTabName, idCol, firstNameCol, lastNameCol, sourceSheetFriendlyName;

  if (sourceType === 'STATS') {
    targetUrl = communityUrls.statsUrl;
    targetTabName = ATTENDANCE_STATS_TAB_NAME;
    idCol = STATS_COL_PERSON_ID;
    firstNameCol = STATS_COL_FIRST_NAME;
    lastNameCol = STATS_COL_LAST_NAME;
    sourceSheetFriendlyName = "Attendance Stats";
    if (!targetUrl) {
      ui.alert('Error', `URL for 'Attendance Stats' not found in Settings for Community ID "${selectedCommunityId}".`, ui.ButtonSet.OK);
      return;
    }
  } else if (sourceType === 'DIRECTORY') {
    targetUrl = communityUrls.directoryUrl;
    targetTabName = DIRECTORY_TAB_NAME;
    idCol = DIR_COL_PERSON_ID;
    firstNameCol = DIR_COL_FIRST_NAME;
    lastNameCol = DIR_COL_LAST_NAME;
    sourceSheetFriendlyName = "Directory";
    if (!targetUrl) {
      ui.alert('Error', `URL for 'Directory' not found in Settings for Community ID "${selectedCommunityId}".`, ui.ButtonSet.OK);
      return;
    }
  } else {
    ui.alert('Error', 'Invalid source type specified for fetching names.', ui.ButtonSet.OK);
    return;
  }

  let sourceSpreadsheets;
  try {
    sourceSpreadsheets = SpreadsheetApp.openByUrl(targetUrl);
  } catch (e) {
    ui.alert('Error', `Could not open the ${sourceSheetFriendlyName} spreadsheet. Please check the URL in Settings. Error: ${e.message}`, ui.ButtonSet.OK);
    return;
  }

  const sourceSheet = sourceSpreadsheets.getSheetByName(targetTabName);
  if (!sourceSheet) {
    ui.alert('Error', `Tab "${targetTabName}" not found in the ${sourceSheetFriendlyName} spreadsheet for ${selectedCommunityId}.`, ui.ButtonSet.OK);
    return;
  }

  const sourceData = sourceSheet.getDataRange().getValues();
  const outputData = [];

  if (sourceData.length < 2) {
    ui.alert('Info', `No data (or only a header row) found in the "${targetTabName}" tab of the ${sourceSheetFriendlyName} sheet for ${selectedCommunityId}.`, ui.ButtonSet.OK);
    const lastRowUAT = uatSheet.getLastRow();
    if (lastRowUAT >= UAT_START_DATA_ROW) {
      uatSheet.getRange(UAT_START_DATA_ROW, UAT_COL_ID, lastRowUAT - UAT_START_DATA_ROW + 1, UAT_COL_ACTIVITY_LEVEL).clearContent();
    }
    return;
  }

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    if (row.length < Math.max(idCol, firstNameCol, lastNameCol)) {
      Logger.log(`Skipping row ${i+1} in ${sourceSheetFriendlyName} - ${targetTabName} due to insufficient columns.`);
      continue;
    }
    const personId = row[idCol - 1];
    const firstName = row[firstNameCol - 1];
    const lastName = row[lastNameCol - 1];
    if (personId && String(personId).trim() !== "" &&
      firstName && String(firstName).trim() !== "" &&
      lastName && String(lastName).trim() !== "") {
      const fullName = String(firstName).trim() + " " + String(lastName).trim();
      outputData.push([
        personId,
        fullName,
        lastName,
        firstName,
        null
      ]);
    }
  }

  const lastRowUAT = uatSheet.getLastRow();
  if (lastRowUAT >= UAT_START_DATA_ROW) {
    uatSheet.getRange(UAT_START_DATA_ROW, UAT_COL_ID, lastRowUAT - UAT_START_DATA_ROW + 1, UAT_COL_ACTIVITY_LEVEL).clearContent();
  }

  if (outputData.length > 0) {
    uatSheet.getRange(UAT_START_DATA_ROW, UAT_COL_ID, outputData.length, outputData[0].length).setValues(outputData);
    ui.alert('Success', `Fetched ${outputData.length} names from "${targetTabName}" in the ${sourceSheetFriendlyName} sheet (headers skipped) and populated them into "Update Attendance Tracker".`, ui.ButtonSet.OK);
  } else {
    ui.alert('Info', `No valid names (with ID, First Name, and Last Name) found to fetch from "${targetTabName}" in the ${sourceSheetFriendlyName} sheet after skipping header. "Update Attendance Tracker" has been cleared.`, ui.ButtonSet.OK);
  }
}

function getNamesFromAttendanceStats() {
  fetchAndPopulateNames('STATS');
}

function getNamesFromDirectory() {
  fetchAndPopulateNames('DIRECTORY');
}


function generateRecentDistinctDates(N, referenceDate, timeZone) {
  if (N <= 0) return [];
  const datesArray = [];
  const year = referenceDate.getFullYear();
  const month = referenceDate.getMonth();
  let quarterStartMonth;
  if (month < 3) quarterStartMonth = 0;
  else if (month < 6) quarterStartMonth = 3;
  else if (month < 9) quarterStartMonth = 6;
  else quarterStartMonth = 9;
  const quarterStartDate = new Date(year, quarterStartMonth, 1);
  for (let i = 0; i < N; i++) {
    let eventDate = new Date(referenceDate.getTime());
    eventDate.setDate(referenceDate.getDate() - i);
    if (eventDate.getTime() < quarterStartDate.getTime()) {
      break;
    }
    datesArray.push(Utilities.formatDate(eventDate, timeZone, "M/d/yyyy"));
  }
  return datesArray;
}
