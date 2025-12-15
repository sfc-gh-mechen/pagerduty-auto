// ==========================
// Configuration
// ==========================
const SCHEDULE_ID = "PGEHI7T";          // PagerDuty SQL-Sev1-OnCall schedule ID
const PD_TOKEN = "xxxxxx";       // replace it with your PagerDuty API token
const ON_CALL_SHEET_NAME = "Q4FY26 24x7 On-Call";  // 24*7 on call tab 
const SR_SHEET_NAME = "Q4FY26: Support Rotation";  // Suppport rotation tab
const DRY_RUN = true;                    // true: print only 
const SYNC_START_TIME = new Date();      // current time
const DEBUG = true;



const TEST_SCHEDULE_ID = "P5W7H6Z";          // test schedule ID
const TEST_PD_TOKEN = "xxxxxx";       // replace it with your PagerDuty API token
const TEST_SHEET_NAME = "Q4FY26 24x7 On-Call";  // 24*7 on call tab 
const TEST_DRY_RUN = true;                    
const TEST_SYNC_START_TIME = new Date();      
const TEST_DEBUG = true;
const WEB = "https://script.google.com/a/macros/snowflake.com/s/AKfycbyRFLNTe3BUwSoXa6ZS4QvCCOtrbyZsyMRA7jRXVus2Ety5IWN6_35TUcQwEKVL3n6A2w/exec";


/***************************************
 * Global variables
 ***************************************/

// Global map: engineer name -> PagerDuty user ID
let ENGINEER_MAP = {};

// Global map for testing purpose: engineer name -> PagerDuty user ID
let TEST_ENGINEER_MAP = {};


/***************************************
 * Utility functions
 ***************************************/

/**
 * Conditional logger
 * @param {string} message - Log message
 */
function cmm_log(message) {
  if (DEBUG) {
    Logger.log(message);
  }
}


/***************************************
 * Web app test
 ***************************************/

/**
 * Test PagerDuty Web API proxy
 */
function test_webapp() {
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Token token=" + PD_TOKEN,
      "Accept": "application/vnd.pagerduty+json;version=2"
    },
    payload: JSON.stringify({ foo: "bar" }),
    muteHttpExceptions: true // Do not throw on 4xx/5xx
  };

  const response = UrlFetchApp.fetch(WEB, options);
  Logger.log(response.getContentText());

  const reqPayload = {
    method: "GET",
    url: "/schedules/SCHEDULE_ID/overrides"
  };

  const resp = UrlFetchApp.fetch(WEB, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(reqPayload)
  });

  Logger.log(resp.getContentText());
}


/***************************************
 * Engineer mapping initialization
 ***************************************/

/**
 * Initialize the global engineer map
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 *        Sheet containing engineer names
 * @param {Array<number>} columns
 *        Column indices (1-based) where engineer names appear
 */
function initEngineerMap(sheet, columns) {
  ENGINEER_MAP = generateEngineerMapping(sheet, columns);
  cmm_log(
    "Initialized ENGINEER_MAP with " +
    Object.keys(ENGINEER_MAP).length +
    " users."
  );
}


/***************************************
 * PagerDuty user lookup
 ***************************************/

/**
 * Look up PagerDuty user ID by calling PagerDuty API directly
 *
 * @param {string} name - Engineer name (may contain parentheses)
 * @returns {string|null} PagerDuty user ID or null if not found
 */
function getUserId_directly(name) {
  if (!name) return null;

  // Remove anything inside parentheses, e.g. "Alice (NY)" -> "Alice"
  const cleanName = name
    .toString()
    .replace(/\s*\(.*?\)\s*/g, "")
    .trim();

  Logger.log("Checking PagerDuty user: " + cleanName);

  const url =
    "https://api.pagerduty.com/users?query=" +
    encodeURIComponent(cleanName);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        "Authorization": "Token token=" + PD_TOKEN,
        "Accept": "application/vnd.pagerduty+json;version=2"
      }
    });

    const result = JSON.parse(response.getContentText());
    const users = result.users || [];

    if (users.length === 0) {
      Logger.log("❌ getUserId_directly: user not found: " + cleanName);
      return null;
    }

    // Return the first matched user ID
    return users[0].id;

  } catch (err) {
    Logger.log("❌ getUserId_directly API call failed: " + err);
    return null;
  }
}


/**
 * Test getUserId_directly()
 */
function test_getUserId_directly() {
  const testNames = [
    "Adrian Neumann",
    "Di Liu",
    "Adrian Neumann (Berlin)",
    "Alice Chen",
    "Nonexistent User"
  ];

  testNames.forEach(name => {
    const id = getUserId_directly(name);
    if (id) {
      Logger.log(`✅ PagerDuty ID for "${name}": ${id}`);
    } else {
      Logger.log(`❌ User not found: "${name}"`);
    }
  });
}


/***************************************
 * Cached lookup using ENGINEER_MAP
 ***************************************/

/**
 * Get PagerDuty user ID from the pre-generated map
 *
 * @param {string} name - Engineer name
 * @returns {string|null} PagerDuty user ID or null if not found
 */
function getUserId(name) {
  if (!name) return null;

  const pdId = ENGINEER_MAP[name.trim()];
  return pdId || null;
}


/***************************************
 * Mapping generation
 ***************************************/

/**
 * Generate engineer name -> PagerDuty user ID mapping
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 *        Google Sheet object
 * @param {Array<number>} columns
 *        Columns to scan (1-based, e.g. F=6, I=9, M=13)
 * @returns {Object} Mapping { "Engineer Name": "PagerDuty User ID" }
 */
function generateEngineerMapping(sheet, columns) {
  const data = sheet.getDataRange().getValues();
  const map = {};

  // Fetch all PagerDuty users once
  const allUsers = getAllUsers(); // [{ id, name }, ...]

  // Build a lowercase name -> id index
  const nameToId = {};
  allUsers.forEach(user => {
    if (user.name) {
      nameToId[user.name.toLowerCase()] = user.id;
    }
  });

  // Start from row 4 (skip headers)
  for (let i = 3; i < data.length; i++) {
    const row = data[i];

    columns.forEach(col => {
      const name = row[col - 1];

      if (name && name.toString().trim() !== "") {
        // Remove parentheses from names
        const cleanName = name
          .toString()
          .replace(/\s*\(.*?\)\s*/g, "")
          .trim();

        const pdId = nameToId[cleanName.toLowerCase()];
        if (pdId) {
          map[cleanName] = pdId;
        } else {
          Logger.log(
            "⚠️ Row " +
            (i + 1) +
            ": PagerDuty user not found: " +
            cleanName
          );
        }
      }
    });
  }

  return map;
}


/***************************************
 * PagerDuty user listing
 ***************************************/

/**
 * Fetch all PagerDuty users using pagination
 *
 * @returns {Array<Object>} List of PagerDuty users
 */
function getAllUsers() {
  const limit = 100; // Page size
  let offset = 0;
  let allUsers = [];
  let more = true;

  while (more) {
    const url = `https://api.pagerduty.com/users?limit=${limit}&offset=${offset}`;

    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        "Authorization": "Token token=" + PD_TOKEN,
        "Accept": "application/json"
      }
    });

    const data = JSON.parse(response.getContentText());
    const users = data.users || [];
    allUsers = allUsers.concat(users);

    // Determine whether there are more pages
    if (users.length < limit || !data.more) {
      more = false;
    } else {
      offset += limit;
    }
  }

  return allUsers;
}


/***************************************
 * Spreadsheet binding checks
 ***************************************/

/**
 * Check whether the script is bound to a spreadsheet
 */
function check() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Active Spreadsheet: " + ss);

  if (ss) {
    Logger.log("Name = " + ss.getName());
    Logger.log("ID = " + ss.getId());
    Logger.log("Sheets = " + ss.getSheets().map(s => s.getName()));
  } else {
    Logger.log("This is NOT a container-bound script.");
  }
}


/**
 * Check spreadsheet binding (simplified)
 */
function checkSpreadsheetBinding() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    Logger.log("❌ This script is not bound to any Spreadsheet");
  } else {
    Logger.log("Spreadsheet name: " + ss.getName());
    Logger.log("Spreadsheet ID: " + ss.getId());
  }
}


/***************************************
 * Sync entry points
 ***************************************/

/**
 * Test sync (dry run)
 */
function testSync() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(ON_CALL_SHEET_NAME);

  Logger.log("Sheet: " + sheet.getName());
  syncPagerDutyFuture(sheet, false, true);
}


/**
 * Full sync (non-test)
 */
function syncALL() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(ON_CALL_SHEET_NAME);

  Logger.log("Sheet: " + sheet.getName());
  syncPagerDutyFuture(sheet, false, false);
}


/**
 * Test sync against test schedule
 */
function testSync_test() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(ON_CALL_SHEET_NAME);

  Logger.log("Sheet: " + sheet.getName());
  syncPagerDutyFuture(sheet, false, true);
}


/***************************************
 * Core row-level update logic
 ***************************************/

/**
 * Update PagerDuty schedule/override based on a single row in Google Sheets
 * - Only applies to columns A/B/C (day, time, engineer)
 * - Supports dry run
 *
 * @param {number} rowNum - Row number (1-based)
 * @param {string} day
 * @param {string} time
 * @param {string} engineer
 * @param {boolean} dryRun - Whether to run in dry-run mode
 * @param {boolean} test - Whether to use test schedule
 */
function updatePagerDutyForRow(
  rowNum,
  day,
  time,
  engineer,
  dryRun = true,
  test = true
) {
  Logger.log(
    `Processing row ${rowNum}: day=${day}, time=${time}, engineer=${engineer}`
  );

  // ---- Validate required fields ----
  if (!day || !time || !engineer) {
    Logger.log(
      `Row ${rowNum} has incomplete data; skipping or deleting corresponding entry`
    );

    if (!dryRun) {
      deletePagerDutyEntry(rowNum); // Deletion logic must be implemented separately
    }
    return;
  }

  // ---- Parse time range ----
  const timeStr = time
    .toString()
    .replace(/[\r\n\u2028\u2029]+/g, " ")
    .trim();

  let startISO, endISO;
  try {
    [startISO, endISO] = parseTime(day, timeStr, rowNum);
  } catch (e) {
    Logger.log(`Row ${rowNum}: parseTime failed: ${e.message}`);
    return;
  }

  // ---- Skip entries ending before SYNC_START_TIME ----
  if (new Date(endISO) < SYNC_START_TIME) {
    Logger.log(`Row ${rowNum}: end < SYNC_START_TIME, skipping`);
    return;
  }

  // ---- Resolve PagerDuty user ID ----
  const cleanName = engineer
    .toString()
    .trim()
    .replace(/\s*\(.*?\)\s*/g, "");

  const userId = getUserId_directly(cleanName);

  if (!userId) {
    Logger.log(`Row ${rowNum}: PagerDuty user not found: ${cleanName}`);
    throw new Error(`PagerDuty User ID not found for "${cleanName}"`);
  }

  Logger.log(
    `Row ${rowNum} user resolved: original="${engineer}", clean="${cleanName}", userId="${userId}"`
  );

  // ---- Fetch future schedule entries and overrides ----
  BEGIN_TIME = new Date(startISO);
  END_TIME = new Date(endISO);

  const scheduleId = test ? TEST_SCHEDULE_ID : SCHEDULE_ID;

  const futureEntries = generateFutureScheduleEntries(
    scheduleId,
    BEGIN_TIME,
    END_TIME
  );
  const pdOverrides = getOverrides(scheduleId, BEGIN_TIME, END_TIME);

  // ---- Matching logic ----
  const relatedFutureEntries = futureEntries.filter(
    entry => entry.user.id === userId
  );

  relatedFutureEntries.forEach(entry => {
    const isMatch = timesMatch(
      entry.start,
      entry.end,
      startISO,
      endISO
    );

    Logger.log(
      `  FutureEntry start=${entry.start}, end=${entry.end}, match=${isMatch}, userId=${entry.user.id}`
    );
  });

  const hasMatch = relatedFutureEntries.some(entry =>
    timesMatch(entry.start, entry.end, startISO, endISO)
  );

  if (!hasMatch) {
    const alreadyOverride = pdOverrides.find(
      o =>
        timesMatch(o.start, o.end, startISO, endISO) &&
        o.user.id === userId
    );

    if (!alreadyOverride) {
      Logger.log(
        `[CREATE OVERRIDE] start=${startISO}, end=${endISO}, userId=${userId}, name=${cleanName}`
      );

      if (!dryRun) {
        const payload = {
          override: {
            start: startISO,
            end: endISO,
            user: {
              id: userId,
              type: "user_reference"
            }
          }
        };

        const url = `https://api.pagerduty.com/schedules/${scheduleId}/overrides`;

        UrlFetchApp.fetch(url, {
          method: "post",
          contentType: "application/json",
          headers: {
            "Authorization": "Token token=" + PD_TOKEN,
            "Accept": "application/vnd.pagerduty+json;version=2"
          },
          payload: JSON.stringify(payload)
        });

        Logger.log("Override successfully created");
      }
    } else {
      Logger.log("Override already exists; skipping creation");
    }
  } else {
    Logger.log("Future schedule entry already matches; no override needed");

    relatedFutureEntries.forEach(entry => {
      if (
        timesMatch(entry.start, entry.end, startISO, endISO)
      ) {
        Logger.log(
          `[MATCHED FUTURE ENTRY] start=${entry.start}, end=${entry.end}, user=${entry.user.summary || entry.user.id}`
        );
      }
    });
  }
}


/***************************************
 * Test helpers
 ***************************************/

/**
 * Test updatePagerDutyForRow (dry run, test schedule)
 */
function test_updatePagerDutyForRow() {
  const rowNum = 5;
  const day = "Wed Dec 24, 2025";      // Simulates column A
  const time = "12am - 8am PT";        // Simulates column B
  const engineer = "Elena Cai";        // Simulates column C

  updatePagerDutyForRow(rowNum, day, time, engineer, true, true);
}


/**
 * Test updatePagerDutyForRow (real run, test schedule)
 */
function test_updatePagerDutyForRow_test() {
  const rowNum = 5;
  const day = "Wed Dec 26, 2025";      // Simulates column A
  const time = "12am - 8am PT";        // Simulates column B
  const engineer = "Elena Cai";        // Simulates column C

  updatePagerDutyForRow(rowNum, day, time, engineer, false, true);
}


/***************************************
 * PagerDuty entry helpers
 ***************************************/

/**
 * Placeholder for deleting PagerDuty entries by row
 *
 * @param {number} rowNum - Row number
 */
function deletePagerDutyEntry(rowNum) {
  Logger.log(`[DRY RUN DELETE] PagerDuty entry for row ${rowNum}`);
  // You may locate overrides by start/end and delete them here
}


/**
 * Push a PagerDuty override entry
 *
 * @param {Object} entry - Override entry { start, end, user }
 */
function pushPagerDutyEntry(entry) {
  const url = "https://api.pagerduty.com/schedules/.../overrides";

  const headers = {
    "Authorization": "Token token=" + PD_TOKEN,
    "Content-Type": "application/json",
    "Accept": "application/vnd.pagerduty+json;version=2"
  };

  const body = {
    override: {
      start: entry.start,
      end: entry.end,
      user: {
        id: entry.user,
        type: "user_reference"
      }
    }
  };

  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers,
    payload: JSON.stringify(body)
  });

  Logger.log("PagerDuty update succeeded: " + JSON.stringify(entry));
}


/***************************************
 * Trigger: onEdit handler
 ***************************************/

/**
 * Triggered when the sheet is edited
 * - Clears previous error notes
 * - Syncs PagerDuty when A/B/C columns change
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit_setNote(e) {
  const cell = e.range;

  try {
    // Clear previous error notes to avoid stacking
    cell.clearNote();

    if (!e || !e.range) return;

    const sheet = e.source.getActiveSheet();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Only handle columns A/B/C (1,2,3) and non-header rows
    if (![1, 2, 3].includes(col) || row < 4) return;

    const oldValue = e.oldValue || "";
    const newValue = e.value || "";

    const day = sheet.getRange(row, 1).getValue();
    const time = sheet.getRange(row, 2).getValue();
    const engineer = sheet
      .getRange(row, 3)
      .getValue()
      .toString()
      .trim(); // Remove trailing spaces

    // Detect cell cleared (engineer removed)
    if (col === 3 && oldValue && !newValue) {
      Logger.log(
        `Row ${row} column ${col} cleared. Old value: ${oldValue}`
      );
      deletePagerDutyOverride(row, day, time, oldValue);
    } else {
      updatePagerDutyForRow(row, day, time, engineer, false, false);
    }

  } catch (err) {
    e.range.setNote("Error: " + err.message);
    SpreadsheetApp.flush();
    Browser.msgBox("Error: " + err.message);
  }
}


/***************************************
 * PagerDuty override deletion
 ***************************************/

/**
 * Delete PagerDuty overrides matching a specific row/time/user
 *
 * @param {number} row
 * @param {string} day
 * @param {string} time
 * @param {string} engineer
 */
function deletePagerDutyOverride(row, day, time, engineer) {
  Logger.log("=== deletePagerDutyOverride called ===");
  Logger.log("Row: " + row);
  Logger.log("Day: " + day);
  Logger.log("Time: " + time);
  Logger.log("Engineer: " + engineer);

  const [startISO, endISO] = parseTime(day, time, row);
  const userId = getUserId_directly(engineer.trim());

  const overrides = getOverrides(
    SCHEDULE_ID,
    new Date(startISO),
    new Date(endISO)
  );

  overrides.forEach(o => {
    if (
      o.user.id === userId &&
      timesMatch(o.start, o.end, startISO, endISO)
    ) {
      const url =
        `https://api.pagerduty.com/schedules/${SCHEDULE_ID}/overrides/${o.id}`;

      UrlFetchApp.fetch(url, {
        method: "delete",
        headers: {
          "Authorization": "Token token=" + PD_TOKEN,
          "Accept": "application/vnd.pagerduty+json;version=2"
        }
      });

      Logger.log(
        `Deleted override ID=${o.id}, user=${engineer}, row=${row}`
      );
    }
  });
}


/**
 * Test helper for deletePagerDutyOverride
 */
function testDeletePagerDutyOverride() {
  const row = 5;
  const day = "Mon Jan 5, 2026";
  const time = "4pm - 12am PT";
  const engineer = "Dmytro Koval";

  Logger.log("=== Test deletePagerDutyOverride ===");
  Logger.log(
    `Row=${row}, Day=${day}, Time=${time}, Engineer=${engineer}`
  );

  try {
    const [startISO, endISO] = parseTime(day, time, row);
    const userId = getUserId_directly(engineer.trim());

    const overrides = getOverrides(
      SCHEDULE_ID,
      new Date(startISO),
      new Date(endISO)
    );

    if (overrides.length === 0) {
      Logger.log("No overrides found for this time/user");
    } else {
      overrides.forEach(o => {
        Logger.log(
          `[DRY RUN] Found override -> ID=${o.id || "undefined"}, ` +
          `start=${o.start}, end=${o.end}, userId=${o.user.id}`
        );
      });
    }

    // Uncomment to perform actual deletion
    deletePagerDutyOverride(row, day, time, engineer);

  } catch (err) {
    Logger.log(
      "Error in testDeletePagerDutyOverride: " + err.message
    );
  }

  Logger.log("=== End test ===");
}


/***************************************
 * Time parsing
 ***************************************/

/**
 * Parse day + PT time range into ISO timestamps
 *
 * @param {string} dayStr
 * @param {string} timeStr
 * @param {number} rowNum
 * @returns {[string, string]} [startISO, endISO]
 */
function parseTime(dayStr, timeStr, rowNum) {
  // Remove holidays, extra annotations, and newlines
  dayStr = dayStr
    .toString()
    .trim()
    .split(/\r?\n/)[0]
    .split(/US holiday|:/)[0]
    .trim();

  const dayDate = new Date(dayStr);

  // Normalize whitespace and invisible characters
  timeStr = timeStr
    .replace(/[\r\n\u2028\u2029]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  // Match the first PT time segment, e.g. "4pm - 12am PT"
  const regex =
    /(\d{1,2}(?::\d{2})?\s*(?:am|pm)(?:\+1d)?\s*-\s*\d{1,2}(?::\d{2})?\s*(?:am|pm)(?:\+1d)?)\s*PT/i;

  const match = regex.exec(timeStr);
  if (!match) {
    throw new Error(
      `parseTime failed to find PT segment (row ${rowNum}): ${timeStr}`
    );
  }

  const [startStr, endStr] = match[1].split(/\s*-\s*/);

  function parseHourMinute(str) {
    const plusDay = str.includes("+1d");
    str = str.replace("+1d", "");

    const m = str.match(/(\d+)(?::(\d+))?\s*(am|pm)/i);
    if (!m) throw new Error("Invalid time format: " + str);

    let hour = parseInt(m[1], 10);
    const minute = m[2] ? parseInt(m[2], 10) : 0;
    const meridiem = m[3].toLowerCase();

    if (meridiem === "pm" && hour !== 12) hour += 12;
    if (meridiem === "am" && hour === 12) hour = 0;

    return { hour, minute, plusDay };
  }

  const startHM = parseHourMinute(startStr);
  const endHM = parseHourMinute(endStr);

  const startDate = new Date(dayDate);
  startDate.setHours(startHM.hour, startHM.minute, 0, 0);

  const endDate = new Date(dayDate);
  endDate.setHours(endHM.hour, endHM.minute, 0, 0);

  // Advance to next day if needed
  if (startHM.plusDay || endHM.plusDay || endDate <= startDate) {
    endDate.setDate(endDate.getDate() + 1);
  }

  const tzOffset = "-08:00"; // PT timezone (PST)

  function toISOStringWithTZ(date, offset) {
    const pad = n => n.toString().padStart(2, "0");
    return (
      `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` +
      `T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}` +
      offset
    );
  }

  const startISO = toISOStringWithTZ(startDate, tzOffset);
  const endISO = toISOStringWithTZ(endDate, tzOffset);

  Logger.log(
    `Row ${rowNum} PT segment parsed: start=${startISO}, end=${endISO}`
  );

  return [startISO, endISO];
}

/***************************************
 * PagerDuty Schedule & Layer Helpers
 ***************************************/

/**
 * Fetch PagerDuty schedule's final_schedule and entries within a time range
 *
 * @param {string} scheduleId - PagerDuty schedule ID
 * @param {Date} syncStartTime - Start datetime
 * @param {Date} syncEndTime - End datetime
 * @returns {Object} { finalSchedule: {}, entries: [] }
 */
function getFinalSchedule(scheduleId, syncStartTime, syncEndTime) {
  const entriesLimit = 100;
  let entriesOffset = 0;
  let allEntries = [];
  let more = true;

  // 1️⃣ Fetch schedule, including final_schedule
  const scheduleUrl = `https://api.pagerduty.com/schedules/${scheduleId}`;
  Logger.log(scheduleUrl);

  const scheduleResponse = UrlFetchApp.fetch(scheduleUrl, {
    method: "get",
    headers: {
      "Authorization": "Token token=" + PD_TOKEN,
      "Accept": "application/vnd.pagerduty+json;version=2"
    }
  });

  const scheduleData = JSON.parse(scheduleResponse.getContentText());
  const finalSchedule = scheduleData.schedule.final_schedule;
  Logger.log("Fetched final_schedule snapshot");
  Logger.log(JSON.stringify(finalSchedule, null, 2));

  // 2️⃣ Fetch schedule entries within the specified time range (paginated)
  while (more) {
    const entriesUrl = `https://api.pagerduty.com/schedules/${scheduleId}/entries` +
      `?since=${syncStartTime.toISOString()}` +
      `&until=${syncEndTime.toISOString()}` +
      `&limit=${entriesLimit}&offset=${entriesOffset}`;

    Logger.log("Fetching entries from: " + entriesUrl);

    const response = UrlFetchApp.fetch(entriesUrl, {
      method: "get",
      headers: {
        "Authorization": "Token token=" + PD_TOKEN,
        "Accept": "application/vnd.pagerduty+json;version=2"
      }
    });

    const result = JSON.parse(response.getContentText());
    const entries = result.schedule_entries || [];
    allEntries = allEntries.concat(entries);

    if (!result.more || entries.length < entriesLimit) {
      more = false;
    } else {
      entriesOffset += entriesLimit;
    }
  }

  Logger.log(`Fetched ${allEntries.length} entries for schedule ${scheduleId}`);

  return {
    finalSchedule: finalSchedule,
    entries: allEntries
  };
}


/**
 * Fetch entries for a specific layer within a PagerDuty schedule
 *
 * @param {string} scheduleId - PagerDuty schedule ID
 * @param {Date} sinceDate - Start datetime
 * @param {Date} untilDate - End datetime
 * @param {string} layerName - Layer name (e.g., "Layer 1 (backup)")
 * @returns {Array} Array of schedule entries
 */
function getScheduleLayerEntries(scheduleId, sinceDate, untilDate, layerName) {
  // 1️⃣ Fetch schedule metadata
  const scheduleUrl = `https://api.pagerduty.com/schedules/${scheduleId}`;
  const scheduleResp = UrlFetchApp.fetch(scheduleUrl, {
    headers: {
      "Authorization": "Token token=" + PD_TOKEN,
      "Accept": "application/vnd.pagerduty+json;version=2"
    }
  });

  const scheduleData = JSON.parse(scheduleResp.getContentText());
  if (!scheduleData.schedule || !scheduleData.schedule.schedule_layers) {
    throw new Error(`Failed to fetch schedule or schedule_layers: ${scheduleId}`);
  }

  // 2️⃣ Locate the specified layer
  const layer = scheduleData.schedule.schedule_layers.find(l => l.name === layerName);
  if (!layer) throw new Error(`Layer not found: ${layerName} in schedule ${scheduleId}`);

  Logger.log("Layer metadata: " + JSON.stringify(layer, null, 2));
  const layerId = layer.id;

  // 3️⃣ Fetch layer entries (paginated)
  const limit = 100;
  let offset = 0;
  let allEntries = [];
  let more = true;

  while (more) {
    const url = `https://api.pagerduty.com/schedules/${scheduleId}/entries` +
      `?since=${sinceDate.toISOString()}` +
      `&until=${untilDate.toISOString()}` +
      `&layer_ids[]=${layerId}&limit=${limit}&offset=${offset}`;

    Logger.log("Fetching schedule entries: " + url);

    const response = UrlFetchApp.fetch(url, {
      headers: {
        "Authorization": "Token token=" + PD_TOKEN,
        "Accept": "application/json"
      }
    });

    const data = JSON.parse(response.getContentText());
    const entries = data.schedule_entries || [];
    allEntries = allEntries.concat(entries);

    if (entries.length < limit || !data.more) {
      more = false;
    } else {
      offset += limit;
    }
  }

  Logger.log(`Fetched ${allEntries.length} entries for layer: ${layerName}`);
  return allEntries;
}


/**
 * Lookup user name from PagerDuty ID
 *
 * @param {string} userId - PagerDuty user ID
 * @returns {string} User name or "Unknown User" if not found
 */
function getUserNameById(userId) {
  if (!ENGINEER_MAP) return "Unknown User";

  for (let name in ENGINEER_MAP) {
    if (ENGINEER_MAP[name] === userId) return name;
  }
  return "Unknown User";
}


/**
 * Compare two time ranges, allowing small end-time difference
 *
 * @param {string|Date} start1
 * @param {string|Date} end1
 * @param {string|Date} start2
 * @param {string|Date} end2
 * @returns {boolean} True if start matches exactly and end differs <= 2 minutes
 */
function timesMatch(start1, end1, start2, end2) {
  const s1 = new Date(start1).getTime();
  const e1 = new Date(end1).getTime();
  const s2 = new Date(start2).getTime();
  const e2 = new Date(end2).getTime();

  return s1 === s2 && Math.abs(e1 - e2) <= 120 * 1000;
}


/**
 * Delete all PagerDuty overrides within a time range
 *
 * @param {Sheet} sheet - Optional sheet for logging
 * @param {string} schedule_id - PagerDuty schedule ID
 * @param {Date} startTime - Start datetime
 * @param {Date} endTime - End datetime
 * @param {boolean} dryRun - If true, only logs without deleting
 */
function delete_all_overrides(sheet, schedule_id, startTime, endTime, dryRun = true) {
  Logger.log(`Deleting overrides: schedule=${schedule_id}, start=${startTime.toISOString()}, end=${endTime.toISOString()}`);

  const overrides = getOverrides(schedule_id, startTime, endTime);

  Logger.log(`Found ${overrides.length} overrides`);

  if (dryRun) {
    overrides.forEach(o => {
      Logger.log(`[DRY RUN] Will delete override -> ID: ${o.id}, Start: ${o.start}, End: ${o.end}, User ID: ${o.user.id}, Name: ${getUserNameById(o.user.id)}`);
    });
    return;
  }

  overrides.forEach(o => {
    const url = `https://api.pagerduty.com/schedules/${schedule_id}/overrides/${o.id}`;
    try {
      UrlFetchApp.fetch(url, {
        method: "delete",
        headers: {
          "Authorization": "Token token=" + PD_TOKEN,
          "Accept": "application/vnd.pagerduty+json;version=2"
        }
      });
      Logger.log(`Deleted override -> Start: ${o.start}, End: ${o.end}, User ID: ${o.user.id}, Name: ${getUserNameById(o.user.id)}`);
    } catch (err) {
      Logger.log(`Error deleting override ${o.id}: ${err.message}`);
    }
  });

  Logger.log("Finished deleting overrides");
}


/**
 * Test deleting all PagerDuty overrides within a time range
 */
function test_delete_all_overrides() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const schedule_id = TEST_SCHEDULE_ID; // Test Schedule ID
  const startTime = new Date("2025-11-25T00:00:00-08:00"); // Test start time
  const endTime = new Date("2026-01-31T00:00:00-08:00");   // Test end time
  const dryRun = false; // false = actually delete, true = only dry run

  Logger.log("==== Testing delete_all_overrides (Dry Run) ====");
  delete_all_overrides(sheet, schedule_id, startTime, endTime, dryRun);
  Logger.log("==== Test completed ====");
}

/**
 * Core function: sync future schedule entries to PagerDuty
 * @param {Sheet} sheet - Google Sheet object
 * @param {boolean} dryRun - Whether to only print Dry Run
 * @param {boolean} test - Whether to use test schedule ID
 */
function syncPagerDutyFuture(sheet, dryRun = true, test = true) {
  // Initialize Engineer map (columns F/I/M)
  initEngineerMap(sheet, [6, 9, 13]);

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(3); // Skip first 3 header rows
  if (!rows.length) return;

  // Get last non-empty row for determining sync end time
  let lastNonEmptyRow = null;
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][1] && rows[i][1].toString().trim() !== "") {
      lastNonEmptyRow = rows[i];
      break;
    }
  }
  if (!lastNonEmptyRow) throw new Error("No valid time data in spreadsheet");

  const [_, lastEndISO] = parseTime(lastNonEmptyRow[0], lastNonEmptyRow[1], rows.length + 3);
  const SYNC_END_TIME = new Date(lastEndISO);
  const SYNC_START_TIME = new Date(); // Start from now

  const spreadsheetEntries = [];

  // Process each row
  rows.forEach((row, i) => {
    let day = row[0], time = row[1], engineer = row[2];
    const rowNum = i + 4; // Actual Sheet row number
    const rowRange = sheet.getRange(rowNum, 1, 1, row.length);
    const cellC = sheet.getRange(rowNum, 3); // Column C for error notes

    // Clear old notes
    rowRange.clearNote();

    if (!day || !time || !engineer) return;

    if (day.toString().includes("\n")) day = day.toString().split("\n")[0];
    const timeStr = time.toString().replace(/[\r\n\u2028\u2029]+/g, " ").trim();

    let startISO, endISO;
    try {
      [startISO, endISO] = parseTime(day, timeStr, rowNum);
    } catch (e) {
      Logger.log(`Row ${rowNum}: parseTime failed: ${e.message}`);
      return;
    }

    if (new Date(endISO) < SYNC_START_TIME) {
      Logger.log(`Row ${rowNum}: End time before SYNC_START_TIME, skip: ${endISO}`);
      return;
    }

    const cleanName = engineer.toString().trim().replace(/\s*\(.*?\)\s*/g, "");
    const userId = getUserId(cleanName);
    if (!userId) {
      const msg = `PagerDuty user ID not found: ${cleanName}`;
      Logger.log(`Row ${rowNum} FATAL: ${msg}`);
      cellC.setNote("Error: " + msg);
      return;
    }

    spreadsheetEntries.push({ start: startISO, end: endISO, user: userId });
  });

  Logger.log(`Generated ${spreadsheetEntries.length} spreadsheet entries`);

  const schedule_id = test ? TEST_SCHEDULE_ID : SCHEDULE_ID;

  // Generate future schedule entries
  Logger.log("Generating future schedule entries...");
  const futureEntries = generateFutureScheduleEntries(schedule_id, SYNC_START_TIME, SYNC_END_TIME);
  Logger.log(`Generated ${futureEntries.length} future schedule entries`);

  // Fetch existing overrides
  const pdOverrides = getOverrides(schedule_id, SYNC_START_TIME, SYNC_END_TIME);
  Logger.log(`Existing overrides: ${pdOverrides.length}`);

  // ========================
  // Compare spreadsheet vs future schedule to determine overrides to create
  // ========================
  const toCreateOverride = [];

  spreadsheetEntries.forEach(sp => {
    if (new Date(sp.end) < SYNC_START_TIME) {
      cmm_log(`Skipping spreadsheet entry (end < SYNC_START_TIME): ${sp.user}, ${sp.start} -> ${sp.end}`);
      return;
    }

    const relatedFutureEntries = futureEntries.filter(f => f.user.id === sp.user);
    const hasMatch = relatedFutureEntries.some(f => timesMatch(f.start, f.end, sp.start, sp.end));

    if (!hasMatch) {
      const alreadyOverride = pdOverrides.find(o => timesMatch(o.start, o.end, sp.start, sp.end) && o.user.id === sp.user);
      if (!alreadyOverride) {
        toCreateOverride.push(sp);
      }
    }
  });

  cmm_log(`Overrides to create: ${toCreateOverride.length}`);

  // ========================
  // Dry Run output or actual creation
  // ========================
  if (dryRun) {
    cmm_log("=== DRY RUN ===");
    toCreateOverride.forEach(c => {
      const userName = getUserNameById(c.user);
      cmm_log(`Override to create -> Start: ${c.start}, End: ${c.end}, User ID: ${c.user}, Name: ${userName}`);
    });
  } else {
    toCreateOverride.forEach(c => {
      const payload = {
        override: {
          start: c.start,
          end: c.end,
          user: { id: c.user, type: "user_reference" }
        }
      };
      const url = `https://api.pagerduty.com/schedules/${schedule_id}/overrides`;
      UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        headers: {
          "Authorization": "Token token=" + PD_TOKEN,
          "Accept": "application/vnd.pagerduty+json;version=2"
        },
        payload: JSON.stringify(payload)
      });
    });
    Logger.log(`Created ${toCreateOverride.length} overrides`);
  }

  Logger.log("PagerDuty future schedule sync complete.");
}

/**
 * Get schedule entries for a given schedule and time range
 */
function getFinalScheduleEntries(scheduleId, syncStartTime, syncEndTime) {
  const PD_TOKEN = "u+NuJGtkb_7xAz2xB6xQ";
  const limit = 100;
  let offset = 0;
  let allEntries = [];
  let more = true;

  while (more) {
    const entriesUrl = `https://api.pagerduty.com/schedules/${scheduleId}/entries` +
                       `?since=${syncStartTime.toISOString()}` +
                       `&until=${syncEndTime.toISOString()}` +
                       `&limit=${limit}&offset=${offset}`;

    Logger.log("Fetching entries from: " + entriesUrl);

    const response = UrlFetchApp.fetch(entriesUrl, {
      method: "get",
      headers: {
        "Authorization": "Token token=" + PD_TOKEN,
        "Accept": "application/vnd.pagerduty+json;version=2"
      },
      muteHttpExceptions: true
    });

    const responseText = response.getContentText();
    const statusCode = response.getResponseCode();
    Logger.log("HTTP status: " + statusCode);
    Logger.log("Response text: " + responseText);

    if (statusCode === 404) {
      Logger.log("No entries generated for this time range, using final_schedule snapshot.");
      return getFinalScheduleSnapshot(scheduleId);
    }

    if (responseText && responseText.trim() !== "") {
      try {
        const result = JSON.parse(responseText);
        const entries = result.schedule_entries || [];
        allEntries = allEntries.concat(entries);
        if (!result.more || entries.length < limit) {
          more = false;
        } else {
          offset += limit;
        }
      } catch (e) {
        Logger.log("JSON parse error: " + e);
        more = false;
      }
    } else {
      Logger.log("Empty response, cannot parse JSON");
      more = false;
    }
  }

  Logger.log(`Fetched ${allEntries.length} entries for schedule ${scheduleId}`);
  return allEntries;
}

/**
 * Get final_schedule snapshot
 */
function getFinalScheduleSnapshot(scheduleId) {
  const PD_TOKEN = "u+NuJGtkb_7xAz2xB6xQ";
  const scheduleUrl = `https://api.pagerduty.com/schedules/${scheduleId}`;

  const response = UrlFetchApp.fetch(scheduleUrl, {
    method: "get",
    headers: {
      "Authorization": "Token token=" + PD_TOKEN,
      "Accept": "application/vnd.pagerduty+json;version=2"
    },
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  const finalEntries = result.schedule.final_schedule.rendered_schedule_entries || [];

  Logger.log(`Final schedule snapshot contains ${finalEntries.length} entries`);
  return finalEntries;
}

/**
 * Test fetching final schedule entries
 */
function testGetFinalScheduleEntries() {
  const scheduleId = "PGEHI7T";
  const start = new Date("2025-11-25T00:00:00Z");
  const end = new Date("2026-02-01T00:00:00Z");

  const entries = getFinalScheduleEntries(scheduleId, start, end);
  Logger.log(JSON.stringify(entries, null, 2));
}

/*********************************************
 * Generate future schedule entries
 * @param {string} scheduleId - PagerDuty Schedule ID
 * @param {Date} syncStartTime - Start time
 * @param {Date} syncEndTime - End time
 * @returns {Array} List of schedule entries
 *********************************************/
function generateFutureScheduleEntries(scheduleId, syncStartTime, syncEndTime) {
  Logger.log("generateFutureScheduleEntries called with:");
  Logger.log("  scheduleId: " + scheduleId);
  Logger.log("  syncStartTime: " + syncStartTime.toISOString());
  Logger.log("  syncEndTime: " + syncEndTime.toISOString());

  // 1️⃣ Fetch schedule config (including layers)
  const scheduleUrl = `https://api.pagerduty.com/schedules/${scheduleId}`;
  const scheduleResp = UrlFetchApp.fetch(scheduleUrl, {
    method: "get",
    headers: {
      "Authorization": "Token token=" + PD_TOKEN,
      "Accept": "application/vnd.pagerduty+json;version=2"
    }
  });
  const schedule = JSON.parse(scheduleResp.getContentText()).schedule;
  const layers = schedule.schedule_layers || [];

  // 2️⃣ Fetch existing overrides
  const overrides = getOverrides(scheduleId, syncStartTime, syncEndTime);

  // 3️⃣ Iterate each layer and generate entries
  let entries = [];
  layers.forEach(layer => {
    const layerStart = new Date(layer.start);
    const layerEnd = layer.end ? new Date(layer.end) : syncEndTime;

    Logger.log("Processing layer: " + layer.name);

    const rangeStart = new Date(Math.max(layerStart, syncStartTime));
    const rangeEnd = new Date(Math.min(layerEnd, syncEndTime));

    const users = layer.users.map(u => u.user);
    const turnLengthMs = layer.rotation_turn_length_seconds * 1000;

    let current = new Date(rangeStart);
    let userIndex = 0;

    while (current < rangeEnd) {
      const next = new Date(current.getTime() + turnLengthMs);
      const entryEnd = next > rangeEnd ? rangeEnd : next;

      entries.push({
        start: current.toISOString(),
        end: entryEnd.toISOString(),
        user: users[userIndex % users.length]
      });

      current = entryEnd;
      userIndex++;
    }
  });

  // 4️⃣ Apply overrides
  overrides.forEach(o => {
    const oStart = new Date(o.start);
    const oEnd = new Date(o.end);

    entries = entries.map(e => {
      const eStart = new Date(e.start);
      const eEnd = new Date(e.end);

      if (!(eEnd <= oStart || eStart >= oEnd)) {
        return { start: e.start, end: e.end, user: o.user };
      }
      return e;
    });
  });

  Logger.log("Generated " + entries.length + " future schedule entries.");
  entries.forEach(e => {
    Logger.log(`Start: ${e.start}, End: ${e.end}, User: ${e.user.summary}`);
  });

  return entries;
}

/**
 * Fetch overrides via WebApp
 */
function getOverridesViaWebApp(scheduleId, sinceDate, untilDate) {
  const webAppUrl = WEB;
  Logger.log("WebApp URL: " + WEB);

  const payload = {
    action: "get_overrides",
    scheduleId,
    sinceDate: sinceDate.toISOString(),
    untilDate: untilDate.toISOString()
  };

  const resp = UrlFetchApp.fetch(webAppUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(resp.getContentText());
  if (result.status === "ok") {
    return result.overrides;
  } else {
    Logger.log("Error fetching overrides: " + result.message);
    return [];
  }
}

/**
 * Fetch overrides from PagerDuty API
 */
function getOverrides(scheduleId, sinceDate, untilDate) {
  const limit = 100;
  let offset = 0;
  let allOverrides = [];
  let more = true;

  while (more) {
    const url = `https://api.pagerduty.com/schedules/${scheduleId}/overrides` +
                `?since=${sinceDate.toISOString()}` +
                `&until=${untilDate.toISOString()}` +
                `&limit=${limit}&offset=${offset}`;

    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        "Authorization": "Token token=" + PD_TOKEN,
        "Accept": "application/vnd.pagerduty+json;version=2"
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log(`Error fetching overrides: HTTP ${response.getResponseCode()}`);
      break;
    }

    const result = JSON.parse(response.getContentText());
    const overrides = (result.overrides || []).map((o, i) => {
      Logger.log(`Override ${i}: id=${o.id}, start=${o.start}, end=${o.end}, user=${o.user.summary}`);
      return { id: o.id, start: o.start, end: o.end, user: o.user };
    });

    allOverrides = allOverrides.concat(overrides);

    if (!result.more || overrides.length < limit) {
      more = false;
    } else {
      offset += limit;
    }
  }

  Logger.log(`Total overrides fetched: ${allOverrides.length}`);
  return allOverrides;
}

/*********************************************
 * Constants
 *********************************************/
const PST_COLUMNS = [3, 4];   // C, D
const CET_COLUMNS = [5, 6];   // E, F

const PST_SHIFT = { start: 8, end: 16 };       // 8am - 4pm PST
const CET_SHIFT = { start: 0, end: 8 };        // 12am - 8am PST

const ENGINEER_COLUMNS = [3, 4, 5, 6]; // C-F
const HEADER_ROW = 3;                  // Header row index

/*********************************************
 * Event handlers for onSelectionChange / onEdit / onChange
 *********************************************/
function onSelectionChange(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  if (sheetName === 'Q4FY26: Support Rotation') {
    handleSelection_SR(e);
  }
}

function handleSelection_SR(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SR_SHEET_NAME) return;

  const cell = sheet.getActiveCell();
  const row = cell.getRow();
  const col = cell.getColumn();

  if (![3, 4, 5, 6].includes(col) || row < 4) return;

  const key = `${sheet.getName()}!R${row}C${col}`;
  const props = PropertiesService.getDocumentProperties();
  const prev = props.getProperty(key);
  const nowBold = cell.getFontWeight() === "bold";
  const now = nowBold ? "bold" : "normal";

  if (prev === now) return; // no change

  props.setProperty(key, now);

  const engineer = cell.getValue() ? cell.getValue().toString().trim() : "";
  const startDate = sheet.getRange(row, 1).getValue();
  const endDate = sheet.getRange(row, 2).getValue();
  if (!startDate || !endDate) return;

  if (nowBold && engineer) {
    Logger.log(`BOLD format → create shifts for ${engineer}`);
    try {
      processBoldEngineerCell(row, col, startDate, engineer, endDate);
    } catch (err) {
      cell.setNote("Error: " + err.message);
      SpreadsheetApp.flush();
      Browser.msgBox("Error: " + err.message);
    }
  } else {
    Logger.log(`UNBOLD format → remove shifts for ${engineer}`);
    cell.clearNote();
    removeShiftsForEngineerDay(startDate, engineer, endDate, col);
  }
}

function onEdit_mm(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  switch(sheetName) {
    case 'Q4FY26: Support Rotation':
      Logger.log("Running onEdit_SR");
      onEdit_SR(e);
      break;
    case 'Q4FY26 24x7 On-Call':
      Logger.log("Running onEdit_setNote");
      onEdit_setNote(e);
      break;
    default:
      return;
  }
}

function onChange_mm(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  if (sheetName !== SR_SHEET_NAME) return;

  Logger.log(`onChange_mm triggered on sheet: ${sheetName}`);

  const range = e.range || sheet.getActiveRange();
  const startRow = range.getRow();
  const endRow = startRow + range.getNumRows() - 1;
  const startCol = range.getColumn();
  const endCol = startCol + range.getNumColumns() - 1;

  for (let row = startRow; row <= endRow; row++) {
    if (row <= HEADER_ROW) continue;

    for (let col = startCol; col <= endCol; col++) {
      if (!ENGINEER_COLUMNS.includes(col)) continue;

      const cell = sheet.getRange(row, col);
      const isBold = cell.getFontWeight() === "bold";
      const engineer = cell.getValue() ? cell.getValue().toString().trim() : "";
      const startDate = sheet.getRange(row, 1).getValue();
      const endDate = sheet.getRange(row, 2).getValue();

      cell.clearNote();

      if (!startDate || !endDate) {
        Logger.log(`Missing start/end date at row ${row}`);
        continue;
      }

      if (isBold && engineer) {
        Logger.log(`BOLD → create shifts for ${engineer} at row ${row}, col ${col}`);
        try {
          processBoldEngineerCell(row, col, startDate, engineer, endDate);
        } catch (err) {
          cell.setNote("Error: " + err.message);
          SpreadsheetApp.flush();
          Browser.msgBox("Error: " + err.message);
        }
      } else {
        const oldEngineer = engineer || (e.oldValue ? e.oldValue.toString().trim() : engineer);
        Logger.log(`EMPTY or UNBOLD → remove shifts for ${oldEngineer || "(no engineer)"} at row ${row}, col ${col}`);
        removeShiftsForEngineerDay(startDate, oldEngineer, endDate, col);
        cell.clearNote();
      }
    }
  }
}

/*
Change matrix for cell edits:

Cell Value Change       | Old Value Bold? | New Value Bold? | Valid PD User? | Action Taken                     | Note / Log
------------------------|----------------|----------------|----------------|---------------------------------|------------
Empty → Bold            | N/A            | Yes            | Yes            | Create shifts for each day       | ✅ Shift created
Empty → Bold            | N/A            | Yes            | No             | Do not create shifts             | ❌ Note added: user not found
Bold → Unbold           | Yes            | No             | Yes/No         | Remove shifts for each day       | Clear any previous Note
Bold → Empty            | Yes            | N/A            | Yes/No         | Remove shifts for each day       | Clear any previous Note
Non-bold → Bold         | No             | Yes            | Yes            | Create shifts for each day       | ✅ Shift created
Non-bold → Bold         | No             | Yes            | No             | Do not create shifts             | ❌ Note added: user not found
Bold → Bold             | Yes            | Yes            | Yes            | Update shifts (recreate if needed) | ✅ Shift created / updated
Bold → Bold             | Yes            | Yes            | No             | Do not create shifts             | ❌ Note remains
Empty → Empty           | N/A            | N/A            | N/A            | No action                        | No log
Non-bold → Non-bold     | No             | No             | N/A            | No action                        | No log
*/

function onEdit_SR(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SR_SHEET_NAME) return;

  const cell = e.range;
  const row = cell.getRow();
  const col = cell.getColumn();

  // Sync bold state cache to avoid duplicate triggers onSelectionChange
  const key = `${sheet.getName()}!R${row}C${col}`;
  const props = PropertiesService.getDocumentProperties();
  const isBold = cell.getFontWeight() === "bold";
  props.setProperty(key, isBold ? "bold" : "normal");

  // Clear old error note
  cell.clearNote();

  Logger.log(`onEdit_SR triggered -> row=${row}, col=${col}`);

  // Skip header row
  if (row < 4) return;

  // Only handle columns C-F (engineer columns)
  if (![3, 4, 5, 6].includes(col)) return;

  const engineer = cell.getValue() ? cell.getValue().toString().trim() : "";
  const startDate = sheet.getRange(row, 1).getValue(); // Column A
  const endDate = sheet.getRange(row, 2).getValue();   // Column B

  if (!startDate || !endDate) {
    Logger.log("Missing startDate or endDate");
    return;
  }

  if (isBold && engineer) {
    // Bold → create shifts
    Logger.log(`Cell turned BOLD → create shifts for ${engineer}`);
    try {
      processBoldEngineerCell(row, col, startDate, engineer, endDate);
    } catch (err) {
      // Only set note if error occurs
      cell.setNote("Error: " + err.message);
      SpreadsheetApp.flush();
      Browser.msgBox("Error: " + err.message);
    }
  } else if (!engineer || !isBold) {
    // Empty or non-bold → remove shifts
    const oldEngineer = e.oldValue ? e.oldValue.toString().trim() : engineer;
    Logger.log(`Cell is EMPTY or UNBOLD → remove shifts for ${oldEngineer || "(no engineer)"}`);
    removeShiftsForEngineerDay(startDate, oldEngineer, endDate, col);
    cell.clearNote();
  }
}

function removeShiftsForEngineerDay(startDate, engineerName, endDate, col) {
  Logger.log("=== removeShiftsForEngineerDay START ===");
  Logger.log(`Engineer: "${engineerName}"`);
  Logger.log(`StartDate: ${startDate} | EndDate: ${endDate}`);
  Logger.log(`Column: ${col}`);

  const pdUserId = getUserId_directly(engineerName);
  if (!pdUserId) {
    Logger.log(`PD user not found: ${engineerName}`);
    Logger.log("=== removeShiftsForEngineerDay END ===");
    return;
  }

  const dailyShifts = computeDailyShifts(startDate, endDate, col);
  Logger.log(`Generated ${dailyShifts.length} daily shifts for deletion.`);

  dailyShifts.forEach(({ startISO, endISO }, idx) => {
    Logger.log(`--- Checking Shift ${idx}: ${startISO} → ${endISO} ---`);

    const shifts = getOverrides(SCHEDULE_ID, new Date(startISO), new Date(endISO));
    if (!shifts || !shifts.length) {
      Logger.log(`No shifts found for ${engineerName} in this period.`);
      return;
    }

    shifts.forEach(s => {
      if (s.user && s.user.id === pdUserId) {
        Logger.log(`Deleting shift ${s.id} for ${engineerName}`);
        deleteShift(SCHEDULE_ID, s.id);
      }
    });

    Logger.log(`--- Shift ${idx} deletion check done ---`);
  });

  Logger.log("=== removeShiftsForEngineerDay END ===");
}

function processBoldEngineerCell(row, col, startDate, engineerName, endDate, suppressError = false) {
  Logger.log("=== processBoldEngineerCell START ===");
  Logger.log(`Input params -> row: ${row}, col: ${col}`);
  Logger.log(`Engineer: "${engineerName}"`);
  Logger.log(`StartDate: ${startDate} | EndDate: ${endDate}`);

  const pdUserId = getUserId_directly(engineerName);
  Logger.log(`PD User Lookup -> engineer="${engineerName}", result userId="${pdUserId}"`);

  if (!pdUserId) {
    const msg = `No PagerDuty user found for "${engineerName}". Skip.`;
    Logger.log(msg);
    if (!suppressError) throw new Error(msg);
    Logger.log("Suppressed error due to suppressError=true");
    return;
  }

  const dailyShifts = computeDailyShifts(startDate, endDate, col);
  Logger.log(`Generated ${dailyShifts.length} daily shifts:`);

  dailyShifts.forEach((s, idx) => {
    Logger.log(`  Shift[${idx}] = ${s.startISO} → ${s.endISO}`);
  });

  dailyShifts.forEach(({ startISO, endISO }, idx) => {
    Logger.log(`--- Creating Shift ${idx} ---`);
    Logger.log(`PD User: ${pdUserId}`);
    Logger.log(`Start: ${startISO}`);
    Logger.log(`End:   ${endISO}`);

    const now = new Date();
    if (new Date(endISO) < now) {
      Logger.log(`Skip Shift ${idx}: end time ${endISO} is already past.`);
      return;
    }

    const result = createShift(SCHEDULE_ID, pdUserId, startISO, endISO);
    Logger.log(`createShift() returned: ${JSON.stringify(result)}`);
    Logger.log(`Shift created for ${engineerName}, ${startISO} → ${endISO}`);
    Logger.log(`--- Shift ${idx} done ---`);
  });

  Logger.log("=== processBoldEngineerCell END ===");
}

function deleteShift(scheduleId, overrideId) {
  const url = `https://api.pagerduty.com/schedules/${scheduleId}/overrides/${overrideId}`;
  Logger.log(`Deleting shift override: schedule=${scheduleId}, override=${overrideId}`);

  const resp = UrlFetchApp.fetch(url, {
    method: "delete",
    headers: {
      "Authorization": "Token token=" + PD_TOKEN,
      "Accept": "application/vnd.pagerduty+json;version=2"
    },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  Logger.log(`Delete shift response code: ${code}`);

  if (code !== 200 && code !== 204) {
    Logger.log("Failed to delete shift: " + resp.getContentText());
    throw new Error("Failed to delete shift override: " + resp.getContentText());
  }

  Logger.log("Shift deleted successfully");
}

function test_deleteShift() {
  const testScheduleId = SCHEDULE_ID;
  const testOverrideId = "Q1HCH6GMTMBWJP";  // example override ID

  Logger.log("=== Running test_deleteShift ===");
  Logger.log(`Target schedule: ${testScheduleId}`);
  Logger.log(`Target override: ${testOverrideId}`);

  try {
    deleteShift(testScheduleId, testOverrideId);
    Logger.log("Test completed: Shift deleted successfully.");
  } catch (err) {
    Logger.log("Test failed: " + err.message);
  }
}

/*********************************************
 * HELPERS — ensure layer exists
 *********************************************/
function ensureLayerExistsFixed(scheduleId, layerName) {
  Logger.log(`ensureLayerExistsFixed called with scheduleId=${scheduleId}, layerName=${layerName}`);

  const url = `https://api.pagerduty.com/schedules/${scheduleId}`;
  const scheduleResp = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': `Token token=${PD_TOKEN}`,
      'Accept': 'application/vnd.pagerduty+json;version=2'
    },
    muteHttpExceptions: true
  });

  if (scheduleResp.getResponseCode() !== 200) {
    Logger.log(`Schedule fetch failed: ${scheduleResp.getContentText()}`);
    throw new Error(`Error fetching schedule: ${scheduleResp.getContentText()}`);
  }

  const scheduleData = JSON.parse(scheduleResp.getContentText());
  const layer = scheduleData.schedule.schedule_layers.find(l => l.name === layerName);

  if (!layer) {
    Logger.log(`Layer ${layerName} does not exist. Please create it manually in PagerDuty.`);
    throw new Error(`Layer ${layerName} not found`);
  }

  Logger.log(`Found layer ${layerName} with id=${layer.id}`);
  return layer.id;
}

/*********************************************
 * HELPERS — compute daily shifts
 *********************************************/
function computeDailyShifts(startDate, endDate, col) {
  const shifts = [];
  const current = new Date(startDate);

  while (current <= endDate) {
    const dayOfWeek = current.getDay(); // 0=Sunday, 6=Saturday
    if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Only Monday to Friday
      let startHour, endHour;
      if (PST_COLUMNS.includes(col)) {
        startHour = PST_SHIFT.start;
        endHour = PST_SHIFT.end;
      } else {
        startHour = CET_SHIFT.start;
        endHour = CET_SHIFT.end;
      }

      const dayStart = new Date(current);
      dayStart.setHours(startHour, 0, 0, 0);

      const dayEnd = new Date(current);
      dayEnd.setHours(endHour, 0, 0, 0);

      shifts.push({
        startISO: toPST(dayStart),
        endISO: toPST(dayEnd)
      });
    }

    // Move to next day
    current.setDate(current.getDate() + 1);
  }

  return shifts;
}

/*********************************************
 * CREATE SHIFT — API call
 *********************************************/
function createShift(scheduleId, userId, startISO, endISO) {
  const url = `https://api.pagerduty.com/schedules/${scheduleId}/overrides`;
  const payload = {
    override: {
      start: startISO,
      end: endISO,
      user: { id: userId, type: "user_reference" }
    }
  };

  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: pdHeaders(),
    payload: JSON.stringify(payload)
  });
}

/*********************************************
 * HELPER — convert date to PST ISO string
 *********************************************/
function toPST(date) {
  const pad = n => String(n).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth()+1)}-${pad(date.getDate())}`
       + `T${pad(date.getHours())}:00:00-08:00`;
}

/*********************************************
 * HELPER — PagerDuty API headers
 *********************************************/
function pdHeaders() {
  return {
    "Authorization": "Token token=" + PD_TOKEN,
    "Accept": "application/vnd.pagerduty+json;version=2"
  };
}
