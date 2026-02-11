/**
 * Code.gs — Weekly Tipout Sheet Script
 * The Local on 50th
 *
 * Receives daily tip data from the calculator app and writes it
 * to weekly sheets in the "Weekly Tipouts" Google Spreadsheet.
 *
 * NEW SHEET LAYOUT (for all newly created weekly sheets):
 *
 *   A: Employee Name
 *   B: Mon Tips  |  C: Mon Hrs
 *   D: Tue Tips  |  E: Tue Hrs
 *   F: Wed Tips  |  G: Wed Hrs
 *   H: Thu Tips  |  I: Thu Hrs
 *   J: Fri Tips  |  K: Fri Hrs
 *   L: Sat Tips  |  M: Sat Hrs
 *   N: Sun Tips  |  O: Sun Hrs
 *   P: Weekly Tips Total  |  Q: Weekly Hours Total
 *   R: (spacer)
 *   S onward: Daily Summary
 *
 * Hosts do NOT get hours — only tips (split evenly).
 * Bussers/Expo, Kitchen (BOH), and Servers/Bartenders have hours.
 *
 * Existing/prior-week sheets are NOT modified. If data is sent for
 * a date that falls in an already-existing old-format sheet, the
 * script detects the old layout and writes using old column positions.
 */

// ═══════════════════════════════════════════════════════
//  COLUMN CONSTANTS — NEW FORMAT (tips + hours)
// ═══════════════════════════════════════════════════════

var NAME_COL         = 1;   // A

var MON_TIPS_COL     = 2;   // B
var MON_HRS_COL      = 3;   // C
var TUE_TIPS_COL     = 4;   // D
var TUE_HRS_COL      = 5;   // E
var WED_TIPS_COL     = 6;   // F
var WED_HRS_COL      = 7;   // G
var THU_TIPS_COL     = 8;   // H
var THU_HRS_COL      = 9;   // I
var FRI_TIPS_COL     = 10;  // J
var FRI_HRS_COL      = 11;  // K
var SAT_TIPS_COL     = 12;  // L
var SAT_HRS_COL      = 13;  // M
var SUN_TIPS_COL     = 14;  // N
var SUN_HRS_COL      = 15;  // O

var WEEKLY_TIPS_COL  = 16;  // P
var WEEKLY_HRS_COL   = 17;  // Q
// Column 18 (R) is a spacer
var SUMMARY_START_COL = 19; // S

// ═══════════════════════════════════════════════════════
//  COLUMN CONSTANTS — OLD FORMAT (backward compatibility)
// ═══════════════════════════════════════════════════════

var OLD_MONDAY_COL        = 2;   // B  (Mon=B … Sun=H)
var OLD_WEEKLY_TOTAL_COL  = 9;   // I
var OLD_SUMMARY_START_COL = 12;  // L

// ═══════════════════════════════════════════════════════
//  SHARED CONSTANTS
// ═══════════════════════════════════════════════════════

var HEADER_ROW     = 1;
var DATA_START_ROW = 2;
var DAY_NAMES      = ['Sunday', 'Monday', 'Tuesday', 'Wednesday',
                      'Thursday', 'Friday', 'Saturday'];

// ═══════════════════════════════════════════════════════
//  WEB-APP ENTRY POINTS
// ═══════════════════════════════════════════════════════

function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var result = processTipData(data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    message: 'Weekly Tipout Script is running'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════
//  MAIN PROCESSING
// ═══════════════════════════════════════════════════════

/**
 * Receives a daily payload from the calculator and writes employee
 * tips (and hours, for new-format sheets) to the correct day columns.
 *
 * Payload shape from calculator:
 *   { date, shiftLead, totalSales, toGoSales, tipOutSales,
 *     ccTips, cashTips, autoGrat, giftCardTips, totalTips,
 *     employees: [{ name, role, hours, tips }] }
 */
function processTipData(data) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var date     = new Date(data.date + 'T12:00:00');
  var jsDay    = date.getDay();           // 0 = Sun … 6 = Sat
  var dayIndex = getDayIndex(jsDay);      // 0 = Mon … 6 = Sun

  // Find or create the weekly sheet
  var sheet  = getOrCreateWeeklySheet(ss, date);
  var newFmt = isNewFormat(sheet);

  // Determine which columns to write to for this day of the week
  var tipsCol, hrsCol;
  if (newFmt) {
    tipsCol = MON_TIPS_COL + dayIndex * 2;  // B, D, F, H, J, L, N
    hrsCol  = tipsCol + 1;                   // C, E, G, I, K, M, O
  } else {
    tipsCol = OLD_MONDAY_COL + dayIndex;     // B … H
    hrsCol  = null;                          // old format has no hours columns
  }

  // Write each employee's data
  var employees = data.employees || [];
  for (var i = 0; i < employees.length; i++) {
    var emp = employees[i];
    var row = findOrCreateEmployeeRow(sheet, emp.name);

    // Always write tips
    sheet.getRange(row, tipsCol).setValue(emp.tips);

    // Write hours only on new-format sheets, and only for non-Host roles
    if (newFmt && hrsCol && emp.role !== 'Host') {
      sheet.getRange(row, hrsCol).setValue(emp.hours);
    }
  }

  // Update weekly total formulas
  updateWeeklyTotals(sheet, newFmt);

  // Update the daily summary section
  updateDailySummary(sheet, data, jsDay, newFmt);

  // Apply number formatting
  formatSheet(sheet, newFmt);

  return { success: true, sheet: sheet.getName() };
}

// ═══════════════════════════════════════════════════════
//  SHEET LOOKUP / CREATION
// ═══════════════════════════════════════════════════════

/**
 * Returns the weekly sheet for the week containing `date`.
 * Creates a new sheet (with the new column layout) if one doesn't exist.
 */
function getOrCreateWeeklySheet(ss, date) {
  var weekStart = getWeekStart(date);
  var weekEnd   = new Date(weekStart);
  weekEnd.setDate(weekEnd.getDate() + 6);

  var name  = formatSheetName(weekStart, weekEnd);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = createNewWeeklySheet(ss, name);
  }
  return sheet;
}

/** Returns the Monday of the week containing `date`. */
function getWeekStart(date) {
  var d    = new Date(date);
  var day  = d.getDay();                       // 0 = Sun
  var diff = (day === 0) ? 6 : (day - 1);     // go back to Monday
  d.setDate(d.getDate() - diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

/** Sheet name: "M/D - M/D" (Monday – Sunday). */
function formatSheetName(weekStart, weekEnd) {
  var s = (weekStart.getMonth() + 1) + '/' + weekStart.getDate();
  var e = (weekEnd.getMonth()   + 1) + '/' + weekEnd.getDate();
  return s + ' - ' + e;
}

/**
 * Creates a brand-new weekly sheet with the NEW column structure
 * (tips + hours for each day). Only new sheets get this layout —
 * existing sheets from prior weeks are never restructured.
 */
function createNewWeeklySheet(ss, name) {
  var sheet = ss.insertSheet(name);

  // ── Main headers (A through Q) ──
  var mainHeaders = [
    'Employee Name',
    'Mon Tips', 'Mon Hrs',
    'Tue Tips', 'Tue Hrs',
    'Wed Tips', 'Wed Hrs',
    'Thu Tips', 'Thu Hrs',
    'Fri Tips', 'Fri Hrs',
    'Sat Tips', 'Sat Hrs',
    'Sun Tips', 'Sun Hrs',
    'Weekly Tips', 'Weekly Hrs'
  ];
  sheet.getRange(HEADER_ROW, 1, 1, mainHeaders.length).setValues([mainHeaders]);

  // ── Daily Summary headers (S onward) ──
  var summaryHeaders = [
    'Date', 'Day', 'Shift Lead',
    'Total Sales', 'To-Go Sales', 'Tip-Out Sales',
    'CC Tips', 'Cash Tips', 'Auto Grat', 'Gift Card Tips',
    'Total Tips'
  ];
  sheet.getRange(HEADER_ROW, SUMMARY_START_COL, 1, summaryHeaders.length)
       .setValues([summaryHeaders]);

  // ── Style the header row ──
  var lastCol = SUMMARY_START_COL + summaryHeaders.length - 1;
  var hdr     = sheet.getRange(HEADER_ROW, 1, 1, lastCol);
  hdr.setFontWeight('bold')
     .setBackground('#4a5568')
     .setFontColor('#ffffff')
     .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  // ── Column widths ──
  sheet.setColumnWidth(NAME_COL, 180);
  for (var c = MON_TIPS_COL; c <= SUN_HRS_COL; c++) {
    sheet.setColumnWidth(c, 85);
  }
  sheet.setColumnWidth(WEEKLY_TIPS_COL, 105);
  sheet.setColumnWidth(WEEKLY_HRS_COL, 105);
  sheet.setColumnWidth(18, 20); // spacer column R

  return sheet;
}

// ═══════════════════════════════════════════════════════
//  FORMAT DETECTION
// ═══════════════════════════════════════════════════════

/**
 * Returns true if the sheet uses the new Tips + Hours column layout.
 * Detection: cell B1 will be "Mon Tips" in new format vs "Mon" in old.
 */
function isNewFormat(sheet) {
  return sheet.getRange(HEADER_ROW, 2).getValue() === 'Mon Tips';
}

// ═══════════════════════════════════════════════════════
//  EMPLOYEE ROW MANAGEMENT
// ═══════════════════════════════════════════════════════

/**
 * Finds the row for an existing employee by name, or appends a
 * new row at the bottom of the sheet and returns its row number.
 */
function findOrCreateEmployeeRow(sheet, name) {
  var lastRow = Math.max(sheet.getLastRow(), HEADER_ROW);

  if (lastRow >= DATA_START_ROW) {
    var names = sheet.getRange(DATA_START_ROW, NAME_COL,
                    lastRow - HEADER_ROW, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (names[i][0] === name) return i + DATA_START_ROW;
    }
  }

  var newRow = lastRow + 1;
  sheet.getRange(newRow, NAME_COL).setValue(name);
  return newRow;
}

// ═══════════════════════════════════════════════════════
//  WEEKLY TOTAL FORMULAS
// ═══════════════════════════════════════════════════════

/**
 * Sets SUM formulas in the weekly total columns for every employee row.
 *
 * New format:
 *   Weekly Tips  (P) = SUM of tip  columns B, D, F, H, J, L, N
 *   Weekly Hours (Q) = SUM of hour columns C, E, G, I, K, M, O
 *
 * Old format:
 *   Weekly Total (I) = SUM(B:H)
 */
function updateWeeklyTotals(sheet, newFmt) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  for (var row = DATA_START_ROW; row <= lastRow; row++) {
    if (!sheet.getRange(row, NAME_COL).getValue()) continue;

    if (newFmt) {
      var tf = '=SUM(' +
        colLetter(MON_TIPS_COL) + row + ',' +
        colLetter(TUE_TIPS_COL) + row + ',' +
        colLetter(WED_TIPS_COL) + row + ',' +
        colLetter(THU_TIPS_COL) + row + ',' +
        colLetter(FRI_TIPS_COL) + row + ',' +
        colLetter(SAT_TIPS_COL) + row + ',' +
        colLetter(SUN_TIPS_COL) + row + ')';

      var hf = '=SUM(' +
        colLetter(MON_HRS_COL) + row + ',' +
        colLetter(TUE_HRS_COL) + row + ',' +
        colLetter(WED_HRS_COL) + row + ',' +
        colLetter(THU_HRS_COL) + row + ',' +
        colLetter(FRI_HRS_COL) + row + ',' +
        colLetter(SAT_HRS_COL) + row + ',' +
        colLetter(SUN_HRS_COL) + row + ')';

      sheet.getRange(row, WEEKLY_TIPS_COL).setFormula(tf);
      sheet.getRange(row, WEEKLY_HRS_COL).setFormula(hf);
    } else {
      var oldFormula = '=SUM(' +
        colLetter(OLD_MONDAY_COL) + row + ':' +
        colLetter(OLD_MONDAY_COL + 6) + row + ')';
      sheet.getRange(row, OLD_WEEKLY_TOTAL_COL).setFormula(oldFormula);
    }
  }
}

// ═══════════════════════════════════════════════════════
//  DAILY SUMMARY
// ═══════════════════════════════════════════════════════

/**
 * Writes (or overwrites) a row in the Daily Summary section for
 * the submitted date.  New-format sheets start the summary at
 * column S; old-format sheets start at column L.
 */
function updateDailySummary(sheet, data, jsDay, newFmt) {
  var summaryCol = newFmt ? SUMMARY_START_COL : OLD_SUMMARY_START_COL;
  var lastRow    = sheet.getLastRow();
  var targetRow  = DATA_START_ROW;

  // Look for existing entry for this date, or find the next empty summary row
  if (lastRow >= DATA_START_ROW) {
    var dates = sheet.getRange(DATA_START_ROW, summaryCol,
                    lastRow - HEADER_ROW, 1).getValues();
    var found = false;

    for (var i = 0; i < dates.length; i++) {
      if (dates[i][0] === data.date) {
        targetRow = DATA_START_ROW + i;
        found = true;
        break;
      }
    }

    if (!found) {
      // Place after the last non-empty summary row
      for (var j = dates.length - 1; j >= 0; j--) {
        if (dates[j][0] !== '' && dates[j][0] !== null) {
          targetRow = DATA_START_ROW + j + 1;
          break;
        }
      }
    }
  }

  var summaryRow = [
    data.date,
    DAY_NAMES[jsDay],
    data.shiftLead    || '',
    data.totalSales   || 0,
    data.toGoSales    || 0,
    data.tipOutSales  || 0,
    data.ccTips       || 0,
    data.cashTips     || 0,
    data.autoGrat     || 0,
    data.giftCardTips || 0,
    data.totalTips    || 0
  ];

  sheet.getRange(targetRow, summaryCol, 1, summaryRow.length)
       .setValues([summaryRow]);
}

// ═══════════════════════════════════════════════════════
//  FORMATTING
// ═══════════════════════════════════════════════════════

/**
 * Applies number formats and alignment to the sheet.
 *
 * New format:
 *   Tips columns  → $#,##0.00
 *   Hours columns → 0.00"/hr"  (displays as e.g. "5.75/hr")
 *   Summary $     → $#,##0.00
 *
 * Old format:
 *   Day columns + weekly total → $#,##0.00
 *   Summary $                  → $#,##0.00
 */
function formatSheet(sheet, newFmt) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;
  var rows = lastRow - DATA_START_ROW + 1;

  if (newFmt) {
    // ── Tips columns: currency ──
    var tipCols = [MON_TIPS_COL, TUE_TIPS_COL, WED_TIPS_COL, THU_TIPS_COL,
                   FRI_TIPS_COL, SAT_TIPS_COL, SUN_TIPS_COL, WEEKLY_TIPS_COL];
    for (var t = 0; t < tipCols.length; t++) {
      sheet.getRange(DATA_START_ROW, tipCols[t], rows, 1)
           .setNumberFormat('$#,##0.00');
    }

    // ── Hours columns: number with "/hr" suffix ──
    var hrCols = [MON_HRS_COL, TUE_HRS_COL, WED_HRS_COL, THU_HRS_COL,
                  FRI_HRS_COL, SAT_HRS_COL, SUN_HRS_COL, WEEKLY_HRS_COL];
    for (var h = 0; h < hrCols.length; h++) {
      sheet.getRange(DATA_START_ROW, hrCols[h], rows, 1)
           .setNumberFormat('0.00"/hr"');
    }

    // ── Summary dollar columns ──
    // Offsets 3–10 from SUMMARY_START_COL:
    //   +3 Total Sales, +4 To-Go Sales, +5 Tip-Out Sales,
    //   +6 CC Tips, +7 Cash Tips, +8 Auto Grat, +9 Gift Card Tips,
    //   +10 Total Tips
    for (var s = 3; s <= 10; s++) {
      sheet.getRange(DATA_START_ROW, SUMMARY_START_COL + s, rows, 1)
           .setNumberFormat('$#,##0.00');
    }

    // ── Center-align the data area (B through Q) ──
    sheet.getRange(DATA_START_ROW, MON_TIPS_COL, rows,
                   WEEKLY_HRS_COL - MON_TIPS_COL + 1)
         .setHorizontalAlignment('center');

  } else {
    // Old format: currency for B through I
    sheet.getRange(DATA_START_ROW, OLD_MONDAY_COL, rows, 8)
         .setNumberFormat('$#,##0.00');

    // Old summary dollar columns
    for (var o = 3; o <= 10; o++) {
      sheet.getRange(DATA_START_ROW, OLD_SUMMARY_START_COL + o, rows, 1)
           .setNumberFormat('$#,##0.00');
    }
  }

  // ── Alternating row colors for readability ──
  var endCol = newFmt ? WEEKLY_HRS_COL : OLD_WEEKLY_TOTAL_COL;
  for (var r = DATA_START_ROW; r <= lastRow; r++) {
    var bg = ((r - DATA_START_ROW) % 2 === 1) ? '#f7fafc' : '#ffffff';
    sheet.getRange(r, 1, 1, endCol).setBackground(bg);
  }
}

// ═══════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════

/**
 * Converts JS getDay() (0=Sun, 1=Mon … 6=Sat) to our
 * 0-based day index where 0=Monday … 6=Sunday.
 */
function getDayIndex(jsDay) {
  return (jsDay === 0) ? 6 : jsDay - 1;
}

/**
 * Converts a 1-based column number to a column letter.
 * 1→A, 2→B, … 26→Z, 27→AA, etc.
 */
function colLetter(n) {
  var s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}
