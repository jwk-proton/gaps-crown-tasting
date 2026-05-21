/**
 * Gap's Crown Vineyard — Blind Tasting Backend
 * Google Apps Script — deployed as a web app (execute as Me, access: Anyone)
 *
 * SHEET STRUCTURE (auto-created if missing):
 * Row 1: headers
 * Rows 2+: one row per taster (upserted by taster name)
 *
 * Columns:
 *   taster | timestamp | rank_A…rank_J (10 cols) | notes_A…notes_J (10 cols)
 *   Total: 23 columns
 *
 * ENDPOINTS:
 *   POST  { taster, rankings: [{id, rank}], notes: [{id, notes}] }
 *         → upserts the taster's row, returns { ok: true }
 *
 *   GET   ?action=aggregate
 *         → returns { aggregate: [{id, avgRank, submissions}], tasters: [...] }
 *
 * DEPLOYMENT:
 *   1. Open script.google.com → New project → paste this file
 *   2. Deploy → New deployment → Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   3. Copy the deployment URL into SCRIPT_URL in index.html
 */

// ── Column layout ──────────────────────────────────────────────────────────
var WINE_IDS   = ['A','B','C','D','E','F','G','H','I','J'];
var COL_TASTER = 1;   // column index (1-based) for taster name
var COL_TS     = 2;   // timestamp
var COL_RANK_A = 3;   // ranks: cols 3–12
var COL_NOTE_A = 13;  // notes: cols 13–22
// Total columns: 22

var SHEET_NAME = 'Tasting';


// ── Helpers ────────────────────────────────────────────────────────────────

/**
 * Returns the active spreadsheet's tasting sheet,
 * creating it (with headers) if it doesn't exist yet.
 */
function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    // Build header row
    var headers = ['taster', 'timestamp'];
    WINE_IDS.forEach(function(id) { headers.push('rank_' + id); });
    WINE_IDS.forEach(function(id) { headers.push('notes_' + id); });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Finds the row number (1-based) for a given taster name,
 * or returns -1 if not found. Skips the header row.
 */
function findTasterRow(sheet, tasterName) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === tasterName) return i + 1; // convert to 1-based row
  }
  return -1;
}

/**
 * Adds CORS headers to every response so the app can call from GitHub Pages.
 * Returns a JSON ContentService output with the given payload.
 */
function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── POST handler ───────────────────────────────────────────────────────────

/**
 * Receives a taster's full ranking + notes payload and upserts their row.
 *
 * Expected body (JSON string):
 * {
 *   taster:   "Jason",
 *   rankings: [ { id: "A", rank: 1 }, … ],   // all 10 wines
 *   notes:    [ { id: "A", notes: "…" }, … ]  // all 10 wines
 * }
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var tasterName = (payload.taster || '').trim();
    if (!tasterName) return jsonResponse({ ok: false, error: 'Missing taster name' });

    var sheet     = getSheet();
    var timestamp = new Date().toISOString();

    // Build the row values array (22 columns)
    var row = new Array(22).fill('');
    row[0] = tasterName;
    row[1] = timestamp;

    // Fill ranks (cols 3–12, index 2–11)
    (payload.rankings || []).forEach(function(r) {
      var idx = WINE_IDS.indexOf(r.id);
      if (idx >= 0) row[2 + idx] = r.rank;
    });

    // Fill notes (cols 13–22, index 12–21)
    (payload.notes || []).forEach(function(n) {
      var idx = WINE_IDS.indexOf(n.id);
      if (idx >= 0) row[12 + idx] = n.notes || '';
    });

    // Upsert: overwrite existing row or append new one
    var existingRow = findTasterRow(sheet, tasterName);
    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return jsonResponse({ ok: true, taster: tasterName, timestamp: timestamp });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}


// ── GET handler ────────────────────────────────────────────────────────────

/**
 * Returns aggregate rankings across all tasters who have submitted.
 *
 * Response:
 * {
 *   ok: true,
 *   tasters: ["Jason", "Katie", …],          // who has submitted
 *   aggregate: [
 *     { id: "A", avgRank: 2.3, submissions: 4 },
 *     …
 *   ]   // sorted by avgRank ascending (lower = better)
 * }
 */
function doGet(e) {
  try {
    var sheet = getSheet();
    var data  = sheet.getDataRange().getValues();

    // Rows 2+ are taster data (row 1 is headers)
    var tasterRows = data.slice(1).filter(function(row) {
      return row[0] && row[0].toString().trim() !== '';
    });

    var tasters = tasterRows.map(function(row) { return row[0]; });

    // Accumulate rank totals per wine
    var totals      = {};  // { wineId: total rank sum }
    var counts      = {};  // { wineId: number of submissions with a rank }
    WINE_IDS.forEach(function(id) { totals[id] = 0; counts[id] = 0; });

    tasterRows.forEach(function(row) {
      WINE_IDS.forEach(function(id, i) {
        var rank = parseFloat(row[2 + i]);  // COL_RANK_A - 1 (0-indexed)
        if (!isNaN(rank) && rank > 0) {
          totals[id] += rank;
          counts[id]++;
        }
      });
    });

    // Build aggregate array
    var aggregate = WINE_IDS.map(function(id) {
      var n = counts[id];
      return {
        id:          id,
        avgRank:     n > 0 ? Math.round((totals[id] / n) * 100) / 100 : 999,
        submissions: n
      };
    });

    // Sort by average rank ascending (best wines first)
    aggregate.sort(function(a, b) { return a.avgRank - b.avgRank; });

    return jsonResponse({ ok: true, tasters: tasters, aggregate: aggregate });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}
