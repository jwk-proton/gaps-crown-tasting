/**
 * Gap's Crown Vineyard — Blind Tasting Backend
 * Google Apps Script — deployed as a web app (execute as Me, access: Anyone)
 *
 * SHEET STRUCTURE (auto-created if missing):
 * Row 1: headers
 * Rows 2+: one row per taster (upserted by taster name)
 *
 * Columns:
 *   taster | timestamp | rank_A…rank_E (Flight I, 1-5) | rank_F…rank_J (Flight II, 1-5) | notes_A…notes_J
 *   Total: 23 columns
 *
 * ENDPOINTS:
 *   POST  { taster, rankings: [{id, rank, flight}], notes: [{id, notes}] }
 *         → upserts the taster's row, returns { ok: true }
 *
 *   GET   ?action=aggregate
 *         → returns top 2 from Flight I + top 2 from Flight II by average rank
 *           { ok: true, tasters: [...], finalists: [{id, flight, avgRank, submissions}] }
 */

var FLIGHT_0 = ['A','B','C','D','E'];  // Flight I wines
var FLIGHT_1 = ['F','G','H','I','J'];  // Flight II wines
var WINE_IDS = FLIGHT_0.concat(FLIGHT_1);

var COL_TASTER = 1;   // 1-based
var COL_TS     = 2;
var COL_RANK_A = 3;   // ranks A-E: cols 3-7
var COL_RANK_F = 8;   // ranks F-J: cols 8-12
var COL_NOTE_A = 13;  // notes A-J: cols 13-22

var SHEET_NAME = 'Tasting';


// ── Helpers ────────────────────────────────────────────────────────────────

function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var headers = ['taster', 'timestamp'];
    WINE_IDS.forEach(function(id) { headers.push('rank_' + id); });
    WINE_IDS.forEach(function(id) { headers.push('notes_' + id); });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function findTasterRow(sheet, tasterName) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === tasterName) return i + 1;
  }
  return -1;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── POST handler ───────────────────────────────────────────────────────────

/**
 * Receives rankings as 1-5 per flight (both flights use 1-5 independently).
 * The flight membership is determined by wine ID (A-E = Flight I, F-J = Flight II).
 */
function doPost(e) {
  try {
    var payload    = JSON.parse(e.postData.contents);
    var tasterName = (payload.taster || '').trim();
    if (!tasterName) return jsonResponse({ ok: false, error: 'Missing taster name' });

    var sheet     = getSheet();
    var timestamp = new Date().toISOString();

    var row = new Array(22).fill('');
    row[0] = tasterName;
    row[1] = timestamp;

    // Ranks (cols 3-12, index 2-11) — stored as 1-5 per flight
    (payload.rankings || []).forEach(function(r) {
      var idx = WINE_IDS.indexOf(r.id);
      if (idx >= 0) row[2 + idx] = r.rank;
    });

    // Notes (cols 13-22, index 12-21)
    (payload.notes || []).forEach(function(n) {
      var idx = WINE_IDS.indexOf(n.id);
      if (idx >= 0) row[12 + idx] = n.notes || '';
    });

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
 * Computes average rank per wine within each flight independently,
 * then returns the top 2 from Flight I and top 2 from Flight II.
 *
 * Response:
 * {
 *   ok: true,
 *   tasters: ["Jason", "Katie", …],
 *   finalists: [
 *     { id: "E", flight: 0, avgRank: 1.5, submissions: 4 },
 *     { id: "A", flight: 0, avgRank: 2.0, submissions: 4 },
 *     { id: "J", flight: 1, avgRank: 1.0, submissions: 4 },
 *     { id: "H", flight: 1, avgRank: 2.5, submissions: 4 }
 *   ]
 * }
 */
function doGet(e) {
  try {
    var sheet = getSheet();
    var data  = sheet.getDataRange().getValues();

    var tasterRows = data.slice(1).filter(function(row) {
      return row[0] && row[0].toString().trim() !== '';
    });

    var tasters = tasterRows.map(function(row) { return row[0]; });

    // Accumulate per-wine totals and counts
    var totals = {}, counts = {};
    WINE_IDS.forEach(function(id) { totals[id] = 0; counts[id] = 0; });

    tasterRows.forEach(function(row) {
      WINE_IDS.forEach(function(id, i) {
        var rank = parseFloat(row[2 + i]);  // cols 3-12 → index 2-11
        if (!isNaN(rank) && rank > 0) {
          totals[id] += rank;
          counts[id]++;
        }
      });
    });

    // Build per-flight aggregate arrays
    function flightAggregate(ids, flightNum) {
      return ids.map(function(id) {
        var n = counts[id];
        return {
          id:          id,
          flight:      flightNum,
          avgRank:     n > 0 ? Math.round((totals[id] / n) * 100) / 100 : 999,
          submissions: n
        };
      }).sort(function(a, b) { return a.avgRank - b.avgRank; });
    }

    var flight0 = flightAggregate(FLIGHT_0, 0);
    var flight1 = flightAggregate(FLIGHT_1, 1);

    // Top 2 from each flight
    var finalists = flight0.slice(0, 2).concat(flight1.slice(0, 2));

    return jsonResponse({ ok: true, tasters: tasters, finalists: finalists });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}
