/**
 * Gap's Crown Vineyard — Blind Tasting Backend
 * Google Apps Script — deployed as a web app (execute as Me, access: Anyone)
 *
 * SHEET STRUCTURE (auto-created if missing):
 * Col 1:     taster
 * Col 2:     timestamp
 * Cols 3-12: rank_A…rank_J  (flight ranks, 1-5 per flight)
 * Cols 13-22: notes_A…notes_J
 * Cols 23-32: fin_rank_A…fin_rank_J (finalist personal ranking, 1-4, sparse)
 * Total: 32 columns
 *
 * ENDPOINTS:
 *   POST  { taster, rankings, notes, finalistRankings }
 *         → upserts taster row
 *
 *   GET   ?action=aggregate
 *         → top 2 from each flight by avg rank
 *
 *   GET   ?action=taster&name=Jason
 *         → that taster's saved record
 *
 *   GET   ?action=results
 *         → group finalist order + each taster's personal finalist ranking
 */

var FLIGHT_0  = ['A','B','C','D','E'];
var FLIGHT_1  = ['F','G','H','I','J'];
var WINE_IDS  = FLIGHT_0.concat(FLIGHT_1);
var SHEET_NAME = 'Tasting';

// Column offsets (1-based)
var COL_TASTER    = 1;
var COL_TS        = 2;
var COL_RANK_0    = 3;   // rank_A … rank_J  → cols 3-12
var COL_NOTES_0   = 13;  // notes_A … notes_J → cols 13-22
var COL_FIN_0     = 23;  // fin_rank_A … fin_rank_J → cols 23-32
var TOTAL_COLS    = 32;


// ── Helpers ────────────────────────────────────────────────────────────────

function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var headers = ['taster', 'timestamp'];
    WINE_IDS.forEach(function(id) { headers.push('rank_' + id); });
    WINE_IDS.forEach(function(id) { headers.push('notes_' + id); });
    WINE_IDS.forEach(function(id) { headers.push('fin_rank_' + id); });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findTasterRow(sheet, name) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === name) return i + 1;
  }
  return -1;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── POST ───────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var payload    = JSON.parse(e.postData.contents);
    var tasterName = (payload.taster || '').trim();
    if (!tasterName) return jsonResponse({ ok: false, error: 'Missing taster name' });

    var sheet     = getSheet();
    var timestamp = new Date().toISOString();
    var row       = new Array(TOTAL_COLS).fill('');

    row[0] = tasterName;
    row[1] = timestamp;

    // Flight rankings (cols 3-12)
    (payload.rankings || []).forEach(function(r) {
      var idx = WINE_IDS.indexOf(r.id);
      if (idx >= 0) row[COL_RANK_0 - 1 + idx] = r.rank;
    });

    // Notes (cols 13-22)
    (payload.notes || []).forEach(function(n) {
      var idx = WINE_IDS.indexOf(n.id);
      if (idx >= 0) row[COL_NOTES_0 - 1 + idx] = n.notes || '';
    });

    // Finalist personal rankings (cols 23-32, sparse)
    (payload.finalistRankings || []).forEach(function(f) {
      var idx = WINE_IDS.indexOf(f.id);
      if (idx >= 0) row[COL_FIN_0 - 1 + idx] = f.rank;
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


// ── GET router ─────────────────────────────────────────────────────────────

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'aggregate';
  if (action === 'taster')    return doGetTaster(e);
  if (action === 'results')   return doGetResults(e);
  return doGetAggregate(e);
}


// ── GET ?action=taster ─────────────────────────────────────────────────────

function doGetTaster(e) {
  try {
    var name = (e && e.parameter && e.parameter.name) ? e.parameter.name.trim() : '';
    if (!name) return jsonResponse({ ok: false, error: 'Missing name' });

    var sheet  = getSheet();
    var rowNum = findTasterRow(sheet, name);
    if (rowNum < 0) return jsonResponse({ ok: true, found: false });

    var row = sheet.getRange(rowNum, 1, 1, TOTAL_COLS).getValues()[0];

    var wines = WINE_IDS.map(function(id, i) {
      var flight = FLIGHT_0.indexOf(id) >= 0 ? 0 : 1;
      var rank   = parseFloat(row[COL_RANK_0 - 1 + i]);
      return {
        id:     id,
        flight: flight,
        rank:   (!isNaN(rank) && rank > 0) ? rank : (flight === 0 ? i + 1 : i - 4),
        notes:  row[COL_NOTES_0 - 1 + i] ? row[COL_NOTES_0 - 1 + i].toString() : ''
      };
    });

    // Restore finalist personal rankings if they exist
    var finalistRankings = [];
    WINE_IDS.forEach(function(id, i) {
      var fr = parseFloat(row[COL_FIN_0 - 1 + i]);
      if (!isNaN(fr) && fr > 0) {
        finalistRankings.push({ id: id, rank: fr });
      }
    });

    return jsonResponse({ ok: true, found: true, wines: wines, finalistRankings: finalistRankings });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}


// ── GET ?action=aggregate ──────────────────────────────────────────────────

function doGetAggregate(e) {
  try {
    var sheet = getSheet();
    var data  = sheet.getDataRange().getValues();

    var tasterRows = data.slice(1).filter(function(r) {
      return r[0] && r[0].toString().trim() !== '';
    });

    var tasters = tasterRows.map(function(r) { return r[0]; });
    var totals = {}, counts = {};
    WINE_IDS.forEach(function(id) { totals[id] = 0; counts[id] = 0; });

    tasterRows.forEach(function(row) {
      WINE_IDS.forEach(function(id, i) {
        var rank = parseFloat(row[COL_RANK_0 - 1 + i]);
        if (!isNaN(rank) && rank > 0) { totals[id] += rank; counts[id]++; }
      });
    });

    function flightAgg(ids, flightNum) {
      return ids.map(function(id) {
        var n = counts[id];
        return { id: id, flight: flightNum, avgRank: n > 0 ? Math.round(totals[id]/n*100)/100 : 999, submissions: n };
      }).sort(function(a, b) { return a.avgRank - b.avgRank; });
    }

    var finalists = flightAgg(FLIGHT_0, 0).slice(0,2).concat(flightAgg(FLIGHT_1, 1).slice(0,2));
    return jsonResponse({ ok: true, tasters: tasters, finalists: finalists });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}


// ── GET ?action=results ────────────────────────────────────────────────────

/**
 * Returns:
 * {
 *   ok: true,
 *   finalists: [{id, flight, avgRank, submissions}],  // group order, top 2 per flight
 *   tasterRankings: {
 *     "Jason":  { "E": 1, "D": 2, "J": 3, "H": 4 },
 *     "Katie":  { "D": 1, "E": 2, "H": 3, "J": 4 },
 *     …
 *   }
 * }
 */
function doGetResults(e) {
  try {
    var sheet = getSheet();
    var data  = sheet.getDataRange().getValues();

    var tasterRows = data.slice(1).filter(function(r) {
      return r[0] && r[0].toString().trim() !== '';
    });

    // Compute group finalists (same as aggregate)
    var totals = {}, counts = {};
    WINE_IDS.forEach(function(id) { totals[id] = 0; counts[id] = 0; });
    tasterRows.forEach(function(row) {
      WINE_IDS.forEach(function(id, i) {
        var rank = parseFloat(row[COL_RANK_0 - 1 + i]);
        if (!isNaN(rank) && rank > 0) { totals[id] += rank; counts[id]++; }
      });
    });

    function flightAgg(ids, flightNum) {
      return ids.map(function(id) {
        var n = counts[id];
        return { id: id, flight: flightNum, avgRank: n > 0 ? Math.round(totals[id]/n*100)/100 : 999, submissions: n };
      }).sort(function(a, b) { return a.avgRank - b.avgRank; });
    }

    var finalists = flightAgg(FLIGHT_0, 0).slice(0,2).concat(flightAgg(FLIGHT_1, 1).slice(0,2));

    // Collect each taster's personal finalist rankings
    var tasterRankings = {};
    tasterRows.forEach(function(row) {
      var name    = row[0].toString();
      var rankings = {};
      WINE_IDS.forEach(function(id, i) {
        var fr = parseFloat(row[COL_FIN_0 - 1 + i]);
        if (!isNaN(fr) && fr > 0) rankings[id] = fr;
      });
      tasterRankings[name] = rankings;
    });

    return jsonResponse({ ok: true, finalists: finalists, tasterRankings: tasterRankings });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}
