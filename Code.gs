/**
 * Gap's Crown Vineyard — Blind Tasting Backend
 * Google Apps Script — deployed as a web app (execute as Me, access: Anyone)
 *
 * SHEET STRUCTURE (auto-created if missing):
 * Row 1: headers
 * Rows 2+: one row per taster (upserted by taster name)
 *
 * Columns:
 *   taster | timestamp | rank_A…rank_E (1-5) | rank_F…rank_J (1-5) | notes_A…notes_J
 *   Total: 23 columns
 *
 * ENDPOINTS:
 *   POST  { taster, rankings: [{id, rank}], notes: [{id, notes}] }
 *         → upserts taster row, returns { ok: true }
 *
 *   GET   ?action=aggregate
 *         → top 2 from Flight I + top 2 from Flight II by avg rank
 *           { ok: true, tasters: [...], finalists: [{id, flight, avgRank, submissions}] }
 *
 *   GET   ?action=taster&name=Jason
 *         → returns that taster's saved rankings + notes, or found:false if not in sheet
 *           { ok: true, found: true, wines: [...], finalists: [] }
 *           { ok: true, found: false }
 */

var FLIGHT_0 = ['A','B','C','D','E'];
var FLIGHT_1 = ['F','G','H','I','J'];
var WINE_IDS = FLIGHT_0.concat(FLIGHT_1);

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

    (payload.rankings || []).forEach(function(r) {
      var idx = WINE_IDS.indexOf(r.id);
      if (idx >= 0) row[2 + idx] = r.rank;
    });

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

function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? e.parameter.action : 'aggregate';

  if (action === 'taster') {
    return doGetTaster(e);
  } else {
    return doGetAggregate(e);
  }
}

/**
 * Returns a single taster's saved record, reconstructed into the wines array
 * format the app expects.
 *
 * Response (found):
 * {
 *   ok: true, found: true,
 *   wines: [ {id:'A', flight:0, rank:1, notes:'...'}, … ],
 *   finalists: []   // finalists are always re-derived from group data on load
 * }
 *
 * Response (not found):
 * { ok: true, found: false }
 */
function doGetTaster(e) {
  try {
    var name = e && e.parameter && e.parameter.name ? e.parameter.name.trim() : '';
    if (!name) return jsonResponse({ ok: false, error: 'Missing name parameter' });

    var sheet = getSheet();
    var rowNum = findTasterRow(sheet, name);

    if (rowNum < 0) {
      return jsonResponse({ ok: true, found: false });
    }

    var row = sheet.getRange(rowNum, 1, 1, 22).getValues()[0];

    // Reconstruct wines array from stored ranks and notes
    var wines = WINE_IDS.map(function(id, i) {
      var flight = FLIGHT_0.indexOf(id) >= 0 ? 0 : 1;
      var rank   = parseFloat(row[2 + i]);
      return {
        id:     id,
        flight: flight,
        rank:   (!isNaN(rank) && rank > 0) ? rank : (i < 5 ? i + 1 : i - 4),
        notes:  row[12 + i] ? row[12 + i].toString() : ''
      };
    });

    return jsonResponse({ ok: true, found: true, wines: wines, finalists: [] });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

/**
 * Returns top 2 from Flight I and top 2 from Flight II by average rank
 * across all tasters who have submitted.
 */
function doGetAggregate(e) {
  try {
    var sheet = getSheet();
    var data  = sheet.getDataRange().getValues();

    var tasterRows = data.slice(1).filter(function(row) {
      return row[0] && row[0].toString().trim() !== '';
    });

    var tasters = tasterRows.map(function(row) { return row[0]; });

    var totals = {}, counts = {};
    WINE_IDS.forEach(function(id) { totals[id] = 0; counts[id] = 0; });

    tasterRows.forEach(function(row) {
      WINE_IDS.forEach(function(id, i) {
        var rank = parseFloat(row[2 + i]);
        if (!isNaN(rank) && rank > 0) {
          totals[id] += rank;
          counts[id]++;
        }
      });
    });

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

    var flight0   = flightAggregate(FLIGHT_0, 0);
    var flight1   = flightAggregate(FLIGHT_1, 1);
    var finalists = flight0.slice(0, 2).concat(flight1.slice(0, 2));

    return jsonResponse({ ok: true, tasters: tasters, finalists: finalists });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}
