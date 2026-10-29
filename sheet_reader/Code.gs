/**
 * Sample Box Reader  -  READ-ONLY (enforced by Google: the only permission this script has is "view spreadsheets").
 * Reads the QA checklist Google Sheet (currently the TEST copy) through the Sheets API and returns the rows
 * that have a Sample Location, found by HEADER NAME (not column letter), in the shape the Sample Box Tracker uses.
 *
 * GET ?action=ping   -> {ok:true}
 * GET ?action=rows   -> {ok, total, withLocation, hash, rows:[[id, tsWallMs, job, item, loc, sheetRow], ...]}
 */
var TAB = 'ChecklistTemplate_Sample';
var SHEET_ID = '1FUumm9OTYdUXq3GDBf1NAO68HaORlTBl9nYGwkkxQgU';   // TEST copy

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'ping') return out({ ok: true, tab: TAB });
    if (p.action === 'rows') return out(readRows());
    return out({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function norm(s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function colLetter(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function readRows() {
  var q = "'" + TAB + "'!";
  var head = Sheets.Spreadsheets.Values.get(SHEET_ID, q + '1:1').values;
  if (!head || !head[0]) throw new Error('no header row');
  var hdr = head[0].map(norm);
  function col(name) { var i = hdr.indexOf(norm(name)); if (i < 0) throw new Error('missing column: ' + name); return i + 1; }
  var cTs = col('Timestamp'), cJob = col('Job_Number'), cItem = col('Inventory_Item'), cLoc = col('Sample Location');

  function rng(c) { var L = colLetter(c); return q + L + '2:' + L; }
  var res = Sheets.Spreadsheets.Values.batchGet(SHEET_ID, {
    ranges: [rng(1), rng(cTs), rng(cJob), rng(cItem), rng(cLoc)],
    valueRenderOption: 'UNFORMATTED_VALUE',            // numbers stay numbers, text stays text
    dateTimeRenderOption: 'SERIAL_NUMBER',             // timestamps as the sheet's own wall-clock serial
    majorDimension: 'COLUMNS'
  }).valueRanges;
  var ids = res[0].values ? res[0].values[0] : [], tss = res[1].values ? res[1].values[0] : [],
      jobs = res[2].values ? res[2].values[0] : [], itms = res[3].values ? res[3].values[0] : [],
      locs = res[4].values ? res[4].values[0] : [];

  var rows = [];
  for (var i = 0; i < locs.length; i++) {
    var loc = String(locs[i] == null ? '' : locs[i]).trim();
    if (!loc) continue;                                  // no sample kept for this job
    rows.push([
      String(ids[i] == null ? '' : ids[i]),
      wallMs(tss[i]),
      cleanJob(jobs[i]),
      cleanItem(itms[i]),
      loc,
      i + 2                                              // sheet row number (used later for write-back)
    ]);
  }
  var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(rows)));
  return { ok: true, total: Math.max(tss.length, locs.length), withLocation: rows.length, hash: hash, rows: rows };
}

// The app stores each row's time as a real instant: the sheet's Pacific wall-clock time converted with the Pacific offset
// (this script's time zone is America/Los_Angeles, so DST is handled the same way the browser does it).
function wallMs(v) {
  if (typeof v !== 'number') return null;
  var d = new Date(Math.round((v - 25569) * 86400000));        // wall clock as UTC components
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds()).getTime();
}
// The app's keys use the job number WITHOUT leading zeros ("041510" -> "41510")
function cleanJob(v) { var s = String(v == null ? '' : v).trim().replace(/\.0$/, ''); return s.replace(/^0+(?=\d)/, ''); }
function cleanItem(v) { var s = String(v == null ? '' : v).trim().replace(/\.0$/, ''); return s === '0' ? '' : s; }
function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
