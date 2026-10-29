/**
 * Sample Box Writer  -  can change ONE thing: the "Sample Location" column of the QA checklist sheet.
 * Bound to the sheet (currently the TEST copy). It refuses to do anything else:
 *   - only the column whose header is "Sample Location" is ever written (found by header name)
 *   - each change must name the row, that row's ID (column A), and the value the cell is expected to hold now;
 *     if any of those don't match the sheet, that change is skipped (so it can never overwrite a newer edit)
 *   - values can't start with = + - @ (no formulas), max 60 characters, no line breaks
 *   - a passphrase is required; at most 1500 changes per call; calls are serialized with a lock
 * POST (text/plain JSON): {pass, batch, changes:[{row, id, expect, set}, ...]}
 * Response: {ok, batch, written, results:["ok"|"already"|"id-mismatch"|"changed-by-someone"|"bad-row"|"bad-value", ...]}
 */
var TAB = 'ChecklistTemplate_Sample';
var TARGET = 'Sample Location';
var MAX_CHANGES = 1500;
var PASS_HASH = 'c1c92f1c857767db838157e4c08d5bf01eb0a8c0319327470ebdeb164006cda2';

function doGet(e) { return out({ ok: true, role: 'writer' }); }

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad json' }); }
  if (!passOk(body.pass)) { Utilities.sleep(2000); return out({ ok: false, error: 'bad passphrase' }); }
  var changes = body.changes;
  if (!Array.isArray(changes) || !changes.length) return out({ ok: false, error: 'no changes' });
  if (changes.length > MAX_CHANGES) return out({ ok: false, error: 'too many changes (max ' + MAX_CHANGES + ')' });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return out({ ok: false, error: 'busy, try again' });
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(TAB);
    if (!sh) return out({ ok: false, error: 'tab not found' });
    var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    var col = hdr.indexOf(norm(TARGET)) + 1;
    if (col < 1) return out({ ok: false, error: 'column not found: ' + TARGET });
    var last = sh.getLastRow();
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    var cur = sh.getRange(2, col, last - 1, 1).getValues();
    var letter = colLetter(col), results = [], groups = {}, written = 0;

    changes.forEach(function (c) {
      var row = Number(c.row);
      if (!(row >= 2 && row <= last) || row !== Math.floor(row)) { results.push('bad-row'); return; }
      var set = String(c.set == null ? '' : c.set).trim();
      if (set.length > 60 || /[\r\n]/.test(set) || /^[=+\-@]/.test(set)) { results.push('bad-value'); return; }
      var i = row - 2;
      if (String(ids[i][0]) !== String(c.id)) { results.push('id-mismatch'); return; }
      var now = String(cur[i][0] == null ? '' : cur[i][0]).trim();
      if (now === set) { results.push('already'); return; }
      if (now !== String(c.expect == null ? '' : c.expect).trim()) { results.push('changed-by-someone'); return; }
      (groups[set] = groups[set] || []).push(letter + row);
      results.push('ok'); written++;
    });

    Object.keys(groups).forEach(function (value) {
      var list = sh.getRangeList(groups[value]);
      if (value === '') list.clearContent(); else list.setValue(value);
    });
    SpreadsheetApp.flush();
    return out({ ok: true, batch: body.batch || '', written: written, results: results });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

function passOk(p) {
  if (!p) return false;
  var h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'sbw1:' + String(p), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  var diff = h.length ^ PASS_HASH.length;
  for (var i = 0; i < h.length && i < PASS_HASH.length; i++) diff |= h.charCodeAt(i) ^ PASS_HASH.charCodeAt(i);
  return diff === 0;
}
function norm(s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function colLetter(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
// Run this once from the editor to grant the permission (it only reads the sheet's name).
function authorize() { return SpreadsheetApp.getActive().getName(); }
