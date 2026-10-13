/**
 * Sample Box Tracker - shared backend (Google Apps Script bound to a Google Sheet).
 *
 * Sheets (created on first use):
 *   Rows    k | box | series | ts | job | item        imported sample history (replaced on every import)
 *   Pulls   k | from | to | at                         samples pulled out of a box (+ where they went)
 *   Manual  k | box | series | ts | job | item         samples added by hand
 *   Meta    box | loc | cap | full                     per-box location / capacity / full flag
 *   Config  key | value                                default capacity, import time, ignored-location summary
 *   Log     at | op | key                              tiny audit trail
 *
 * Quota design: every write bumps a version number in Script Properties. A client polling with the
 * current version gets back {v} only - no sheet is opened - so idle polling costs ~nothing.
 * All writes take a script lock and are idempotent upserts by key.
 */
var HEADERS = {
  Rows:   ['k','box','series','ts','job','item'],
  Pulls:  ['k','from','to','at'],
  Manual: ['k','box','series','ts','job','item'],
  Meta:   ['box','loc','cap','full'],
  Config: ['key','value'],
  Log:    ['at','op','key']
};

function doGet(e){
  var p = (e && e.parameter) || {};
  if (p.action === 'ping') return out({ok:true, v: version()});
  var v = version();
  if (p.since !== undefined && String(p.since) === String(v)) return out({ok:true, v:v, unchanged:true});
  return out({ok:true, v:v, state: readState()});
}

function doPost(e){
  var body;
  try { body = JSON.parse(e.postData.contents); } catch(err){ return out({ok:false, error:'bad json'}); }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return out({ok:false, error:'busy'});
  try {
    var ops = body.ops || [];
    for (var i=0;i<ops.length;i++) applyOp(ops[i]);
    var v = bump();
    return out({ok:true, v:v});
  } catch(err){
    return out({ok:false, error:String(err)});
  } finally { lock.releaseLock(); }
}

function applyOp(o){
  switch(o.op){
    case 'setPull':   upsert('Pulls', o.pull.k, [o.pull.k, o.pull.from||'', o.pull.to||'', o.pull.at||'']); log(o.op,o.pull.k); break;
    case 'clearPull': remove('Pulls', o.k); log(o.op,o.k); break;
    case 'addManual': var m=o.entry; upsert('Manual', m.k, [m.k,m.box,m.series||'Box',m.ts||'',m.job||'',m.item||'']); log(o.op,m.k); break;
    case 'delManual': remove('Manual', o.k); remove('Pulls', o.k); log(o.op,o.k); break;
    case 'setMeta':   setMeta(o.box, o.meta); log(o.op,o.box); break;
    case 'setCap':    setConfig('cap', o.cap); log(o.op,String(o.cap)); break;
    case 'import':    doImport(o); log(o.op, String((o.rows||[]).length)); break;
    default: throw new Error('unknown op '+o.op);
  }
}

/* ---------- import: replace Rows (+ pruned Pulls/Manual) in one go ---------- */
function doImport(o){
  var rows = (o.rows||[]).map(function(r){ return [r.k,r.box,r.series||'Box',r.ts||'',r.job||'',r.item||'']; });
  if (rows.length === 0) throw new Error('refusing to import 0 rows');          // data-safety guard
  replaceAll('Rows', rows);
  replaceAll('Pulls',  (o.pulled||[]).map(function(p){ return [p.k,p.from||'',p.to||'',p.at||'']; }));
  replaceAll('Manual', (o.manual||[]).map(function(m){ return [m.k,m.box,m.series||'Box',m.ts||'',m.job||'',m.item||'']; }));
  setConfig('stamp', o.stamp || Date.now());
  setConfig('ignored', JSON.stringify(o.ignored||{}).slice(0,40000));
}

/* ---------- read ---------- */
function readState(){
  var cfg = {};
  table('Config').forEach(function(r){ cfg[r[0]] = r[1]; });
  var meta = {};
  table('Meta').forEach(function(r){ var m={}; if(r[1]) m.loc=r[1]; if(r[2]!=='') m.cap=Number(r[2]); if(r[3]===true||r[3]==='TRUE'||r[3]==='true') m.full=true; meta[r[0]]=m; });
  var pulled = {};
  table('Pulls').forEach(function(r){ pulled[r[0]] = {from:String(r[1]), to:String(r[2]), at:Number(r[3])||0}; });
  var ign = {}; try { ign = JSON.parse(cfg.ignored||'{}'); } catch(e){}
  return {
    cap: Number(cfg.cap)||100,
    stamp: Number(cfg.stamp)||null,
    ignored: ign,
    boxMeta: meta,
    pulled: pulled,
    rows:   table('Rows').map(rowObj),
    manual: table('Manual').map(function(r){ var o=rowObj(r); o.manual=true; return o; })
  };
}
function rowObj(r){ return {k:String(r[0]), box:String(r[1]), series:String(r[2]||'Box'), ts: r[3]===''?null:Number(r[3]), job:String(r[4]), item:String(r[5])}; }

/* ---------- sheet helpers ---------- */
function sheet(name){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh){
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2,1,sh.getMaxRows()-1,HEADERS[name].length).setNumberFormat('@');   // keep job/item/keys as text
  }
  return sh;
}
function table(name){
  var sh = sheet(name), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2,1,n-1,HEADERS[name].length).getValues();
}
function findRow(sh, key){
  var n = sh.getLastRow(); if (n < 2) return -1;
  var ks = sh.getRange(2,1,n-1,1).getValues();
  for (var i=0;i<ks.length;i++) if (String(ks[i][0]) === String(key)) return i+2;
  return -1;
}
function upsert(name, key, vals){
  var sh = sheet(name), r = findRow(sh, key);
  if (r < 0) r = Math.max(sh.getLastRow(),1) + 1;
  sh.getRange(r,1,1,vals.length).setNumberFormat('@').setValues([vals.map(function(x){return x===''||x==null?'':String(x);})]);
}
function remove(name, key){
  var sh = sheet(name), r = findRow(sh, key);
  if (r > 0) sh.deleteRow(r);
}
function replaceAll(name, rows){
  var sh = sheet(name), n = sh.getLastRow(), w = HEADERS[name].length;
  if (n > 1) sh.getRange(2,1,n-1,w).clearContent();
  if (rows.length){
    if (sh.getMaxRows() < rows.length+1) sh.insertRowsAfter(sh.getMaxRows(), rows.length+1-sh.getMaxRows());
    var rng = sh.getRange(2,1,rows.length,w);
    rng.setNumberFormat('@');
    rng.setValues(rows.map(function(r){ return r.map(function(x){return x===''||x==null?'':String(x);}); }));
  }
}
function setMeta(box, m){
  var sh = sheet('Meta'), r = findRow(sh, box);
  if (!m || (!m.loc && m.cap==null && !m.full)){ if (r>0) sh.deleteRow(r); return; }
  upsert('Meta', box, [box, m.loc||'', m.cap==null?'':m.cap, m.full?'true':'']);
}
function setConfig(key, value){ upsert('Config', key, [key, value]); }

/* ---------- version / log / output ---------- */
function version(){ return Number(PropertiesService.getScriptProperties().getProperty('v')||'0'); }
function bump(){ var v=version()+1; PropertiesService.getScriptProperties().setProperty('v', String(v)); return v; }
function log(op,key){ sheet('Log').appendRow([new Date().toISOString(), op, key]); }
function out(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
