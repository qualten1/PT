// ============================================================
// PRODTRACK Apps Script
// ============================================================
// HOW TO UPDATE SAFELY (future updates):
//   1. Paste new code here — replace everything
//   2. Save → Deploy → New Deployment → Web App
//      → Execute as Me → Anyone → Deploy
//   3. Copy new URL → update const API in your HTML files
//   4. NEVER run setupSheets() again after first setup
//   5. If you add new columns to JOB_COLS: run addMissingColumns()
// ============================================================

const JOB_COLS  = ['PO No','PO Date','Customer','Item Desc','Qty','Price','NBD','Notes','Status','StatusUpdatedAt','Mat Notes','Est Times','Lathe','Machine','Urgent'];
const TIME_COLS = ['OPERATOR','DRG NUMBER','OPERATION','START TIME','END TIME','QUANTITY','DATE','REMARKS','TOTAL TIME','TIME PER PIECE'];

function makeResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getAll';
    if (action === 'getAll')  return makeResponse({ ok:true, jobs:readSheet('Jobs',JOB_COLS), timeLogs:readSheet('Timings',TIME_COLS) });
    if (action === 'getJobs') return makeResponse({ ok:true, jobs:readSheet('Jobs',JOB_COLS) });
    if (action === 'getLogs') return makeResponse({ ok:true, timeLogs:readSheet('Timings',TIME_COLS) });
    return makeResponse({ ok:false, error:'Unknown action' });
  } catch(err) { return makeResponse({ ok:false, error:err.toString() }); }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    switch(data.action) {
      case 'addJob':       appendRow('Jobs',    data, JOB_COLS);                                        return makeResponse({ok:true});
      case 'updateJob':    var res=updateJob(data); return makeResponse({ok: res==='ok', result: res});
      case 'addLog':       appendRow('Timings', data, TIME_COLS);                                       return makeResponse({ok:true});
      case 'updateStatus': var sres=updateStatus(data.item, data.status); return makeResponse({ok: sres==='ok', result: sres});
      case 'deleteJob':    deleteRow('Jobs',    'Item Desc', data.item);                                return makeResponse({ok:true});
      case 'deleteLog':    deleteRow('Timings', 'OPERATOR',  data.operator, data.drgNumber, data.date); return makeResponse({ok:true});
      default:             return makeResponse({ok:false, error:'Unknown action: '+data.action});
    }
  } catch(err) { return makeResponse({ ok:false, error:err.toString() }); }
}

function readSheet(name, expectedCols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
  var headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  var rows    = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  var result  = [];
  var tz      = Session.getScriptTimeZone();
  for (var i=0; i<rows.length; i++) {
    var row = rows[i];
    if (row.every(function(c){ return c===''||c===null||c===undefined; })) continue;
    var obj = {};
    for (var j=0; j<headers.length; j++) {
      var h=headers[j], val=row[j];
      if (!h) continue;
      if (val instanceof Date) {
        var isTimeCol=(h==='START TIME'||h==='END TIME'||h==='TOTAL TIME'||h==='TIME PER PIECE');
        val = isTimeCol ? Utilities.formatDate(val,tz,'HH:mm') : Utilities.formatDate(val,tz,'dd/MM/yyyy');
      }
      obj[h] = (val===null||val===undefined) ? '' : String(val);
    }
    // Migrate old lathe IDs in Est Times to new merged IDs
    if (obj['Est Times']) obj['Est Times'] = migrateEstTimes(obj['Est Times']);
    result.push(obj);
  }
  return result;
}

// Remap old individual lathe IDs to merged group IDs
// Safe to run repeatedly — already-migrated IDs pass through unchanged
function migrateEstTimes(raw) {
  if (!raw || raw === '') return raw;
  var REMAP = {
    'lathe_big1': 'lathe_big', 'lathe_big2': 'lathe_big',
    'lathe_med3': 'lathe_med', 'lathe_med4': 'lathe_med', 'lathe_med5': 'lathe_med',
    'lathe_sm6':  'lathe_sm',  'lathe_sm7':  'lathe_sm'
  };
  try {
    var parsed = JSON.parse(raw);
    var migrated = {}, changed = false;
    Object.keys(parsed).forEach(function(k) {
      var newKey = REMAP[k] || k;
      if (newKey !== k) changed = true;
      // If new key already exists (both lathe_big1 + lathe_big2 set), keep whichever came first
      if (!migrated[newKey]) migrated[newKey] = parsed[k];
    });
    return changed ? JSON.stringify(migrated) : raw;
  } catch(e) { return raw; }
}

function appendRow(sheetName, data, cols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    // Sheet missing — create headers only, no data touched
    sh = ss.insertSheet(sheetName);
    sh.getRange(1,1,1,cols.length).setValues([cols]);
    styleHeader(sh, 1, cols.length);
  }
  ensureColumns(sh, cols);
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var row = headers.map(function(h){ var v=data[h]; return (v===undefined||v===null)?'':v; });
  sh.appendRow(row);
}

// Update every field of a job row — matched by original Item Desc + PO No
function updateJob(data) {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName('Jobs');
  if (!sh||sh.getLastRow()<2) { Logger.log('updateJob: Jobs sheet missing or empty'); return 'no_sheet'; }
  
  // Ensure columns exist FIRST, then re-read headers fresh
  ensureColumns(sh, JOB_COLS);
  var lc=sh.getLastColumn(), lr=sh.getLastRow();
  var hdrs=sh.getRange(1,1,1,lc).getValues()[0]; // read AFTER ensureColumns
  
  var ic=hdrs.indexOf('Item Desc'), pc=hdrs.indexOf('PO No');
  if (ic<0) { Logger.log('updateJob: Item Desc column not found'); return 'no_col'; }
  
  var rows=sh.getRange(2,1,lr-1,lc).getValues();
  var origItem = String(data._origItem||data['Item Desc']||'').trim();
  var origPO   = String(data._origPO||'').trim();
  
  Logger.log('updateJob: looking for item="'+origItem+'" po="'+origPO+'" in '+rows.length+' rows');
  
  for (var i=0;i<rows.length;i++) {
    var rowItem = String(rows[i][ic]||'').trim();
    var rowPO   = pc>=0 ? String(rows[i][pc]||'').trim() : '';
    var itemMatch = rowItem.toLowerCase()===origItem.toLowerCase();
    var poMatch   = !origPO || !rowPO || rowPO===origPO;
    if (itemMatch && poMatch) {
      Logger.log('updateJob: found match at row '+(i+2)+' rowItem="'+rowItem+'"');
      JOB_COLS.forEach(function(col){
        var ci=hdrs.indexOf(col);
        if (ci>=0 && data[col]!==undefined && data[col]!==null && col!=='_origItem' && col!=='_origPO') {
          sh.getRange(i+2,ci+1).setValue(data[col]);
        }
      });
      var tc=hdrs.indexOf('StatusUpdatedAt');
      if(tc>=0 && data['Status']) sh.getRange(i+2,tc+1).setValue(new Date().toISOString());
      return 'ok';
    }
  }
  Logger.log('updateJob: NO MATCH for item="'+origItem+'" po="'+origPO+'" — first 3 items: '+rows.slice(0,3).map(function(r){return String(r[ic]);}).join(' | '));
  return 'not_found';
}

function updateStatus(item, status) {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName('Jobs');
  if (!sh||sh.getLastRow()<2) return 'no_sheet';
  ensureColumns(sh, JOB_COLS);
  var lc=sh.getLastColumn(), lr=sh.getLastRow();
  var hdrs=sh.getRange(1,1,1,lc).getValues()[0];
  var ic=hdrs.indexOf('Item Desc'), sc=hdrs.indexOf('Status'), tc=hdrs.indexOf('StatusUpdatedAt');
  if (ic<0||sc<0) return 'no_col';
  var vals=sh.getRange(2,ic+1,lr-1,1).getValues();
  var now = new Date().toISOString();
  for (var i=0;i<vals.length;i++) {
    if (String(vals[i][0]).trim().toLowerCase()===String(item).trim().toLowerCase()) {
      sh.getRange(i+2,sc+1).setValue(status);
      if(tc>=0) sh.getRange(i+2,tc+1).setValue(now);
      return 'ok';
    }
  }
  Logger.log('updateStatus: no row found for item="'+item+'"');
  return 'not_found';
}

function deleteRow(sheetName, matchCol, matchVal, matchVal2, matchVal3) {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(sheetName);
  if (!sh||sh.getLastRow()<2) return;
  var lc=sh.getLastColumn(), lr=sh.getLastRow();
  var hdrs=sh.getRange(1,1,1,lc).getValues()[0];
  var mc=hdrs.indexOf(matchCol);
  if (mc<0) return;
  var rows=sh.getRange(2,1,lr-1,lc).getValues();
  for (var i=rows.length-1;i>=0;i--) {
    var match=String(rows[i][mc]).trim()===String(matchVal).trim();
    if (match&&matchVal2) {
      var drgC=hdrs.indexOf('DRG NUMBER'), datC=hdrs.indexOf('DATE');
      if (drgC>=0) match=match&&String(rows[i][drgC]).trim()===String(matchVal2).trim();
      if (datC>=0&&matchVal3) match=match&&String(rows[i][datC]).trim()===String(matchVal3).trim();
    }
    if (match) { sh.deleteRow(i+2); return; }
  }
}

// ════════════════════════════════════════════════════════════
// SAFE — adds missing columns, NEVER touches data rows
// ════════════════════════════════════════════════════════════
function styleHeader(sh, fromCol, toCol) {
  sh.getRange(1, fromCol, 1, (toCol||fromCol)-fromCol+1)
    .setBackground('#1a2236').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sh.setFrozenRows(1);
}

function ensureColumns(sh, expectedCols) {
  var cur=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var changed=false;
  expectedCols.forEach(function(col){
    if (cur.indexOf(col)===-1) {
      var ni=sh.getLastColumn()+1;
      sh.getRange(1,ni).setValue(col);
      styleHeader(sh,ni,ni);
      cur.push(col);
      changed=true;
    }
  });
  return changed;
}

// ── RUN THIS INSTEAD OF setupSheets() when adding new columns ──
// Zero risk — only touches header row, data is completely safe.
function addMissingColumns() {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), log=[];
  var js=ss.getSheetByName('Jobs');
  if (js) log.push('Jobs: '+(ensureColumns(js,JOB_COLS)?'new columns added ✅':'already up to date ✅'));
  else log.push('Jobs sheet not found ⚠');
  var ts=ss.getSheetByName('Timings');
  if (ts) log.push('Timings: '+(ensureColumns(ts,TIME_COLS)?'new columns added ✅':'already up to date ✅'));
  else log.push('Timings sheet not found ⚠');
  try { SpreadsheetApp.getUi().alert('Column Check\n\n'+log.join('\n')+'\n\nAll data rows untouched.'); }
  catch(e) { Logger.log(log.join('\n')); }
}

// ════════════════════════════════════════════════════════════
// FIRST-TIME SETUP ONLY — run once on a brand new empty sheet
// After first setup, NEVER run this again.
// ════════════════════════════════════════════════════════════
function setupSheets() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var js=ss.getSheetByName('Jobs'), ts=ss.getSheetByName('Timings');
  var hasData=(js&&js.getLastRow()>1)||(ts&&ts.getLastRow()>1);

  if (hasData) {
    try {
      var ui=SpreadsheetApp.getUi();
      var resp=ui.alert(
        '⚠ YOUR DATA EXISTS — STOP!',
        'Your sheets already have data.\n\nDo NOT run setupSheets() — it will delete everything.\n\nInstead:\n• To add new columns → run addMissingColumns()\n• To update the script → just paste and redeploy, no function needed\n\nClick CANCEL to exit safely.',
        ui.ButtonSet.OK_CANCEL
      );
      if (resp!==ui.Button.OK) return;
      // Second warning
      var resp2=ui.alert('ARE YOU SURE?','This will DELETE ALL EXISTING DATA. Are you absolutely sure this is a brand new empty sheet?',ui.ButtonSet.OK_CANCEL);
      if (resp2!==ui.Button.OK) return;
    } catch(e) { Logger.log('setupSheets() aborted safely — data exists.'); return; }
  }

  function makeSheet(name, cols, widths) {
    var sh=ss.getSheetByName(name);
    if (!sh) sh=ss.insertSheet(name);
    if (sh.getLastRow()<=1) { sh.clearContents(); sh.clearFormats(); try{sh.getBandings().forEach(function(b){b.remove();});}catch(e){} }
    sh.getRange(1,1,1,cols.length).setValues([cols]);
    styleHeader(sh,1,cols.length);
    widths.forEach(function(w,i){ try{sh.setColumnWidth(i+1,w);}catch(e){} });
    try{ sh.getRange(2,1,Math.min(500,sh.getMaxRows()-1),cols.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY); }catch(e){}
    return sh;
  }

  makeSheet('Jobs',    JOB_COLS,  [80,90,120,180,50,80,90,180,120,180,200,110,110]);
  makeSheet('Timings', TIME_COLS, [100,180,130,80,80,70,90,180,80,100]);
  ['Operations','Operators','TimeLogs'].forEach(function(n){
    var s=ss.getSheetByName(n);
    if(s&&s.getLastRow()<=1){try{ss.deleteSheet(s);}catch(e){}}
  });

  try { SpreadsheetApp.getUi().alert('✅ Setup done!\n\nNow deploy as Web App and copy the URL.\n\n⚠ Do NOT run setupSheets() again — use addMissingColumns() for future changes.'); }
  catch(e) {}
}

// ════════════════════════════════════════════════════════════
// RUN ONCE — fixes old lathe IDs directly in the sheet
// Safe: only rewrites Est Times cells that contain old IDs
// After running, old IDs are gone from the sheet permanently
// ════════════════════════════════════════════════════════════
function migrateLatheIdsInSheet() {
  var REMAP = {
    'lathe_big1':'lathe_big', 'lathe_big2':'lathe_big',
    'lathe_med3':'lathe_med', 'lathe_med4':'lathe_med', 'lathe_med5':'lathe_med',
    'lathe_sm6' :'lathe_sm',  'lathe_sm7' :'lathe_sm'
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Jobs');
  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No Jobs sheet or no data found.');
    return;
  }

  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var estCol  = headers.indexOf('Est Times');
  if (estCol < 0) {
    SpreadsheetApp.getUi().alert('Est Times column not found.');
    return;
  }

  var rows    = sh.getRange(2, estCol+1, sh.getLastRow()-1, 1).getValues();
  var updated = 0;
  var log     = [];

  for (var i = 0; i < rows.length; i++) {
    var raw = String(rows[i][0] || '').trim();
    if (!raw) continue;
    try {
      var parsed   = JSON.parse(raw);
      var migrated = {};
      var changed  = false;
      Object.keys(parsed).forEach(function(k) {
        var newKey = REMAP[k] || k;
        if (newKey !== k) changed = true;
        // If new key already exists keep whichever has a value
        if (!migrated[newKey]) {
          migrated[newKey] = parsed[k];
        }
      });
      if (changed) {
        var newVal = JSON.stringify(migrated);
        sh.getRange(i+2, estCol+1).setValue(newVal);
        updated++;
        log.push('Row '+(i+2)+': '+raw+' → '+newVal);
      }
    } catch(e) {
      log.push('Row '+(i+2)+': SKIPPED (bad JSON): '+raw);
    }
  }

  var msg = updated + ' row(s) updated.\n\n' + (log.join('\n') || 'Nothing to change — all IDs already up to date.');
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('Migration Complete\n\n' + msg);
}
