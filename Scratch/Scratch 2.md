# Main Hub

```js
// STANT PM — Monthly Imports + Monthly Debits (Dynamic Month Tabs, Row 4 Inserts)
// Time zone: set your Apps Script project to America/New_York before creating triggers.

/* =========================
   CONFIG
   ========================= */
var CONFIG = {
  MASTER_NAME: 'September 25 Income Sheets',

  // Current TX sheet URL (update per month if you make a new TX file)
  TRANSACTION_SHEET_URL: 'https://docs.google.com/spreadsheets/....',

  // TX columns (1-based)
  TX_COLUMNS: { PROPERTY: 1, UNIT: 2, DATE: 7, EXPLANATION: 8, DEBIT: 5, MARKUP_INCLUDED: 9 },

  // Property source files (we will import only the tab that matches the current month name, e.g., Sep25)
  PROPERTY_SOURCES: [
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....'
  ],

  // Sheet structure
  HEADER_ROW: 1,
  DATA_START_ROW: 4,      // data begins at row 4
  WRITABLE_START_ROW: 4,  // always insert at row 4

  // Fallback header locations if detection fails
  AMOUNT_COL_INDEX: 4,
  MARKUP_FLAG_COL_INDEX: 8,
  MARKUP_REV_COL_INDEX: 9,

  // Canonical mapping for property names
  PROPERTY_ALIASES: {
    '1505-1515 Franklin Park': ['1505-1515 Franklin Park','Franklin Park'],
    'Ohio & Bryden':           ['Ohio & Bryden','Ohio and Bryden','224 S Ohio Ave','1096-1104 Bryden Rd','Ohio Bryden'],
    '196 Miller':              ['196 Miller','196 Miller Ave'],
    '705 Ann':                 ['705 Ann'],
    'Schiller Terrace':        ['Schiller Terrace'],
    '189 Patterson':           ['189 Patterson','189 W Patterson Ave'],
    '2536 Adams':              ['2536 Adams','2536 Adams Ave','Adams'],
    '22 Wilson':               ['22 Wilson','22 Wilson Ave'],
    'Park and State':          ['Park & State','Park and State','Park State'],
    '1476 S High St':          ['1476 S High St','1476 S High']
  },

  // Cover
  COVER_SHEET_NAME: 'Cover',
  LABEL_DUE_TO_OWNERS: /due\s*to\s*owners/i,
  LABEL_TOTAL_TO_STANT: /total\s*to\s*stant/i,
  CURRENCY_FORMAT: '$#,##0.00;-$#,##0.00',

  // Template tab used only if we ever need to create a new property tab
  TEMPLATE_TAB_NAME: 'Data Entry Pop',

  // Preferred ordering
  PREFERRED_ORDER: [
    { label: 'Cover', patterns: [/^cover$/i] },
    { label: '1505-1515 Franklin Park', patterns: [/franklin\s*park/i] },
    { label: 'Ohio & Bryden', patterns: [/224\b.*ohio/i, /bryden/i] },
    { label: '196 Miller', patterns: [/(^|\s)196\b.*miller/i] },
    { label: '705 Ann', patterns: [/(^|\s)705\b.*ann/i] },
    { label: 'Schiller Terrace', patterns: [/schiller\s*terrace/i] },
    { label: '189 Patterson', patterns: [/(^|\s)189\b.*patterson/i] },
    { label: '2536 Adams', patterns: [/(^|\s)2536\b.*adams/i] },
    { label: '22 Wilson', patterns: [/(^|\s)22\b.*wilson/i] },
    { label: 'Park and State', patterns: [/park\s*(and|&)\s*state/i] },
    { label: '1476 S High St', patterns: [/(^|\s)1476\b.*high/i] }
  ],

  FLUSH_INTERVAL: 10
};

/* =========================
   SMALL UTILITIES
   ========================= */
function notify_(msg){ try{ SpreadsheetApp.getActive().toast(msg); }catch(_){} Logger.log(msg); }
function colLetter_(n){ var s=''; while(n>0){ s=String.fromCharCode(65+((n-1)%26))+s; n=Math.floor((n-1)/26);} return s; }
function ensureMinRows_(sh, minIndex){ var mr=sh.getMaxRows(); if(minIndex>mr) sh.insertRowsAfter(mr, minIndex-mr); }
function ensureMinColumns_(sh, minIndex){ var mc=sh.getMaxColumns(); if(minIndex>mc) sh.insertColumnsAfter(mc, minIndex-mc); }
function canon_(s){ var x=String(s||'').toLowerCase(); x=x.replace(/[\u2010-\u2015]/g,'-').replace(/[^\w\- ]+/g,'').replace(/\s*-\s*/g,'-').replace(/\s+/g,' ').trim(); return x; }

/* =========================
   MONTH TAB NAME — Dynamic (e.g., "Sep25")
   ========================= */
function monthTabNameForDate_(d){
  var m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()];
  var yy = String(d.getFullYear()).slice(-2);
  return m + yy;
}
function currentMonthTabName_(){
  // Use New York time so month flips at local midnight
  var tz = 'America/New_York';
  var now = new Date(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd\'T\'HH:mm:ss'));
  return monthTabNameForDate_(now);
}

/* =========================
   CANONICAL NAMES / ALIASES
   ========================= */
function canonProp_(s){
  var x = String(s || '').toLowerCase();
  x = x.replace(/&/g,' and ');
  x = x.replace(/[^\w ]+/g,' ');
  x = x.replace(/\b(and)\b/g,' ');
  x = x.replace(/\s+/g,' ').trim();
  return x;
}
function aliasesForTab_(tabName, a1Text){
  var list=[], seen={};
  function add(v){ if(!v) return; if(!seen[v]){ list.push(v); seen[v]=true; } }
  var map = CONFIG.PROPERTY_ALIASES || {};
  var fromKey = map[tabName] || map[a1Text] || [];
  for (var i=0;i<fromKey.length;i++) add(fromKey[i]);
  add(tabName); add(a1Text);
  return list;
}
function preferredTabName_(nameOrA1){
  var raw = String(nameOrA1 || '').trim(); if(!raw) return raw;
  var keys = Object.keys(CONFIG.PROPERTY_ALIASES || {});
  var rc = canonProp_(raw);
  for (var i=0;i<keys.length;i++){
    var key=keys[i]; if (canonProp_(key)===rc) return key;
    var al = CONFIG.PROPERTY_ALIASES[key]||[];
    for (var j=0;j<al.length;j++){ if (canonProp_(al[j])===rc) return key; }
  }
  return raw;
}

/* =========================
   HEADER / COLUMNS / ANCHOR
   ========================= */
function findHeaderRow_(sh){
  var maxScan = Math.min(15, sh.getLastRow());
  for (var r=1;r<=maxScan;r++){
    var row = sh.getRange(r,1,1,sh.getLastColumn()).getDisplayValues()[0]
               .map(function(x){return String(x||'').trim().toLowerCase();}).join(' ');
    if (/\bunit\b/.test(row) && /\bdebits?\b/.test(row)) return r;
  }
  return 1;
}
function detectTargetColumnsAtRow_(sh, headerRow){
  var raw = sh.getRange(headerRow,1,1,sh.getLastColumn()).getValues()[0];
  var H = raw.map(function(h){ return String(h||'').replace(/\s+/g,' ').trim(); });
  function find(re){ for (var i=0;i<H.length;i++) if (re.test(H[i])) return i+1; return null; }
  var unitCol   = find(/^unit$/i) || find(/unit\s*#?|unit\s*number/i) || 1;
  var debitsCol = find(/^debits?$/i) || find(/amount|price|total/i) || CONFIG.AMOUNT_COL_INDEX;
  var dateCol   = find(/^date$/i) || find(/trans.*date|transaction.*date|date\s*paid/i);
  if (!dateCol){ var H2=H.map(function(x){return x.replace(/[^\w ]/g,'');});
    for (var i=0;i<H2.length;i++) if (/\bdate\b/i.test(H2[i])) { dateCol=i+1; break; } }
  var explCol   = find(/^debit\/?credit explanation$/i) || find(/debit.*credit.*explan|explan|descr|description|memo/i);
  if (!explCol){ var H3=H.map(function(x){return x.replace(/[^\w ]/g,'');});
    for (var j=0;j<H3.length;j++) if (/(debit.*credit.*explan|explan|descr|description|memo)/i.test(H3[j])) { explCol=j+1; break; } }
  var markupCol = find(/^markup included$/i) || find(/markup.*included|markup/i) || CONFIG.MARKUP_FLAG_COL_INDEX;
  var mrevCol   = find(/^markup revenue$/i)  || find(/markup.*revenue|markup.*rev/i) || CONFIG.MARKUP_REV_COL_INDEX;
  ensureMinColumns_(sh, Math.max(unitCol, debitsCol, dateCol||0, explCol||0, markupCol, mrevCol));
  return { unitCol:unitCol, debitsCol:debitsCol, dateCol:dateCol, explCol:explCol, markupCol:markupCol, mrevCol:mrevCol };
}
function findAnchorRow_(sh){
  var vals = sh.getDataRange().getDisplayValues();
  var rMR=null, rMAF=null;
  for (var r=CONFIG.HEADER_ROW; r<vals.length; r++){
    for (var c=0;c<vals[r].length;c++){
      var t = String(vals[r][c]||'').trim();
      if (!rMR && /^markup\s*revenue$/i.test(t)) rMR=r+1;
      if (!rMAF && /^maf$/i.test(t)) rMAF=r+1;
    }
  }
  var candidates=[]; if (rMR) candidates.push(rMR); if (rMAF) candidates.push(rMAF);
  if (candidates.length) return Math.min.apply(null,candidates);
  var tr = getTotalsRow_(sh);
  return tr ? tr : (sh.getLastRow()+1);
}
function getTotalsRow_(sh){
  var v=sh.getDataRange().getValues();
  for (var r=0;r<v.length;r++){
    for (var c=0;c<v[r].length;c++){
      var cell=v[r][c];
      if (cell && cell.toString && /^totals?\s*:?\s*$/i.test(String(cell).trim())) return r+1;
    }
  }
  return null;
}

/* =========================
   TRANSACTIONS (Debits only)
   ========================= */
function getTransactionsDynamic_(){
  var m = CONFIG.TRANSACTION_SHEET_URL.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!m) return {txRows:[]};
  var ss = SpreadsheetApp.openById(m[1]);
  var sh = ss.getSheets()[0];
  var values = sh.getDataRange().getValues();
  if (!values.length) return {txRows:[]};
  var H = CONFIG.TX_COLUMNS||{};
  var rows = values.slice(CONFIG.HEADER_ROW).map(function(r){
    var property = String((r[(H.PROPERTY||1)-1]||'')).trim();
    var unit     = String((r[(H.UNIT||2)-1]||'')).trim();
    var date     = r[(H.DATE||7)-1];
    var expl     = String((r[(H.EXPLANATION||8)-1]||'')).trim();
    var debitAmt = r[(H.DEBIT||5)-1];
    if (typeof debitAmt==='string'){ var cl = debitAmt.replace(/[^0-9.\-]/g,''); debitAmt = cl===''?null:Number(cl); }
    var markup   = !!r[(H.MARKUP_INCLUDED||9)-1];
    return { property:property, unit:unit, date:date, expl:expl, debitAmount:debitAmt, markup:markup };
  }).filter(function(t){ return t.debitAmount!=null && t.debitAmount!=='' && !isNaN(t.debitAmount); });
  return {txRows:rows};
}

/* =========================
   IMPORTS — Current Month Only
   ========================= */
// ========== STANT PM — IMPORTS ONLY (Dynamic Month, Safe Deletes, Cover, Order) ==========
// Project Time Zone: set to America/New_York

/* =========================
   CONFIG
   ========================= */
var CONFIG = {
  MASTER_NAME: 'September 25 Income Sheets',

  // Property source files (we import only the current month tab from each)
  PROPERTY_SOURCES: [
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....',
    'https://docs.google.com/spreadsheets/....'
  ],

  // Canonical mapping for property names (used to normalize imported tab names)
  PROPERTY_ALIASES: {
    '1505-1515 Franklin Park': ['1505-1515 Franklin Park','Franklin Park'],
    'Ohio & Bryden':           ['Ohio & Bryden','Ohio and Bryden','224 S Ohio Ave','1096-1104 Bryden Rd','Ohio Bryden'],
    '196 Miller':              ['196 Miller','196 Miller Ave'],
    '705 Ann':                 ['705 Ann'],
    'Schiller Terrace':        ['Schiller Terrace'],
    '189 Patterson':           ['189 Patterson','189 W Patterson Ave'],
    '2536 Adams':              ['2536 Adams','2536 Adams Ave','Adams'],
    '22 Wilson':               ['22 Wilson','22 Wilson Ave'],
    'Park and State':          ['Park & State','Park and State','Park State'],
    '1476 S High St':          ['1476 S High St','1476 S High']
  },

  // Cover + formatting
  COVER_SHEET_NAME: 'Cover',
  LABEL_DUE_TO_OWNERS: /due\s*to\s*owners/i,
  LABEL_TOTAL_TO_STANT: /total\s*to\s*stant/i,
  CURRENCY_FORMAT: '$#,##0.00;-$#,##0.00',

  // Preferred tab order
  PREFERRED_ORDER: [
    { label: 'Cover',                        patterns: [/^cover$/i] },
    { label: '1505-1515 Franklin Park',      patterns: [/franklin\s*park/i] },
    { label: 'Ohio & Bryden',                patterns: [/224\b.*ohio/i, /bryden/i] },
    { label: '196 Miller',                   patterns: [/(^|\s)196\b.*miller/i] },
    { label: '705 Ann',                      patterns: [/(^|\s)705\b.*ann/i] },
    { label: 'Schiller Terrace',             patterns: [/schiller\s*terrace/i] },
    { label: '189 Patterson',                patterns: [/(^|\s)189\b.*patterson/i] },
    { label: '2536 Adams',                   patterns: [/(^|\s)2536\b.*adams/i] },
    { label: '22 Wilson',                    patterns: [/(^|\s)22\b.*wilson/i] },
    { label: 'Park and State',               patterns: [/park\s*(and|&)\s*state/i] },
    { label: '1476 S High St',               patterns: [/(^|\s)1476\b.*high/i] }
  ]
};

/* =========================
   UTILS
   ========================= */
function notify_(msg){ try{ SpreadsheetApp.getActive().toast(msg); }catch(_){} Logger.log(msg); }
function colLetter_(n){ var s=''; while(n>0){ s=String.fromCharCode(65+((n-1)%26))+s; n=Math.floor((n-1)/26);} return s; }
function canon_(s){ var x=String(s||'').toLowerCase(); x=x.replace(/[\u2010-\u2015]/g,'-').replace(/[^\w\- ]+/g,'').replace(/\s*-\s*/g,'-').replace(/\s+/g,' ').trim(); return x; }
function canonProp_(s){
  var x = String(s || '').toLowerCase();
  x = x.replace(/&/g,' and ').replace(/[^\w ]+/g,' ').replace(/\b(and)\b/g,' ').replace(/\s+/g,' ').trim();
  return x;
}
function preferredTabName_(nameOrA1){
  var raw = String(nameOrA1 || '').trim(); if(!raw) return raw;
  var keys = Object.keys(CONFIG.PROPERTY_ALIASES || {}), rc = canonProp_(raw);
  for (var i=0;i<keys.length;i++){
    var key=keys[i]; if (canonProp_(key)===rc) return key;
    var al = CONFIG.PROPERTY_ALIASES[key]||[];
    for (var j=0;j<al.length;j++) if (canonProp_(al[j])===rc) return key;
  }
  return raw;
}

/* =========================
   DYNAMIC MONTH TAB NAME (e.g., "Sep25")
   ========================= */
function monthTabNameForDate_(d){
  var m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()];
  var yy = String(d.getFullYear()).slice(-2);
  return m + yy;
}
function currentMonthTabName_(){
  var tz = 'America/New_York';
  var now = new Date(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd\'T\'HH:mm:ss'));
  return monthTabNameForDate_(now);
}

/* =========================
   SAFE DELETE HELPERS
   ========================= */
function safeDeleteSheet_(ss, sh) {
  if (!ss || !sh) return;
  try {
    if (ss.getSheets().length <= 1) { sh.clear(); return; } // never delete last sheet
    var active = ss.getActiveSheet();
    if (active && active.getSheetId() === sh.getSheetId()) {
      var alt = ss.getSheets().find(function(s){ return s.getSheetId() !== sh.getSheetId(); });
      if (alt) ss.setActiveSheet(alt);
    }
    var latest = ss.getSheetByName(sh.getName());
    if (latest) ss.deleteSheet(latest);
  } catch (e) { Logger.log('safeDeleteSheet_ skipped: ' + (e && e.message)); }
}
function removeSheetIfExists_(master, name) {
  try {
    var sh = master.getSheetByName(name);
    if (!sh) return;
    if (master.getSheets().length <= 1) {
      if (name !== CONFIG.COVER_SHEET_NAME) { sh.clear(); sh.setName(CONFIG.COVER_SHEET_NAME); }
      return;
    }
    safeDeleteSheet_(master, sh);
  } catch (e) { Logger.log('removeSheetIfExists_("'+name+'") -> ' + (e && e.message)); }
}
function removeLegendSheets_(master) {
  master.getSheets().forEach(function(s){
    var nm = String(s.getName() || '');
    if (/legend/i.test(nm)) safeDeleteSheet_(master, s);
  });
  // Delete ALL sheets whose canonical name matches the given property (keep zero before re-import)
function deleteSheetsByCanon_(ss, propertyDisplayName) {
  var targetCanon = canonProp_(propertyDisplayName);
  var sheets = ss.getSheets();
  var victims = [];

  // Build canonical for each sheet: prefer A1 if present, else sheet name
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var a1 = '';
    try { a1 = String(sh.getRange('A1').getDisplayValue() || '').trim(); } catch (_) {}
    var nameCanon = canonProp_(a1 || sh.getName());
    if (nameCanon === targetCanon) victims.push(sh);
  }

  // Don’t delete the last sheet in the doc; if it’s the only one, just clear & rename
  if (victims.length === 0) return;
  if (sheets.length <= victims.length) {
    // leave one victim to avoid 0-sheet error
    victims.pop();
  }

  victims.forEach(function(sh){ safeDeleteSheet_(ss, sh); });
}

}

/* =========================
   IMPORT CORE
   ========================= */
function withRetry_(fn,label){ var n=0,last; while(n<3){ try{ return fn(); } catch(e){ last=e; Utilities.sleep(500+n*800);} n++; } throw new Error((label||'op')+' failed: '+(last&&last.message)); }
function openSource_(url){ var m=url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/); if(!m) throw new Error('Bad spreadsheet URL: '+url); return SpreadsheetApp.openById(m[1]); }
function findSourceTabExact_(ss, targetName){
  var exact = ss.getSheetByName(targetName);
  if (exact) return exact;
  var sheets = ss.getSheets();
  for (var i=0;i<sheets.length;i++){
    var n = String(sheets[i].getName()||'').trim();
    if (n.toLowerCase() === targetName.toLowerCase()) return sheets[i];
  }
  return null;
}
function trimToSource_(src, dst){
  var lr=src.getLastRow(), lc=src.getLastColumn();
  if (dst.getMaxRows()>lr) dst.deleteRows(lr+1, dst.getMaxRows()-lr);
  if (dst.getMaxColumns()>lc) dst.deleteColumns(lc+1, dst.getMaxColumns()-lc);
}
function importPropertySheets_CurrentMonth_(master){
  var monthTab = currentMonthTabName_(); // e.g., "Sep25"
  var imported = [];

  for (var i = 0; i < CONFIG.PROPERTY_SOURCES.length; i++) {
    var url = CONFIG.PROPERTY_SOURCES[i];
    try {
      var ss = withRetry_(function(){ return openSource_(url); }, 'open '+url);
      var srcSheet = withRetry_(function(){ return findSourceTabExact_(ss, monthTab); }, 'find '+monthTab);
      if (!srcSheet) { Logger.log('No '+monthTab+' tab in '+url); continue; }

      // Decide the FINAL, SINGLE name we’ll use for this property
      var desiredRaw  = String(srcSheet.getRange('A1').getDisplayValue() || ss.getName() || 'Property').trim();
      var desiredName = preferredTabName_(desiredRaw);          // normalize via PROPERTY_ALIASES
      var canonical   = desiredName;                            // we’ll always use this as the tab name

      // 🚨 Key change: wipe ALL existing tabs that resolve to this canonical location
      deleteSheetsByCanon_(master, canonical);

      // Copy in the new one and name it to the canonical
      var clone = withRetry_(function(){ return srcSheet.copyTo(master); }, 'copyTo '+canonical);
      clone.setName(canonical);

      // Trim any extra rows/cols and stamp A1 with the canonical too
      trimToSource_(srcSheet, clone);
      clone.getRange('A1').setValue(canonical);

      SpreadsheetApp.flush();
      Utilities.sleep(120);
      imported.push(canonical);
      Logger.log('Imported (canonical): '+canonical);
    } catch (e) {
      Logger.log('Import skipped '+url+' -> '+(e && e.message));
    }
  }
  return imported;
}


/* =========================
   COVER (cell ABOVE label)
   ========================= */
function buildCoverAboveLabelsOnly_(master, tabs){
  var cover = master.getSheetByName(CONFIG.COVER_SHEET_NAME) || master.insertSheet(CONFIG.COVER_SHEET_NAME);
  cover.clear();
  cover.getRange(1,1,1,3).setValues([['Building','Due to Owners','Total to Stant']]).setFontWeight('bold');

  var rows=[];
  for (var i=0;i<tabs.length;i++){
    var tab=tabs[i], sh=master.getSheetByName(tab); if(!sh) continue;
    var building = String(sh.getRange('A1').getDisplayValue() || tab).trim();
    var dueA1   = findValueImmediatelyAboveLabelA1_(sh, CONFIG.LABEL_DUE_TO_OWNERS);
    var stantA1 = findValueImmediatelyAboveLabelA1_(sh, CONFIG.LABEL_TOTAL_TO_STANT);
    var dueFx   = dueA1   ? "='" + tab + "'!" + dueA1   : '';
    var stantFx = stantA1 ? "='" + tab + "'!" + stantA1 : '';
    rows.push([building, dueFx, stantFx]);
  }

  if (rows.length){
    cover.getRange(2,1,rows.length,3).setValues(rows);
    for (var r=0;r<rows.length;r++){
      var rr=2+r;
      if (rows[r][1]) cover.getRange(rr,2).setFormula(rows[r][1]).setNumberFormat(CONFIG.CURRENCY_FORMAT);
      if (rows[r][2]) cover.getRange(rr,3).setFormula(rows[r][2]).setNumberFormat(CONFIG.CURRENCY_FORMAT);
    }
    var last=1+rows.length;
    cover.getRange(last+2,1).setValue('TOTALS').setFontWeight('bold');
    cover.getRange(last+2,2).setFormula('=SUM(B2:B' + (last+1) + ')').setNumberFormat(CONFIG.CURRENCY_FORMAT);
    cover.getRange(last+2,3).setFormula('=SUM(C2:C' + (last+1) + ')').setNumberFormat(CONFIG.CURRENCY_FORMAT);
    cover.autoResizeColumns(1,3);
  }
}
function findValueImmediatelyAboveLabelA1_(sh, regex){
  var rng=sh.getDataRange(), v=rng.getValues();
  for (var r=0;r<v.length;r++){
    for (var c=0;c<v[r].length;c++){
      var cell=v[r][c];
      if (cell && cell.toString && regex.test(String(cell))){
        var valueRow=(r+1)-1, valueCol=c+1;
        if (valueRow>=1) return colLetter_(valueCol)+String(valueRow);
      }
    }
  }
  return null;
}

/* =========================
   TAB ORDER
   ========================= */
function reorderTabsPreferred_(master){
  var sheets=master.getSheets(); if(!sheets||!sheets.length) return;
  var info=sheets.map(function(sh){ return { sh:sh, name:sh.getName(), canon:canon_(sh.getName()) }; });
  function move(sh,idx){ master.setActiveSheet(sh); master.moveActiveSheet(idx); }
  var next=1, placed={};
  var cover = info.find(function(i){ return /^cover$/i.test(i.name); });
  if (cover){ move(cover.sh,next++); placed[cover.name]=true; }

  var pref = CONFIG.PREFERRED_ORDER || CONFIG.PREFERRED_ORDER /* keep */;
  if (!pref) pref = [];

  for (var p=0;p<pref.length;p++){
    var entry=pref[p]; if (entry.label && entry.label.toLowerCase()==='cover') continue;
    if (!entry.patterns) continue;
    for (var r=0;r<entry.patterns.length;r++){
      var re=entry.patterns[r];
      var hit=info.find(function(i){ return !placed[i.name] && (re.test(i.name)||re.test(i.canon)); });
      if (hit){ move(hit.sh,next++); placed[hit.name]=true; }
    }
  }
  for (var k=0;k<info.length;k++){
    var i=info[k];
    if (!placed[i.name] && !/^cover$/i.test(i.name)){ move(i.sh,next++); placed[i.name]=true; }
  }
}
function getPropertyTabs_(master){
  var all = master.getSheets().map(function(s){ return s.getName(); });
  return all.filter(function(n){ return n!==CONFIG.COVER_SHEET_NAME && !/legend/i.test(n); });
}

/* =========================
   ENTRY POINTS
   ========================= */
function importOnlyForCurrentMonth(){
  var master=SpreadsheetApp.getActive();
  master.rename(CONFIG.MASTER_NAME);

  // If only one default "Sheet1" exists, create Cover to anchor
  if (master.getSheets().length === 1 && /^sheet\s*1$/i.test(master.getSheets()[0].getName())) {
    master.insertSheet(CONFIG.COVER_SHEET_NAME);
    master.setActiveSheet(master.getSheetByName(CONFIG.COVER_SHEET_NAME));
  }

  removeSheetIfExists_(master,'Sheet1');
  removeSheetIfExists_(master,'Sheet 1');
  removeLegendSheets_(master);

  importPropertySheets_CurrentMonth_(master);

  var tabs = getPropertyTabs_(master);
  buildCoverAboveLabelsOnly_(master, tabs);
  reorderTabsPreferred_(master);

  notify_('Import complete for tab: ' + currentMonthTabName_());
}

// Create/remove monthly import trigger (1st at 6:15 AM ET)
function createMonthlyImportTrigger1(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction && t.getHandlerFunction()==='importOnlyForCurrentMonth') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importOnlyForCurrentMonth').timeBased().onMonthDay(1).atHour(6).nearMinute(15).create();
}
function removeMonthlyImportTrigger1(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction && t.getHandlerFunction()==='importOnlyForCurrentMonth') ScriptApp.deleteTrigger(t);
  });
}

// Simple menu
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('Consolidation (Imports)')
    .addItem('Import Only (Current Month Tab)', 'importOnlyForCurrentMonth')
    .addItem('Create Monthly Import Trigger (1st 6:15 AM ET)', 'createMonthlyImportTrigger1')
    .addItem('Remove Monthly Import Trigger', 'removeMonthlyImportTrigger1')
    .addToUi();
}

/* =========================
   APPEND — NO DEDUPE, ALWAYS ROW 4
   ========================= */
function appendAllDebits_Monthly_(master, tabs, buildingNames, txInfo){
  var total=0;
  for (var j=0;j<tabs.length;j++){
    var tab=tabs[j], sh=master.getSheetByName(tab); if(!sh) continue;

    var headerRow = findHeaderRow_(sh);
    var cols = detectTargetColumnsAtRow_(sh, headerRow);

    var a1 = buildingNames[tab] || tab;
    var canonAliases = aliasesForTab_(tab, a1).map(canonProp_);
    var belongs = txInfo.txRows.filter(function(t){
      var cp = canonProp_(t.property);
      for (var i=0;i<canonAliases.length;i++){ if (cp===canonAliases[i]) return true; }
      return false;
    });
    if (!belongs.length) continue;

    var INSERT_AT = Math.max(CONFIG.WRITABLE_START_ROW, headerRow+1);

    for (var k=0;k<belongs.length;k++){
      var t = belongs[k];
      var amt = (typeof t.debitAmount==='string') ? Number(t.debitAmount.replace(/[^0-9.\-]/g,'')) : t.debitAmount;
      if (amt==null || amt==='' || isNaN(amt)) continue;

      ensureMinRows_(sh, INSERT_AT);
      sh.insertRowsBefore(INSERT_AT, 1);
      var r = INSERT_AT;

      if (cols.unitCol)   sh.getRange(r, cols.unitCol).setValue(String(t.unit||''));
      if (cols.dateCol)   sh.getRange(r, cols.dateCol).setValue(t.date||'');
      if (cols.explCol)   sh.getRange(r, cols.explCol).setValue(String(t.expl||''));
      if (cols.debitsCol) sh.getRange(r, cols.debitsCol).setValue(Number(amt));

      if (cols.markupCol){
        var dv = SpreadsheetApp.newDataValidation().requireCheckbox().build();
        sh.getRange(r, cols.markupCol).setDataValidation(dv).setValue(!!t.markup);
      }
      if (cols.mrevCol && cols.debitsCol && cols.markupCol){
        var fx = '=IF(' + colLetter_(cols.markupCol) + r + ',(' + colLetter_(cols.debitsCol) + r + '*0.1))';
        sh.getRange(r, cols.mrevCol).setFormula(fx);
      }

      total++;
      if (total % (CONFIG.FLUSH_INTERVAL||10)===0){ SpreadsheetApp.flush(); Utilities.sleep(50); }
    }
    try{ enforceTotalsRowSum_(sh); }catch(_){}
    SpreadsheetApp.flush(); Utilities.sleep(50);
  }
  Logger.log('Total debits posted: '+total);
  return total;
}

/* =========================
   TOTALS — SUM B..E over data region
   ========================= */
function enforceTotalsRowSum_(sh){
  var vals=sh.getDataRange().getValues();
  var totalRow=-1;
  for (var r=0;r<vals.length;r++){
    for (var c=0;c<vals[r].length;c++){
      var cell=vals[r][c];
      if (cell && cell.toString && /^totals?\s*:?\s*$/i.test(String(cell).trim())) { totalRow=r+1; break; }
    }
    if (totalRow!==-1) break;
  }
  if (totalRow===-1) return;
  var lastSumRow = totalRow-1; if (lastSumRow<CONFIG.DATA_START_ROW) return;
  ensureMinColumns_(sh,5);
  for (var col=2; col<=5; col++){
    var fx = '=SUM(' + colLetter_(col) + CONFIG.DATA_START_ROW + ':' + colLetter_(col) + lastSumRow + ')';
    sh.getRange(totalRow, col).setFormula(fx).setNumberFormat(CONFIG.CURRENCY_FORMAT);
  }
}

/* =========================
   COVER — cell ABOVE label
   ========================= */
function buildCoverAboveLabelsOnly_(master, tabs){
  var cover = master.getSheetByName(CONFIG.COVER_SHEET_NAME) || master.insertSheet(CONFIG.COVER_SHEET_NAME);
  cover.clear();
  cover.getRange(1,1,1,3).setValues([['Building','Due to Owners','Total to Stant']]).setFontWeight('bold');

  var rows=[];
  for (var i=0;i<tabs.length;i++){
    var tab=tabs[i], sh=master.getSheetByName(tab); if(!sh) continue;
    var building = String(sh.getRange('A1').getDisplayValue() || tab).trim();
    var dueA1   = findValueImmediatelyAboveLabelA1_(sh, CONFIG.LABEL_DUE_TO_OWNERS);
    var stantA1 = findValueImmediatelyAboveLabelA1_(sh, CONFIG.LABEL_TOTAL_TO_STANT);
    var dueFx   = dueA1   ? "='" + tab + "'!" + dueA1   : '';
    var stantFx = stantA1 ? "='" + tab + "'!" + stantA1 : '';
    rows.push([building, dueFx, stantFx]);
  }

  if (rows.length){
    cover.getRange(2,1,rows.length,3).setValues(rows);
    for (var r=0;r<rows.length;r++){
      var rr = 2+r;
      if (rows[r][1]) cover.getRange(rr,2).setFormula(rows[r][1]).setNumberFormat(CONFIG.CURRENCY_FORMAT);
      if (rows[r][2]) cover.getRange(rr,3).setFormula(rows[r][2]).setNumberFormat(CONFIG.CURRENCY_FORMAT);
    }
    var last = 1+rows.length;
    cover.getRange(last+2,1).setValue('TOTALS').setFontWeight('bold');
    cover.getRange(last+2,2).setFormula('=SUM(B2:B' + (last+1) + ')').setNumberFormat(CONFIG.CURRENCY_FORMAT);
    cover.getRange(last+2,3).setFormula('=SUM(C2:C' + (last+1) + ')').setNumberFormat(CONFIG.CURRENCY_FORMAT);
    cover.autoResizeColumns(1,3);
  }
}
function findValueImmediatelyAboveLabelA1_(sh, regex){
  var rng=sh.getDataRange(), v=rng.getValues();
  for (var r=0;r<v.length;r++){
    for (var c=0;c<v[r].length;c++){
      var cell=v[r][c];
      if (cell && cell.toString && regex.test(String(cell))){
        var valueRow = (r+1)-1, valueCol=c+1;
        if (valueRow>=1) return colLetter_(valueCol)+String(valueRow);
      }
    }
  }
  return null;
}

/* =========================
   TAB ORDER / UTIL
   ========================= */
function reorderTabsPreferred_(master){
  var sheets=master.getSheets(); if(!sheets||!sheets.length) return;
  var info=sheets.map(function(sh){ return {sh:sh,name:sh.getName(),canon:canon_(sh.getName())}; });
  function move(sh,idx){ master.setActiveSheet(sh); master.moveActiveSheet(idx); }
  var next=1, placed={};
  var cover = info.find(function(i){return /^cover$/i.test(i.name);});
  if (cover){ move(cover.sh,next++); placed[cover.name]=true; }
  var pref=CONFIG.PREFERRED_ORDER||[];
  for (var p=0;p<pref.length;p++){
    var entry=pref[p]; if (entry.label.toLowerCase()==='cover') continue;
    for (var r=0;r<entry.patterns.length;r++){
      var re=entry.patterns[r];
      var hit=info.find(function(i){ return !placed[i.name] && (re.test(i.name)||re.test(i.canon)); });
      if (hit){ move(hit.sh,next++); placed[hit.name]=true; }
    }
  }
  for (var k=0;k<info.length;k++){ var i=info[k];
    if (!placed[i.name] && !/^cover$/i.test(i.name)){ move(i.sh,next++); placed[i.name]=true; }
  }
}
function getPropertyTabs_(master){
  var all = master.getSheets().map(function(s){return s.getName();});
  return all.filter(function(n){ return n!==CONFIG.COVER_SHEET_NAME && !/legend/i.test(n) && n!==CONFIG.TEMPLATE_TAB_NAME; });
}
function readBuildingNamesFromA1_(master, tabs){
  var map={};
  for (var i=0;i<tabs.length;i++){
    var sh=master.getSheetByName(tabs[i]); var a1 = sh ? String(sh.getRange('A1').getDisplayValue()||'').trim() : '';
    map[tabs[i]] = a1 || tabs[i];
  }
  return map;
}
function safeDeleteSheet_(ss, sh) {
  if (!ss || !sh) return;
  try {
    // Don’t delete if it’s the last sheet
    if (ss.getSheets().length <= 1) { sh.clear(); return; }

    // Don’t delete the active sheet; switch away first
    var active = ss.getActiveSheet();
    if (active && active.getSheetId() === sh.getSheetId()) {
      var alt = ss.getSheets().find(function(s){ return s.getSheetId() !== sh.getSheetId(); });
      if (alt) ss.setActiveSheet(alt);
    }

    // Re-fetch by name right before deleting to avoid stale handle issues
    var latest = ss.getSheetByName(sh.getName());
    if (latest) ss.deleteSheet(latest);
  } catch (e) {
    // If it’s already gone or ID mismatch, just log and continue
    Logger.log('safeDeleteSheet_ skipped: ' + (e && e.message));
  }
}

function removeSheetIfExists_(master, name) {
  try {
    var sh = master.getSheetByName(name);
    if (!sh) return;

    // If only 1 sheet remains, convert it into Cover instead of deleting
    if (master.getSheets().length <= 1) {
      if (name !== CONFIG.COVER_SHEET_NAME) {
        sh.clear();
        sh.setName(CONFIG.COVER_SHEET_NAME);
      }
      return;
    }

    safeDeleteSheet_(master, sh);
  } catch (e) {
    Logger.log('removeSheetIfExists_("' + name + '") -> ' + (e && e.message));
  }
}
function removeLegendSheets_(master) {
  var sheets = master.getSheets();
  var toDelete = sheets.filter(function(s){
    var nm = String(s.getName() || '');
    return /legend/i.test(nm);
  });
  toDelete.forEach(function(s){ safeDeleteSheet_(master, s); });
}


/* =========================
   ENTRY POINTS
   ========================= */
// Run on the 1st (imports current month only), also builds cover + order
function importOnlyForCurrentMonth(){
  var master=SpreadsheetApp.getActive();
  master.rename(CONFIG.MASTER_NAME);
  removeSheetIfExists_(master,'Sheet1'); removeSheetIfExists_(master,'Sheet 1'); removeLegendSheets_(master);
  importPropertySheets_CurrentMonth_(master);
  var tabs = getPropertyTabs_(master);
  buildCoverAboveLabelsOnly_(master, tabs);
  reorderTabsPreferred_(master);
  notify_('Import complete for tab: ' + currentMonthTabName_());
}

// Run on the 19th (post ALL debits, row 4, no dedupe)
function postDebitsMonthlyNow(){
  var master=SpreadsheetApp.getActive();
  master.rename(CONFIG.MASTER_NAME);

  // Optional: keep these to refresh current month tabs before posting
  removeSheetIfExists_(master,'Sheet1'); removeSheetIfExists_(master,'Sheet 1'); removeLegendSheets_(master);
  importPropertySheets_CurrentMonth_(master);

  var tabs=getPropertyTabs_(master);
  var txInfo=getTransactionsDynamic_();
  var buildingNames=readBuildingNamesFromA1_(master, tabs);

  var posted = appendAllDebits_Monthly_(master, tabs, buildingNames, txInfo);
  buildCoverAboveLabelsOnly_(master, tabs);
  reorderTabsPreferred_(master);
  notify_('Debits posted: '+posted);
}

// Create monthly triggers
function createMonthlyImportTrigger1(){
  // 1st of each month at 6:15 AM ET
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction && t.getHandlerFunction()==='importOnlyForCurrentMonth') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importOnlyForCurrentMonth').timeBased().onMonthDay(1).atHour(6).nearMinute(15).create();
}
function createMonthlyPostTrigger19(){
  // 19th of each month at 9:15 PM ET
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction && t.getHandlerFunction()==='postDebitsMonthlyNow') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('postDebitsMonthlyNow').timeBased().onMonthDay(19).atHour(21).nearMinute(15).create();
}
function removeAllMonthlyTriggers(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    var fn = t.getHandlerFunction ? t.getHandlerFunction() : '';
    if (fn==='importOnlyForCurrentMonth' || fn==='postDebitsMonthlyNow') ScriptApp.deleteTrigger(t);
  });
}
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('Consolidation')
    .addItem('Import Only (Current Month Tab)', 'importOnlyForCurrentMonth')
    .addItem('Post Debits (Row 4, No Dedupe)', 'postDebitsMonthlyNow')
    .addSeparator()
    .addItem('Create Monthly Import Trigger (1st 6:15 AM ET)', 'createMonthlyImportTrigger1')
    .addItem('Create Monthly Post Trigger (19th 9:15 PM ET)', 'createMonthlyPostTrigger19')
    .addItem('Remove Monthly Triggers', 'removeAllMonthlyTriggers')
    .addToUi();
}
```
