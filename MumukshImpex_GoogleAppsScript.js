// ============================================================
//  MUMUKSH IMPEX LLP — Production Report Google Sheets Backend
//  Version 2.0 — Fixed
// ============================================================
//
//  ⚠️  DO NOT RUN "doPost" FROM THE EDITOR — it will always error.
//
//  CORRECT STEPS (do in this exact order):
//  ─────────────────────────────────────────────────────────
//  STEP 1: Select "setupSheets" from function dropdown → ▶ Run
//          (Grant permissions when Google asks)
//          → This creates all 4 sheets with headers
//
//  STEP 2: Select "testSave" from dropdown → ▶ Run
//          → This inserts 1 dummy row to verify sheets work
//
//  STEP 3: Deploy → New Deployment
//          • Type: Web App
//          • Execute as: Me
//          • Who has access: Anyone
//          → Copy the Web App URL
//
//  STEP 4: Set ADMIN_EMAIL variable at top of this file
//          to the admin's email address
//
//  STEP 5: Paste the URL into the production app
//          (Step 5 → Export → Google Sheet URL field, Admin only)
// ============================================================

// ─────────────────────────────────────────────────────────
//  ✉️  ADMIN EMAIL — change this to your email address
// ─────────────────────────────────────────────────────────
var ADMIN_EMAIL = 'admin@mumukshimpex.com';   // ← Change this

var SHEET_MAIN     = 'Production_Reports';
var SHEET_SKU      = 'SKU_Data';
var SHEET_DOWNTIME = 'Downtime_Log';
var SHEET_RM       = 'RM_Consumption';

var HEADERS_MAIN = [
  'Report_ID','Submitted_At','Submitted_By',
  'Date','Shift','MC','Size','MC_Speed_RPM','AVG_Speed',
  'Use_Time_min','Run_Time_min','Downtime_min','EFF_Percent',
  'KWH_Opening','KWH_Closing','KWH_Diff',
  'Count_Start','Count_End','Count_Diff',
  'Operator','ASS_Operator','RM_Operator_1','RM_Operator_2','Line_Checker',
  'Handover','CA_Dai','EA_Dai','WL_Dai','Total_Downtime_min',
  'Hold_Brand','Hold_Configure','Hold_Bag_Count','Hold_Pieces','Hold_Remark',
  'Quality_B_Grade','Quality_C_Grade','Quality_Unfiltered','Quality_Remarks',
  'Pkg_Non_Woven','Pkg_White_Poly','Pkg_Core','Pkg_PE','Pkg_Pulf',
  'Pkg_Cartoon','Pkg_Plan_Gutta','Pkg_Sealing_Cut',
  'Verify_Operator','Verify_Quality','Verify_Production','Verify_Store'
];
var HEADERS_SKU = [
  'Report_ID','Date','Shift','MC',
  'SKU_Number','SKU_Name','Pulf_WT','Pulf_Rat','Sap_WT','Sap_Rat',
  'TPC','GPC','B_Grade','C_Grade','T_Rej','Rej_Percent'
];
var HEADERS_DT = [
  'Report_ID','Date','Shift','MC',
  'DT_Row','From_Time','To_Time','Duration_min','Reason'
];
var HEADERS_RM = [
  'Report_ID','Date','Shift','MC',
  'RM_Name','UOM','Opening','Receive','Total','Consumed','Closing','Supplier'
];

// ─────────────────────────────────────────────────────────
//  STEP 1 — RUN THIS FIRST from the editor dropdown
//  Creates all 4 sheets with styled headers
// ─────────────────────────────────────────────────────────
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var toDelete = [];

  // Collect existing sheets to delete AFTER creating new ones.
  // Google Sheets requires at least 1 sheet to exist at all times,
  // so we must create first then delete — never delete all at once.
  [SHEET_MAIN, SHEET_SKU, SHEET_DOWNTIME, SHEET_RM].forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) toDelete.push(s);
  });

  // Step 1: Insert 4 new sheets with temp names to avoid conflicts
  var s1 = ss.insertSheet(SHEET_MAIN     + '_new');
  var s2 = ss.insertSheet(SHEET_SKU      + '_new');
  var s3 = ss.insertSheet(SHEET_DOWNTIME + '_new');
  var s4 = ss.insertSheet(SHEET_RM       + '_new');

  // Step 2: Now safe to delete old sheets (new ones exist)
  toDelete.forEach(function(s) { ss.deleteSheet(s); });

  // Step 3: Rename to correct names and write headers
  s1.setName(SHEET_MAIN);     writeHeader(s1, HEADERS_MAIN);
  s2.setName(SHEET_SKU);      writeHeader(s2, HEADERS_SKU);
  s3.setName(SHEET_DOWNTIME); writeHeader(s3, HEADERS_DT);
  s4.setName(SHEET_RM);       writeHeader(s4, HEADERS_RM);

  // Step 4: Remove default blank Sheet1 if still present
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getNumSheets() > 1) ss.deleteSheet(def);

  SpreadsheetApp.getUi().alert(
    '\u2705 All 4 sheets created!\n\n' +
    '\u2022 Production_Reports\n' +
    '\u2022 SKU_Data\n' +
    '\u2022 Downtime_Log\n' +
    '\u2022 RM_Consumption\n\n' +
    'Run "testSave" next to verify, then Deploy as Web App.'
  );
}

// ─────────────────────────────────────────────────────────
//  STEP 2 — RUN THIS to verify sheets work correctly
//  Inserts one dummy test row into all 4 sheets
// ─────────────────────────────────────────────────────────
function testSave() {
  var dummy = {
    formData: {
      date:'2025-03-11', shift:'A', mc:'MC-01', size:'M',
      mcSpeed:'320', avgSpeed:'12.50', useTime:'480', runTime:'440',
      dTime:'40', eff:'91.7%', kwhOp:'1000', kwhCl:'1120', kwhDiff:'120',
      cntStart:'10000', cntEnd:'16000', cntDiff:'6000',
      operator:'Test Operator', assOp:'Asst Operator',
      rmOp1:'RM Op 1', rmOp2:'RM Op 2', lChecker:'Line Checker',
      handover:'RUN', caDai:'0.76', eaDai:'0.34', wlDai:'0',
      dtTotal: 40, submittedBy:'Test User',
      skus:[
        {name:'Test SKU 1',pulfWt:'5',pulfRat:'2',sapWt:'3',sapRat:'1',
         tpc:'500',gpc:'6000',bGrade:'10',cGrade:'5',tRej:'15',rejPct:'0.25%'},
        {name:'',pulfWt:'',pulfRat:'',sapWt:'',sapRat:'',
         tpc:'',gpc:'',bGrade:'',cGrade:'',tRej:'',rejPct:''},
        {name:'',pulfWt:'',pulfRat:'',sapWt:'',sapRat:'',
         tpc:'',gpc:'',bGrade:'',cGrade:'',tRej:'',rejPct:''},
        {name:'',pulfWt:'',pulfRat:'',sapWt:'',sapRat:'',
         tpc:'',gpc:'',bGrade:'',cGrade:'',tRej:'',rejPct:''}
      ],
      dtRows:[
        {from:'08:00',to:'08:40',min:'40',reason:'Machine Breakdown — Mechanical Fault'}
      ]
    },
    holdData: {
      brand:'TestBrand', configure:'Config-A', bag:'10', noPic:'500',
      holdRemark:'Test remark', bGrade:'10', cGrade:'5',
      unselved:'2', remarks:'All OK',
      nw:'3.5', whitePoly:'12', corNos:'50', pe:'8', pulf:'6',
      cartoon:'20', planGutta:'4', sealCut:'2',
      op:'Test OP', qly:'495', production:'495', storeVerify:'Yes',
      rmData:[
        {rm:'Sap',  uom:'KG', opening:'100', receive:'50',  total:'150', consumed:'120', closing:'30', supplier:'Supplier A'},
        {rm:'CA',   uom:'KG', opening:'20',  receive:'10',  total:'30',  consumed:'25',  closing:'5',  supplier:'Supplier B'},
        {rm:'EA',   uom:'KG', opening:'15',  receive:'5',   total:'20',  consumed:'18',  closing:'2',  supplier:'Supplier C'}
      ]
    }
  };

  var id = saveAllSheets(dummy);
  SpreadsheetApp.getUi().alert(
    '✅ Test record saved!\n\n' +
    'Report ID: ' + id + '\n\n' +
    'Check all 4 sheets — you should see 1 row of test data.\n\n' +
    'If data looks good → Deploy as Web App!'
  );
}

// ─────────────────────────────────────────────────────────
//  HTTP ENDPOINTS (called automatically — do NOT run manually)
// ─────────────────────────────────────────────────────────

function doGet(e) {
  // Admin Dashboard: fetch all sheet data
  if (e && e.parameter && e.parameter.action === 'getData') {
    try {
      return getSheetData();
    } catch(err) {
      Logger.log('getData error: ' + err.message);
      return ContentService
        .createTextOutput(JSON.stringify({error: err.message}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({status:'ok', message:'Mumuksh Impex API is live ✅'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────
//  Returns all sheet data as JSON — used by Admin Dashboard
//  Called via: GET <webAppUrl>?action=getData
// ─────────────────────────────────────────────────────────
function getSheetData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // DATE COLUMNS — Sheets auto-converts date strings to Date objects.
  // We must check instanceof Date and re-format to yyyy-MM-dd to match
  // what the HTML date filter inputs produce.
  var DATE_COLS  = {'Date': true};
  var TS_COLS    = {'Submitted_At': true};

  function sheetToJSON(name) {
    var s = ss.getSheetByName(name);
    if (!s || s.getLastRow() < 2) return [];
    var vals = s.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h); });
    return vals.slice(1).map(function(row) {
      var obj = {};
      hdrs.forEach(function(h, i) {
        var val = row[i];
        if (val === null || val === undefined || val === '') {
          obj[h] = '';
        } else if (val instanceof Date) {
          // Re-format dates so they match HTML <input type="date"> yyyy-MM-dd
          if (DATE_COLS[h]) {
            obj[h] = Utilities.formatDate(val, 'Asia/Kolkata', 'yyyy-MM-dd');
          } else {
            obj[h] = Utilities.formatDate(val, 'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss');
          }
        } else {
          obj[h] = String(val);
        }
      });
      return obj;
    });
  }

  var result = {
    reports:  sheetToJSON(SHEET_MAIN),
    downtime: sheetToJSON(SHEET_DOWNTIME),
    skus:     sheetToJSON(SHEET_SKU),
    rm:       sheetToJSON(SHEET_RM)
  };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────
//  EMAIL NOTIFICATION TO ADMIN
// ─────────────────────────────────────────────────────────
function sendAdminNotification(reportId, fd) {
  try {
    var subject = '\uD83D\uDCCA New Production Report: ' + fd.shift + ' Shift | ' + fd.mc + ' | ' + fd.date;
    var body = [
      'Hello Admin,',
      '',
      'A new production report has been submitted to Google Sheets.',
      '',
      '\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014',
      'Report ID    : ' + reportId,
      'Date         : ' + fd.date,
      'Shift        : ' + fd.shift,
      'Machine      : ' + (fd.mc || '\u2014'),
      'Size         : ' + (fd.size || '\u2014'),
      'Operator     : ' + (fd.operator || '\u2014'),
      'Submitted By : ' + (fd.submittedBy || '\u2014'),
      'Submitted At : ' + (fd.submittedAt || '\u2014'),
      '\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014',
      'PERFORMANCE',
      'EFF%         : ' + (fd.eff || '\u2014'),
      'Run Time     : ' + (fd.runTime || '\u2014') + ' min',
      'Downtime     : ' + (fd.dTime || '\u2014') + ' min',
      '\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014',
      '',
      'Please log in to the production app (Admin) and download the PDF report to send to management.',
      '',
      'This is an automated message from Mumuksh Impex LLP Production System.'
    ].join('\n');

    GmailApp.sendEmail(ADMIN_EMAIL, subject, body);
    Logger.log('Notification email sent to ' + ADMIN_EMAIL);
  } catch(err) {
    // Email failure should NOT block the save — just log it
    Logger.log('Email notification failed: ' + err.message);
  }
}

function doPost(e) {
  // This function CANNOT be run from the editor.
  // It is called automatically when the HTML app sends data.
  try {
    var payload = JSON.parse(e.postData.contents);
    var id = saveAllSheets(payload);
    sendAdminNotification(id, payload.formData || {});
    return ContentService
      .createTextOutput(JSON.stringify({status:'success', reportId: id}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    Logger.log('Error: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({status:'error', message: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────────────────
//  CORE SAVE LOGIC
// ─────────────────────────────────────────────────────────

function saveAllSheets(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fd = payload.formData || {};
  var hr = payload.holdData || {};
  var ts = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd-HHmmss');
  var reportId = 'MMI-' + ts;
  var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss');

  // ── Sheet 1: Production_Reports ──
  var s1 = getSheet(ss, SHEET_MAIN, HEADERS_MAIN);
  s1.appendRow([
    reportId, now, fd.submittedBy||'',
    fd.date||'', fd.shift||'', fd.mc||'', fd.size||'',
    fd.mcSpeed||'', fd.avgSpeed||'',
    fd.useTime||'', fd.runTime||'', fd.dTime||'', (parseFloat((fd.eff||'').toString().replace('%',''))||''),
    fd.kwhOp||'', fd.kwhCl||'', fd.kwhDiff||'',
    fd.cntStart||'', fd.cntEnd||'', fd.cntDiff||'',
    fd.operator||'', fd.assOp||'', fd.rmOp1||'', fd.rmOp2||'', fd.lChecker||'',
    fd.handover||'', fd.caDai||'', fd.eaDai||'', fd.wlDai||'', fd.dtTotal||0,
    hr.brand||'', hr.configure||'', hr.bag||'', hr.noPic||'', hr.holdRemark||'',
    hr.bGrade||'', hr.cGrade||'', hr.unselved||'', hr.remarks||'',
    hr.nw||'', hr.whitePoly||'', hr.corNos||'', hr.pe||'', hr.pulf||'',
    hr.cartoon||'', hr.planGutta||'', hr.sealCut||'',
    hr.op||'', hr.qly||'', hr.production||'', hr.storeVerify||''
  ]);
  styleRow(s1, HEADERS_MAIN.length);

  // ── Sheet 2: SKU_Data ──
  var s2 = getSheet(ss, SHEET_SKU, HEADERS_SKU);
  (fd.skus || []).forEach(function(s, i) {
    if (!s || (!s.name && !s.tpc && !s.gpc && !s.pulfWt && !s.sapWt)) return;
    s2.appendRow([
      reportId, fd.date||'', fd.shift||'', fd.mc||'',
      'SKU '+(i+1), s.name||'',
      s.pulfWt||'', s.pulfRat||'', s.sapWt||'', s.sapRat||'',
      s.tpc||'', s.gpc||'', s.bGrade||'', s.cGrade||'', s.tRej||'', s.rejPct||''
    ]);
    styleRow(s2, HEADERS_SKU.length);
  });

  // ── Sheet 3: Downtime_Log ──
  var s3 = getSheet(ss, SHEET_DOWNTIME, HEADERS_DT);
  (fd.dtRows || []).forEach(function(r, i) {
    if (!r || (!r.from && !r.reason)) return;
    s3.appendRow([
      reportId, fd.date||'', fd.shift||'', fd.mc||'',
      i+1, r.from||'', r.to||'', r.min||'', r.reason||''
    ]);
    styleRow(s3, HEADERS_DT.length);
  });

  // ── Sheet 4: RM_Consumption ──
  var s4 = getSheet(ss, SHEET_RM, HEADERS_RM);
  (hr.rmData || []).forEach(function(r) {
    if (!r || !r.rm) return;
    s4.appendRow([
      reportId, fd.date||'', fd.shift||'', fd.mc||'',
      r.rm||'', r.uom||'KG',
      r.opening||'', r.receive||'', r.total||'',
      r.consumed||'', r.closing||'', r.supplier||''
    ]);
    styleRow(s4, HEADERS_RM.length);
  });

  return reportId;
}

// ─────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────

function getSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = createSheet(ss, name, headers);
  } else if (sheet.getLastRow() === 0) {
    writeHeader(sheet, headers);
  }
  return sheet;
}

function createSheet(ss, name, headers) {
  var sheet = ss.insertSheet(name);
  writeHeader(sheet, headers);
  return sheet;
}

function writeHeader(sheet, headers) {
  var r = sheet.getRange(1, 1, 1, headers.length);
  r.setValues([headers]);
  r.setBackground('#0d1117');
  r.setFontColor('#f0a500');
  r.setFontWeight('bold');
  r.setFontSize(10);
  r.setFontFamily('Arial');
  r.setBorder(true,true,true,true,true,true,
    '#30363d', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 34);
  for (var i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, i <= 2 ? 170 : 120);
  }
}

function styleRow(sheet, colCount) {
  var row = sheet.getLastRow();
  var r   = sheet.getRange(row, 1, 1, colCount);
  r.setBackground(row % 2 === 0 ? '#f0f4f8' : '#ffffff');
  r.setFontSize(10);
  r.setFontFamily('Arial');
  r.setBorder(false,false,true,false,false,false,
    '#dde1e7', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(row, 24);
}
