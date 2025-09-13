/** RT Bootstrap — creates Sheets, seeds data, and injects helper scripts.
 *  Required services (Resources > Advanced Google services):
 *  - Drive API
 *  - Apps Script API
 */

// ---------- CONFIG ----------
const ROOT_FOLDER_NAME = 'Web3WizardsHub';
const EXPORTS_SUBFOLDER = 'Exports';
const FEATURED_SUBFOLDER = 'FeaturedListings';

const ENABLE_APPSHEET_API = true;
const APPSHEET_API_KEY = process.env.APPSHEET_API_KEY || ''; // Load from environment variable

// Optional: if you want to hit AppSheet API later, drop key here and set ENABLE_APPSHEET_API = true
const ENABLE_APPSHEET_API = true;
const APPSHEET_API_KEY = 'ewv8A-hs7TM-r81Ny-LBW87-yBwYR-55IS2-zXPUn-rrE93'; // put your key here when ready

// ---------- ENTRYPOINT ----------
function runSetup() {
  const root = getOrCreateFolder_(ROOT_FOLDER_NAME);
  const exports = getOrCreateChild_(root, EXPORTS_SUBFOLDER);
  const featured = getOrCreateChild_(root, FEATURED_SUBFOLDER);

  // 1) Spreadsheets to create/seed
  const books = [
    makeBookSpec_(
      'Recovered Treasures Inventory',
      [
        {name:'Inventory', headers:['ItemID','Title','Category','Price','Qty','ImageURL','Vendor','Status'], samples:[
          ['INV-001','Vintage Stamp Lot','Stamps',49.99,1,'https://example.com/stamp.jpg','House','Listed'],
          ['INV-002','Antique Coin','Coins',129.00,1,'','Consignment','Draft'],
        ]},
        {name:'Images', headers:['ItemID','ImageURL','Notes'], samples:[
          ['INV-001','https://example.com/stamp.jpg','hero'],
        ]},
        {name:'Vendors', headers:['Vendor','Contact','Phone','Email'], samples:[
          ['House','N/A','',''],
          ['Consignment','John Doe','',''],
        ]},
      ]
    ),
    makeBookSpec_(
      'Central Management',
      [
        {name:'Intake', headers:['IntakeID','Date','Source','Contact','ItemCount','Notes'], samples:[
          ['IN-0001',new Date(),'Seller','Jane',12,'garage pickup'],
        ]},
        {name:'Metrics', headers:['Date','Listings','Sales','GMV','AvgDaysToSell'], samples:[
        {name:'Sync', headers:['Action','Target','When','Status','Details'], 
         samples: SITE_TARGETS.map(site => ['Push Featured', site, 'daily 2am', 'pending', 'first run'])},
        {name:'Sync', headers:['Action','Target','When','Status','Details'], samples: SITE_TARGETS.map(site => [
          'Push Featured', site, 'daily 2am', 'pending', 'first run'
        ])},
        {name:'Registry', headers:['Site','URL','SheetID','Notes'], samples:[]},
      ]
    ),
    makeBookSpec_(
      'Philatelic-Collectors-Companion',
      [
        {name:'Stamps', headers:['StampNo','Country','Year','Denom','Condition','ScottRef','ImageURL','Tags'], samples:[
          ['27','USA','1861','10¢','Fine','Scott 27','', 'grill, early-issue'],
        ]},
        {name:'Catalog', headers:['ScottRef','Title','Notes','Link'], samples:[
          ['Scott 27','10¢ Washington','Type V','https://example.com/ref'],
        ]},
      ]
    )
  ];

  const created = books.map(spec => createAndSeed_(root, spec));

  // 2) Inject container-bound helper script into each spreadsheet
  created.forEach(c => attachBoundHelpers_(c.id));

  // 3) Register central registry with basic links
  const central = created.find(b => b.name === 'Central Management');
  if (central) {
    const sh = SpreadsheetApp.openById(central.id).getSheetByName('Registry');
    created.forEach(b => {
      if (b.name !== 'Central Management') {
        sh.appendRow([b.name, b.url, b.id, '']);
      }
    });
  }

  // 4) (Optional) AppSheet API stubs
  if (ENABLE_APPSHEET_API && APPSHEET_API_KEY) {
    // Example: createAppSheetAppFromSheet_(created[0].id, 'RecoveredTreasures Inventory App');
  }

  SpreadsheetApp.flush();
  Logger.log('Bootstrap complete.');
  Browser.msgBox('Bootstrap complete.\nRoot: ' + root.getUrl() + '\nExports: ' + exports.getUrl());
}

// ---------- BUILDERS ----------
function makeBookSpec_(title, tabs){ return {title, tabs}; }

function createAndSeed_(parentFolder, bookSpec){
  const ss = SpreadsheetApp.create(bookSpec.title);

  // Move file into your root folder (requires Advanced Drive API enabled)
  Drive.Files.update({ parents: [{ id: parentFolder.getId() }] }, ss.getId());

  const tabs = bookSpec.tabs || [];
  const base = ss.getActiveSheet(); // the auto-created default sheet

  if (tabs.length === 0) {
    base.setName('Data');
    return { id: ss.getId(), url: ss.getUrl(), name: bookSpec.title };
  }

  // 1) Reuse default sheet for the FIRST tab
  const first = tabs[0];
  base.setName(first.name);
  if (first.headers && first.headers.length) {
    base.getRange(1,1,1,first.headers.length).setValues([first.headers]);
  }
  if (first.samples && first.samples.length) {
    base.getRange(2,1,first.samples.length, first.headers.length).setValues(first.samples);
  }
  base.setFrozenRows(1);
  autoWidth_(base, first.headers ? first.headers.length : 10);

  // 2) Create remaining tabs
  for (let i = 1; i < tabs.length; i++) {
    const t = tabs[i];
    const s = ss.insertSheet(t.name);
    if (t.headers && t.headers.length) {
      s.getRange(1,1,1,t.headers.length).setValues([t.headers]);
    }
    if (t.samples && t.samples.length) {
      s.getRange(2,1,t.samples.length,t.headers.length).setValues(t.samples);
    }
    s.setFrozenRows(1);
    autoWidth_(s, t.headers ? t.headers.length : 10);
  }

  SpreadsheetApp.flush();
  return { id: ss.getId(), url: ss.getUrl(), name: bookSpec.title };
}


// ---------- HELPERS ----------
function autoWidth_(sheet, colCount){
  for (let c=1;c<=colCount;c++) sheet.autoResizeColumn(c);
}

function getOrCreateFolder_(name){
  const existing = DriveApp.getFoldersByName(name);
  return existing.hasNext() ? existing.next() : DriveApp.createFolder(name);
}
function getOrCreateChild_(parent, name){
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

// ---------- BOUND SCRIPT INJECTION ----------
/**
 * Creates a container-bound Apps Script inside a Spreadsheet and writes helper code.
 * Requires: Apps Script API enabled in this project.
 */
/**
 * Creates a container-bound script inside the given Spreadsheet and writes helper code.
 * Prereqs (in this very project): Services ➜ + ➜ enable **Apps Script API** (aka “Script”) and **Drive API**.
 * Also enable those APIs in your Cloud project.
 */
// 1) Safe attach: use Advanced Service if present; else write helper code to Drive for manual paste.
function attachBoundHelpers_(spreadsheetId){
  if (typeof Script !== 'undefined' && Script.Projects && Script.Projects.create) {
    // Advanced Service is available → create bound project
    const project = Script.Projects.create({ title: 'RT Sheet Helpers', parentId: spreadsheetId });
    const files = [
      { name: 'Code', type: 'SERVER_JS', source: boundHelperSource_() },
      { name: 'appsscript', type: 'JSON',
        source: JSON.stringify({ timeZone: Session.getScriptTimeZone(), exceptionLogging: 'STACKDRIVER' }) }
    ];
    Script.Projects.updateContent({ files }, project.scriptId);
    Logger.log('Bound helpers created: ' + project.scriptId);
    return;
  }
  // Fallback: write helper code to a .txt file beside the Sheet
  emitHelperCodeFile_(spreadsheetId);
}

// 2) Helper: writes “RT-Helpers.code.txt” next to the spreadsheet with the code to paste.
function emitHelperCodeFile_(spreadsheetId){
  const file = DriveApp.getFileById(spreadsheetId);
  const parent = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
  const name = 'RT-Helpers.code.txt';
  // remove stale copy if any
  const existing = parent.getFilesByName(name);
  while (existing.hasNext()) { existing.next().setTrashed(true); }
  const blob = Utilities.newBlob(boundHelperSource_(), 'text/plain', name);
  const txt = parent.createFile(blob);
  Logger.log('Advanced Service not available. Wrote helper code file: ' + txt.getUrl());
}

/** Code that will live inside each spreadsheet’s bound project. */
function boundHelperSource_(){
  return `
/** ==== RT Bound Helpers (lives inside the spreadsheet) ==== */

// Adds a menu every time the spreadsheet is opened.
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('RT Tools')
    .addItem('Reverse Image Search (selected cells)','menuReverseImageSearch')
    .addItem('Export CSV Bundle (all tabs)','menuExportCsvBundle')
    .addSeparator()
    .addItem('Sync Featured Listings (demo)','menuSyncFeatured')
    .addToUi();
}

// Builds Google Images "search by image" URLs for selected cells containing URLs.
function menuReverseImageSearch(){
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  const out = [];
  for (let r=0;r<values.length;r++){
    const row = [];
    for (let c=0;c<values[0].length;c++){
      const v = String(values[r][c]||'').trim();
      row.push(v ? 'https://www.google.com/searchbyimage?image_url=' + encodeURIComponent(v) : '');
    }
    out.push(row);
  }
  // write starting next column to avoid overwrite
  sheet.getRange(range.getRow(), range.getColumn() + range.getNumColumns(), out.length, out[0].length).setValues(out);
  SpreadsheetApp.getUi().alert('Reverse-image search URLs added to the right of your selection.');
}

// Exports every visible tab to Drive as CSV into sibling "Exports" folder.
function menuExportCsvBundle(){
  const ss = SpreadsheetApp.getActive();
  const file = DriveApp.getFileById(ss.getId());
  const parent = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();

  // find or create Exports folder beside this file
  let exports = null;
  const it = parent.getFoldersByName('Exports');
  exports = it.hasNext() ? it.next() : parent.createFolder('Exports');

  ss.getSheets().forEach(sh => {
    if (!sh.isSheetHidden()){
      const csv = sheetToCsv_(sh);
      const blob = Utilities.newBlob(csv, 'text/csv', ss.getName() + ' - ' + sh.getName() + '.csv');
      exports.createFile(blob);
    }
  });
  SpreadsheetApp.getUi().alert('CSV bundle exported to: ' + exports.getUrl());
}

function sheetToCsv_(sheet){
  const range = sheet.getDataRange();
  const vals = range.getValues();
  return vals.map(row => row.map(cell => {
    const s = (cell===null||cell===undefined) ? '' : String(cell);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s;
  }).join(',')).join('\\n');
}

// Demo sync — copies first 100 "listed" items into sibling FeaturedListings folder as CSV
function menuSyncFeatured(){
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName('Inventory');
  if (!inv) { SpreadsheetApp.getUi().alert('No "Inventory" tab found.'); return; }

  const file = DriveApp.getFileById(ss.getId());
  const parent = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
  let featured = parent.getFoldersByName('FeaturedListings');
  featured = featured.hasNext() ? featured.next() : parent.createFolder('FeaturedListings');

  const data = inv.getDataRange().getValues();
  const head = data.shift();
  const idxStatus = head.indexOf('Status');
  const listed = data.filter(r => (r[idxStatus]||'').toString().toLowerCase().indexOf('listed') >= 0).slice(0,100);
  const csv = [head].concat(listed).map(row => row.map(cell => {
    const s = (cell===null||cell===undefined) ? '' : String(cell);
    return /[",\\n]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s;
  }).join(',')).join('\\n');
  const blob = Utilities.newBlob(csv, 'text/csv', 'featured.csv');
  featured.createFile(blob);
  SpreadsheetApp.getUi().alert('Pushed featured.csv to: ' + featured.getUrl());
}
`;
}

// ---------- (Optional) AppSheet API stubs ----------
function createAppSheetAppFromSheet_(sheetId, appName){
  // Requires APPSHEET_API_KEY + AppSheet account configured for API access.
  const url = 'https://api.appsheet.com/api/v2/apps';
  const payload = {
    Name: appName,
    Sources: [{ DataSourceType: 'GoogleSheets', SpreadsheetId: sheetId }]
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: {'ApplicationAccessKey': APPSHEET_API_KEY},
    muteHttpExceptions: true
  });
  Logger.log(res.getResponseCode() + ' ' + res.getContentText());
}

// ---------- (Reference) Sites embeds ----------
function sitesEmbedSnippet_(appUrl, sheetUrl){
  return `
  <!-- Put this in Google Sites (Embed > Website) -->
  <div style="display:grid;gap:12px">
    <iframe src="${appUrl}" style="width:100%;height:80vh;border:0" allow="fullscreen"></iframe>
    <iframe src="${sheetUrl}" style="width:100%;height:60vh;border:0"></iframe>
  </div>`;
}
