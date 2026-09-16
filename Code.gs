/**
 * দাওয়াতী রিপোর্ট ব্যবস্থাপনা সিস্টেম - Backend (Google Apps Script)
 * ডাটা সংরক্ষণ হবে Google Sheet এ।
 *
 * ইনস্টলেশন নির্দেশনা (README.md এ বিস্তারিত আছে):
 * ১. script.google.com এ নতুন প্রজেক্ট বানান
 * ২. এই কোডটি Code.gs এ পেস্ট করুন
 * ৩. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * ৪. Deploy করে যে URL পাবেন সেটা index.html এর CONFIG.API_URL এ বসান
 */

// ==================== কনফিগারেশন ====================
const SHEET_NAMES = {
  SETTINGS: 'Settings',       // মূল শাখা সেটাপ
  PROGRAMS: 'Programs',       // কর্মসূচী তালিকা
  BRANCHES: 'Branches',       // শাখা তালিকা
  REPORTS: 'Reports',         // দৈনিক রিপোর্ট এন্ট্রি
  USERS: 'Users'              // লগইন ইউজার (এডমিন + শাখা)
};

// ==================== এন্ট্রি পয়েন্ট ====================

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let result;
  try {
    ensureSheetsExistCached();

    let params = {};
    if (e.parameter && Object.keys(e.parameter).length > 0) {
      params = e.parameter;
    }
    if (e.postData && e.postData.contents) {
      try {
        const body = JSON.parse(e.postData.contents);
        params = Object.assign({}, params, body);
      } catch (err) {
        // postData not JSON, ignore
      }
    }

    const action = params.action;
    if (!action) {
      result = { success: false, message: 'action প্যারামিটার প্রয়োজন' };
    } else {
      result = routeAction(action, params);
    }
  } catch (err) {
    result = { success: false, message: 'সার্ভার ত্রুটি: ' + err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function routeAction(action, params) {
  switch (action) {
    // ---------- অথেনটিকেশন ----------
    case 'login': return login(params);

    // ---------- প্রাথমিক ডাটা (একসাথে) ----------
    case 'getBootstrap': return getBootstrap();
    case 'getLoginInfo': return getLoginInfo();
    case 'forceRefresh': return forceRefreshData();

    // ---------- মূল শাখা সেটাপ ----------
    case 'getSettings': return getSettings();
    case 'saveSettings': return saveSettings(params);

    // ---------- কর্মসূচী ----------
    case 'getPrograms': return getPrograms();
    case 'saveProgram': return saveProgram(params);
    case 'deleteProgram': return deleteProgram(params);

    // ---------- শাখা ----------
    case 'getBranches': return getBranches();
    case 'saveBranch': return saveBranch(params);
    case 'deleteBranch': return deleteBranch(params);

    // ---------- রিপোর্ট ----------
    case 'getReports': return getReports(params);
    case 'saveReport': return saveReport(params);
    case 'deleteReport': return deleteReport(params);

    // ---------- ড্যাশবোর্ড / সারসংক্ষেপ ----------
    case 'getDashboard': return getDashboard(params);
    case 'getBranchReportSummary': return getBranchReportSummary(params);
    case 'getMissingBranches': return getMissingBranches(params);

    // ---------- প্রিন্ট রিপোর্ট ----------
    case 'getBranchDateReport': return getBranchDateReport(params);
    case 'getBranchPeriodReport': return getBranchReportSummary(params);
    case 'getMainBranchDateReport': return getMainBranchDateReport(params);
    case 'getMainBranchPeriodReport': return getMainBranchPeriodReport(params);

    default:
      return { success: false, message: 'অজানা action: ' + action };
  }
}

// ==================== শীট প্রস্তুতি ====================

let _ssCache = null;
function getSS() {
  if (!_ssCache) _ssCache = SpreadsheetApp.getActiveSpreadsheet();
  return _ssCache;
}

// প্রতিটি রিকোয়েস্টে শীট আছে কিনা / মাইগ্রেশন লাগবে কিনা যাচাই করা অনেক সময় নেয়
// (একাধিক getSheetByName + getRange কল)। এটা মাত্র একবার (৬ ঘণ্টায় একবার) চেক করাই যথেষ্ট,
// তাই ক্যাশে "সব ঠিক আছে" ফ্ল্যাগ রাখা হয় — এতে প্রতিটি লগইন/সেভ/লোড অনেক দ্রুত হয়ে যায়।
function ensureSheetsExistCached() {
  const cache = CacheService.getScriptCache();
  if (cache.get('sheetsReady') === '1') return;
  ensureSheetsExist();
  try { cache.put('sheetsReady', '1', 21600); } catch (e) { /* ক্যাশ ব্যর্থ হলেও সমস্যা নেই */ }
}

function ensureSheetsExist() {
  const ss = getSS();

  if (!ss.getSheetByName(SHEET_NAMES.SETTINGS)) {
    const sh = ss.insertSheet(SHEET_NAMES.SETTINGS);
    sh.appendRow(['key', 'value']);
    sh.appendRow(['mainBranchName', '']);
    sh.appendRow(['mainBranchAddress', '']);
    sh.appendRow(['branchPassword', 'branch123']);
  }

  let programsSh = ss.getSheetByName(SHEET_NAMES.PROGRAMS);
  if (!programsSh) {
    programsSh = ss.insertSheet(SHEET_NAMES.PROGRAMS);
    programsSh.appendRow(['id', 'name', 'year', 'startDate', 'endDate', 'extendedDate', 'totalDays', 'createdAt']);
  } else {
    migrateProgramsYearIfNeeded(programsSh);
  }
  // তারিখ কলামগুলো Plain Text ফরম্যাটে রাখা হয় যাতে Google Sheets কখনো এগুলোকে
  // স্বয়ংক্রিয়ভাবে Date অবজেক্টে রূপান্তর না করে (এতে ১ দিন এদিক-ওদিক দেখানোর সমস্যা হতো)।
  // মাইগ্রেশন সরাসরি সেল বদলায়, তাই এতে আসলেই কোনো মান বদলালে ক্যাশ-ভার্সনও বাড়িয়ে দিতে হয় —
  // নাহলে আগে থেকে ক্যাশ হয়ে থাকা (ভুল তারিখের) বুটস্ট্র্যাপ/ড্যাশবোর্ড ৫ মিনিট পর্যন্ত পুরনো
  // অবস্থাতেই থেকে যেত এবং রিফ্রেশ বাটন চাপার আগ পর্যন্ত ভুল তারিখ দেখাত।
  if (migrateDateColumnsToPlainTextIfNeeded(programsSh, ['startDate', 'endDate', 'extendedDate'])) {
    bumpVersion(SHEET_NAMES.PROGRAMS);
  }
  // id কলাম সরাসরি অ্যাপের ভেতরের লিংকিং-এর জন্য দরকার (এডিট/ডিলিট/ফিল্টার এর জন্য), ফরমেও
  // নেই — তাই এটা মুছে ফেলা হয় না, শুধু চোখের আড়ালে (Hide) রাখা হয় যাতে শীট খুললে অগোছালো না লাগে
  hideColumnsByName(programsSh, ['id']);

  let branchesSh = ss.getSheetByName(SHEET_NAMES.BRANCHES);
  if (!branchesSh) {
    branchesSh = ss.insertSheet(SHEET_NAMES.BRANCHES);
    branchesSh.appendRow(['id', 'name', 'totalRukon', 'totalKormi', 'totalUnit', 'createdAt']);
  } else {
    migrateBranchesSheetIfNeeded(branchesSh);
  }
  hideColumnsByName(branchesSh, ['id']);

  let reportsSh = ss.getSheetByName(SHEET_NAMES.REPORTS);
  if (!reportsSh) {
    reportsSh = ss.insertSheet(SHEET_NAMES.REPORTS);
    reportsSh.appendRow(['id', 'date', 'programId', 'branchId', 'branchName', 'rukonParticipated', 'kormiParticipated', 'unitParticipated', 'groupCount', 'dawatCount', 'shohojogiCount', 'createdAt']);
  } else {
    migrateReportsSheetIfNeeded(reportsSh);
    migrateReportsGroupCountIfNeeded(reportsSh);
  }
  if (migrateDateColumnsToPlainTextIfNeeded(reportsSh, ['date'])) {
    bumpVersion(SHEET_NAMES.REPORTS);
  }
  // id/programId/branchId — এই ৩টা কলাম অ্যাপের ভেতরের লিংকিং-এর জন্য দরকার (কোন রিপোর্ট কোন
  // কর্মসূচী/শাখার, এবং এডিট-ডিলিটের জন্য টার্গেট রো খুঁজে বের করা) — ফরমে এগুলো টাইপ করতে হয় না,
  // branchName দেখেই মানুষ শাখা বোঝে, তাই এই টেকনিক্যাল কলামগুলো Hide করে দেওয়া হলো (ডাটা মুছে ফেলা হয়নি,
  // শুধু চোখের আড়ালে)।
  hideColumnsByName(reportsSh, ['id', 'programId', 'branchId']);

  if (!ss.getSheetByName(SHEET_NAMES.USERS)) {
    const sh = ss.insertSheet(SHEET_NAMES.USERS);
    sh.appendRow(['username', 'password', 'role', 'branchId']);
    // ডিফল্ট এডমিন
    sh.appendRow(['admin', 'admin123', 'admin', '']);
  }
}

// হেডার-নাম দিয়ে একটা কলাম খুঁজে সেটা Hide করে দেয় (মুছে ফেলে না, শুধু লুকিয়ে রাখে) —
// একাধিকবার চললেও সমস্যা নেই, আগে থেকে Hide করা কলাম আবার Hide করলে কিছু হয় না
function hideColumnsByName(sh, columnNames) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  columnNames.forEach(colName => {
    const idx = headers.indexOf(colName); // 0-indexed
    if (idx === -1) return;
    try { sh.hideColumns(idx + 1); } catch (e) { /* হাইড করতে ব্যর্থ হলেও অ্যাপ চলবে */ }
  });
}

// পুরনো Branches শীটে (id, name, createdAt) নতুন ৩টি কলাম (totalRukon, totalKormi, totalUnit) যোগ করে দেয়
function migrateBranchesSheetIfNeeded(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('totalRukon') !== -1) return; // আগে থেকেই মাইগ্রেট করা আছে

  const nameIdx = headers.indexOf('name'); // 0-indexed
  if (nameIdx === -1) return;
  const insertAfterCol = nameIdx + 1; // 1-indexed কলাম নম্বর (name এর কলাম)
  sh.insertColumnsAfter(insertAfterCol, 3);
  sh.getRange(1, insertAfterCol + 1, 1, 3).setValues([['totalRukon', 'totalKormi', 'totalUnit']]);

  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const rows = lastRow - 1;
    const defaults = [];
    for (let i = 0; i < rows; i++) defaults.push([0, 0, 0]);
    sh.getRange(2, insertAfterCol + 1, rows, 3).setValues(defaults);
  }
}

// পুরনো Programs শীটে (id, name, startDate, ...) নতুন 'year' কলাম যোগ করে দেয় (name এর পরে),
// বিদ্যমান রো-গুলোর জন্য startDate থেকে বছর অনুমান করে ডিফল্ট মান বসায়
function migrateProgramsYearIfNeeded(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('year') !== -1) return; // আগে থেকেই মাইগ্রেট করা আছে

  const nameIdx = headers.indexOf('name'); // 0-indexed
  const startDateIdx = headers.indexOf('startDate'); // 0-indexed, মাইগ্রেশনের আগের অবস্থান
  if (nameIdx === -1) return;
  const insertAfterCol = nameIdx + 1; // 1-indexed কলাম নম্বর (name এর কলাম)
  sh.insertColumnsAfter(insertAfterCol, 1);
  sh.getRange(1, insertAfterCol + 1, 1, 1).setValue('year');

  const lastRow = sh.getLastRow();
  if (lastRow > 1 && startDateIdx !== -1) {
    const rows = lastRow - 1;
    // startDate কলাম insertColumnsAfter এর কারণে ১ ঘর ডানে সরে গেছে
    const newStartDateCol = startDateIdx + 1 >= insertAfterCol ? startDateIdx + 2 : startDateIdx + 1;
    const startDates = sh.getRange(2, newStartDateCol, rows, 1).getValues();
    const years = startDates.map(row => {
      const d = row[0];
      let y = new Date().getFullYear();
      if (Object.prototype.toString.call(d) === '[object Date]') {
        y = d.getFullYear();
      } else if (typeof d === 'string' && d.length >= 4) {
        const parsedY = parseInt(d.slice(0, 4), 10);
        if (!isNaN(parsedY)) y = parsedY;
      }
      return [y];
    });
    sh.getRange(2, insertAfterCol + 1, rows, 1).setValues(years);
  }
}

// পুরনো Reports শীটের কলাম নাম পরিবর্তন করে (rukonCount→rukonParticipated,
// kormiCount→kormiParticipated, groupCount→unitParticipated) — কলামের অবস্থান/ডাটা অপরিবর্তিত থাকে
function migrateReportsSheetIfNeeded(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('rukonParticipated') !== -1) return; // আগে থেকেই মাইগ্রেট করা আছে

  const renameMap = { rukonCount: 'rukonParticipated', kormiCount: 'kormiParticipated', groupCount: 'unitParticipated' };
  const newHeaders = headers.map(h => renameMap[h] || h);
  sh.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
}

// "মোট দাওয়াত দেয়া হয়েছে" এর আগে নতুন "groupCount" (মোট গ্রুপ বের হয়েছে) কলাম যোগ করে (ডিফল্ট মান ০)
function migrateReportsGroupCountIfNeeded(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('groupCount') !== -1) return; // আগে থেকেই আছে

  const dawatIdx = headers.indexOf('dawatCount'); // 0-indexed
  if (dawatIdx === -1) return;
  const insertBeforeCol = dawatIdx + 1; // 1-indexed কলাম নম্বর (dawatCount এর কলাম)
  sh.insertColumnsBefore(insertBeforeCol, 1);
  sh.getRange(1, insertBeforeCol, 1, 1).setValue('groupCount');

  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const rows = lastRow - 1;
    const defaults = [];
    for (let i = 0; i < rows; i++) defaults.push([0]);
    sh.getRange(2, insertBeforeCol, rows, 1).setValues(defaults);
  }
}

// তারিখ কলামে টাইপ/পেস্ট করা বা আগে সেভ হওয়া মান Google Sheets স্বয়ংক্রিয়ভাবে Date অবজেক্টে
// রূপান্তর করে ফেলতে পারে, এবং সেই Date স্প্রেডশিটের নিজস্ব টাইমজোন অনুযায়ী সংরক্ষিত হয় —
// যা APP_TIMEZONE এর সাথে না মিললে রিপোর্টে তারিখ ১ দিন আগে/পরে দেখাতে পারে
// (যেমন শীটে ১ সেপ্টেম্বর থাকলেও রিপোর্টে ৩১ আগস্ট দেখানো)।
// এই ফাংশন কলামগুলো Plain Text ফরম্যাটে বদলে দেয় (ভবিষ্যতে আর কখনো এই সমস্যা হবে না) এবং
// ইতিমধ্যে Date হয়ে যাওয়া সেলগুলোকে শীটে যেভাবে দেখাচ্ছে ঠিক সেই তারিখ অনুযায়ী প্লেইন টেক্সটে ফিরিয়ে আনে।
// আগে থেকেই প্লেইন টেক্সট থাকা সেলগুলো স্পর্শ করা হয় না, তাই বারবার চললেও সমস্যা নেই।
function migrateDateColumnsToPlainTextIfNeeded(sh, columnNames) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return false;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let anyChanged = false;

  columnNames.forEach(colName => {
    const idx = headers.indexOf(colName); // 0-indexed
    if (idx === -1) return;
    const col = idx + 1; // 1-indexed কলাম নম্বর

    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const rows = lastRow - 1;
      const range = sh.getRange(2, col, rows, 1);
      const values = range.getValues();
      let changed = false;
      const fixed = values.map(row => {
        const v = row[0];
        if (Object.prototype.toString.call(v) === '[object Date]') {
          changed = true;
          return [formatDateStr(v)];
        }
        // ইতিমধ্যে সঠিক yyyy-MM-dd স্ট্রিং না হলে (যেমন হাতে "31-08-2026" আকারে টাইপ করা থাকলে)
        // formatDateStr দিয়ে ঠিক করে সেভ করে রাখা হয়, যাতে সাজানো/তুলনা করার সময় ভুল না হয়
        if (v !== '' && v !== null && v !== undefined && !/^\d{4}-\d{2}-\d{2}$/.test(String(v).trim())) {
          const normalized = formatDateStr(v);
          if (normalized && normalized !== v) {
            changed = true;
            return [normalized];
          }
        }
        return [v];
      });
      range.setNumberFormat('@'); // ফরম্যাট আগে বদলাতে হয়, নাহলে setValues আবার Date বানিয়ে ফেলবে
      if (changed) {
        range.setValues(fixed);
        anyChanged = true;
      }
    }

    // নিচের ফাঁকা রো-গুলোও প্লেইন টেক্সট রাখা হলো, যাতে ভবিষ্যতে নতুন এন্ট্রিও Date-এ রূপান্তরিত না হয়
    const maxRows = sh.getMaxRows();
    if (maxRows > lastRow) {
      sh.getRange(lastRow + 1, col, maxRows - lastRow, 1).setNumberFormat('@');
    }
  });

  return anyChanged;
}

// ==================== হেল্পার ====================

function sheetToObjects(sheetName) {
  const sh = getSS().getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(c => c === '')) continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });
    rows.push(obj);
  }
  return rows;
}

function generateId() {
  return Utilities.getUuid();
}

function findRowIndexById(sheetName, id) {
  const sh = getSS().getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return i + 1; // 1-indexed row number
  }
  return -1;
}

// ==================== DATE / TIMEZONE ====================

const APP_TIMEZONE = 'Asia/Dhaka'; // শুধু "আজ" হিসাব করতে ব্যবহৃত হয়, তারিখ-কলাম পড়ার জন্য নয়

// স্প্রেডশিটের নিজস্ব টাইমজোন (File > Settings এ যা সেট করা আছে)। কোনো সেল যদি (টাইপ/পেস্ট
// করার কারণে) Date অবজেক্ট হয়ে যায়, Google Sheets সবসময় সেই মান এই টাইমজোন অনুযায়ীই দেখায় —
// তাই সেই একই টাইমজোন দিয়ে ফেরত ফরম্যাট করলে শীটে যা দেখা যাচ্ছে ঠিক সেই তারিখই পাওয়া যাবে
// (আগে হার্ডকোড করা APP_TIMEZONE ব্যবহার হতো, যা স্প্রেডশিটের প্রকৃত টাইমজোনের সাথে না মিললে
// তারিখ ১ দিন আগে/পরে দেখানোর সমস্যা করত)।
let _sheetTzCache = null;
function getSheetTimeZone() {
  if (!_sheetTzCache) {
    try {
      _sheetTzCache = getSS().getSpreadsheetTimeZone();
    } catch (e) {
      _sheetTzCache = APP_TIMEZONE;
    }
  }
  return _sheetTzCache;
}

/**
 * সব তারিখ yyyy-MM-dd আকারে রাখে।
 * শীটে যেভাবে দেখা যাচ্ছে ঠিক সেভাবেই ফেরত দেয় (স্প্রেডশিটের নিজস্ব টাইমজোন অনুযায়ী)।
 *
 * গুরুত্বপূর্ণ:
 * yyyy-MM-dd string কখনো Date হিসেবে parse করা হবে না।
 * এতে ১ সেপ্টেম্বর → ৩১ আগস্ট হওয়ার সমস্যা হবে না।
 */
function formatDateStr(d) {
  if (d === null || d === undefined || d === '') {
    return '';
  }

  // Google Sheet থেকে Date object এলে
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, getSheetTimeZone(), 'yyyy-MM-dd');
  }

  const s = String(d).trim();

  // আগে থেকেই সঠিক ISO date হলে যেমন আছে তেমনই ফেরত দিন
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return s;
  }

  // yyyy/MM/dd বা yyyy-MM-dd HH:mm:ss ধরনের string
  const m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);

  if (m) {
    return [
      m[1],
      String(m[2]).padStart(2, '0'),
      String(m[3]).padStart(2, '0')
    ].join('-');
  }

  // dd-MM-yyyy বা dd/MM/yyyy ধরনের string (বাংলাদেশে প্রায়ই এভাবে হাতে টাইপ করা হয়,
  // যেমন "31-08-2026")। এটা আগে ধরা না পড়লে "31" সংখ্যাটা তারিখ হিসেবে বোঝা যেত না,
  // ফলে সিরিয়াল/সাজানোর সময় এলোমেলো জায়গায় (যেমন ১ তারিখের আগে) চলে যেত।
  const m2 = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (m2) {
    return [
      m2[3],
      String(m2[2]).padStart(2, '0'),
      String(m2[1]).padStart(2, '0')
    ].join('-');
  }

  return s;
}

// সব জায়গায় কর্মসূচীর তারিখ সঠিকভাবে ফরম্যাট করা অবস্থায় পাওয়ার জন্য শেয়ার্ড হেল্পার
function getFormattedPrograms() {
  return getCachedSheetObjects(SHEET_NAMES.PROGRAMS).map(p => ({
    ...p,
    year: Number(p.year) || new Date().getFullYear(),
    startDate: formatDateStr(p.startDate),
    endDate: formatDateStr(p.endDate),
    extendedDate: p.extendedDate ? formatDateStr(p.extendedDate) : ''
  }));
}

// ==================== ক্যাশ (দ্রুত লোডের জন্য) ====================
// প্রতিটি শীটের জন্য আলাদা ভার্সন নম্বর রাখা হয় (আগে সব শীটের জন্য একটাই ভার্সন ছিল —
// তার মানে একটা রিপোর্ট সেভ করলেই Settings/Programs/Branches এর মতো অপরিবর্তিত ডাটার
// ক্যাশও অকারণে বাতিল হয়ে যেত, ফলে প্রতিটি রিপোর্ট এন্ট্রির পর পুরো অ্যাপকেই আবার শীট থেকে
// সব পড়তে হতো — এটাই ছিল ধীরগতির প্রধান কারণ)। এখন শুধু যেই শীট পরিবর্তন হয়েছে তারই
// ভার্সন বাড়ে, বাকি সব ক্যাশ অক্ষত থাকে — তাই ঘন ঘন রিপোর্ট এন্ট্রি করলেও বুটস্ট্র্যাপ/
// প্রোগ্রাম/শাখার মতো কম পরিবর্তনশীল ডাটা প্রায় সবসময় ক্যাশ থেকেই সাথে সাথে চলে আসে।
const CACHE_TTL_SECONDS = 300;

function getVersion(sheetName) {
  const cache = CacheService.getScriptCache();
  const key = 'v_' + sheetName;
  let v = cache.get(key);
  if (!v) {
    v = '1';
    cache.put(key, v, 21600);
  }
  return v;
}

function bumpVersion(sheetName) {
  try {
    const cache = CacheService.getScriptCache();
    const key = 'v_' + sheetName;
    const v = Number(getVersion(sheetName)) + 1;
    cache.put(key, String(v), 21600);
  } catch (e) { /* ক্যাশ ব্যর্থ হলেও অ্যাপ যেন থেমে না যায় */ }
}

function bumpAllVersions() {
  Object.keys(SHEET_NAMES).forEach(k => bumpVersion(SHEET_NAMES[k]));
}

function cacheGetJSON(key) {
  try {
    const cache = CacheService.getScriptCache();
    const val = cache.get(key);
    return val ? JSON.parse(val) : null;
  } catch (e) {
    return null;
  }
}

function cacheSetJSON(key, value) {
  try {
    const cache = CacheService.getScriptCache();
    cache.put(key, JSON.stringify(value), CACHE_TTL_SECONDS);
  } catch (e) { /* বড় ডাটার কারণে ক্যাশ ব্যর্থ হলেও সমস্যা নেই, শুধু ক্যাশ ছাড়া কাজ করবে */ }
}

// deps: এই ক্যাশ-এন্ট্রি কোন কোন শীটের উপর নির্ভরশীল তার তালিকা (SHEET_NAMES.* থেকে) —
// শুধু সেই শীটগুলোর ভার্সন দিয়েই কী তৈরি হয়, তাই অন্য কোনো শীট পরিবর্তন হলে এই ক্যাশ নষ্ট হয় না
function versionedKey(prefix, deps, suffix) {
  const vs = deps.map(getVersion).join('.');
  return prefix + '_v' + vs + '_' + (suffix || 'x');
}

// যেকোনো "শুধু পড়ার" ফাংশনকে ক্যাশ দিয়ে মুড়িয়ে দেওয়ার শর্টকাট —
// একই ভার্সনে একই ইনপুটের জন্য দ্বিতীয়বার শীট না পড়ে সরাসরি ক্যাশ থেকে উত্তর দেয়
function withCache(prefix, deps, keySuffix, computeFn) {
  const cacheKey = versionedKey(prefix, deps, keySuffix);
  const cached = cacheGetJSON(cacheKey);
  if (cached !== null) {
    return { success: true, data: cached };
  }
  const data = computeFn();
  cacheSetJSON(cacheKey, data);
  return { success: true, data };
}

// কাঁচা শীট-ডাটা (sheetToObjects এর ফলাফল) নিজেও ক্যাশ করা থাকে — এতে করে কম্পিউটেড ক্যাশ
// (যেমন কোনো নির্দিষ্ট কর্মসূচীর ড্যাশবোর্ড) মিস হলেও, সেই একই শীট আগের কোনো রিকোয়েস্টে
// পড়া হয়ে থাকলে আবার শীট থেকে না পড়ে ক্যাশ থেকেই কাজ চালানো যায় — এটা সবচেয়ে বড় স্পিড বুস্ট।
// শুধু ওই একটা শীটের ভার্সনের উপর নির্ভর করে, তাই অন্য শীট পরিবর্তন হলে এটা এখনো ক্যাশ থেকেই আসবে।
function getCachedSheetObjects(sheetName) {
  const cacheKey = versionedKey('rawsheet', [sheetName], sheetName);
  const cached = cacheGetJSON(cacheKey);
  if (cached !== null) return cached;
  const data = sheetToObjects(sheetName);
  cacheSetJSON(cacheKey, data);
  return data;
}

// রিফ্রেশ বাটনের জন্য — জোর করে সব শীটের ভার্সন বাড়িয়ে সব ক্যাশ সাথে সাথে অকার্যকর করে দেয়,
// ফলে পরের রিকোয়েস্টেই একদম তাজা ডাটা শীট থেকে আনা হবে
function forceRefreshData() {
  bumpAllVersions();
  return { success: true, message: 'রিফ্রেশ হয়েছে' };
}

// ==================== অথেনটিকেশন ====================

function login(params) {
  const loginType = String(params.loginType || '').trim();

  if (loginType === 'admin') {
    const password = String(params.password || '').trim();
    const users = getCachedSheetObjects(SHEET_NAMES.USERS);
    const user = users.find(u => u.role === 'admin' && String(u.password).trim() === password);
    if (!user) {
      return { success: false, message: 'পাসওয়ার্ড ভুল' };
    }
    return { success: true, data: { username: user.username, role: 'admin' } };
  }

  if (loginType === 'branch') {
    const password = String(params.password || '').trim();
    const settings = getSettings().data;
    const branchPassword = String(settings.branchPassword || '').trim();
    if (password !== branchPassword) {
      return { success: false, message: 'পাসওয়ার্ড ভুল' };
    }
    return {
      success: true,
      data: { role: 'branch' }
    };
  }

  return { success: false, message: 'অবৈধ লগইন অনুরোধ' };
}

// ==================== মূল শাখা সেটাপ ====================

function getRawSettings() {
  const sh = getSS().getSheetByName(SHEET_NAMES.SETTINGS);
  const data = sh.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  if (!settings.branchPassword) settings.branchPassword = 'branch123';
  return settings;
}

function getSettings() {
  return { success: true, data: getRawSettings() };
}

function saveSettings(params) {
  const sh = getSS().getSheetByName(SHEET_NAMES.SETTINGS);
  const data = sh.getDataRange().getValues();
  const updates = {};
  if (params.mainBranchName !== undefined) updates.mainBranchName = params.mainBranchName || '';
  if (params.mainBranchAddress !== undefined) updates.mainBranchAddress = params.mainBranchAddress || '';
  if (params.branchPassword !== undefined) updates.branchPassword = params.branchPassword || '';

  const foundKeys = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    if (updates.hasOwnProperty(key)) {
      sh.getRange(i + 1, 2).setValue(updates[key]);
      foundKeys[key] = true;
    }
  }
  // পুরনো শীটে যদি কোনো key না থাকে (যেমন branchPassword), তাহলে নতুন রো যোগ করে দাও
  Object.keys(updates).forEach(key => {
    if (!foundKeys[key]) {
      sh.appendRow([key, updates[key]]);
    }
  });
  bumpVersion(SHEET_NAMES.SETTINGS);
  return { success: true, message: 'সেটিংস সংরক্ষণ হয়েছে' };
}

// লগইন পেইজে দেখানোর জন্য শুধু মূল শাখার নাম — লগইনের আগেই প্রয়োজন, তাই আলাদা হালকা এন্ডপয়েন্ট
function getLoginInfo() {
  const cacheKey = versionedKey('loginInfo', [SHEET_NAMES.SETTINGS], 'x');
  const cached = cacheGetJSON(cacheKey);
  if (cached) {
    return { success: true, data: cached };
  }
  const settings = getRawSettings();
  const result = { mainBranchName: settings.mainBranchName || '' };
  cacheSetJSON(cacheKey, result);
  return { success: true, data: result };
}

// লগইনের পর একবারে সেটিংস + কর্মসূচী + শাখা — সব একসাথে পাঠানো হয়,
// যাতে ৩টি আলাদা রিকোয়েস্টের বদলে মাত্র ১টি রিকোয়েস্টেই অ্যাপ শুরু হতে পারে (দ্রুত লোড)
function getBootstrap() {
  const cacheKey = versionedKey('bootstrap', [SHEET_NAMES.SETTINGS, SHEET_NAMES.PROGRAMS, SHEET_NAMES.BRANCHES], 'all');
  const cached = cacheGetJSON(cacheKey);
  if (cached) {
    return { success: true, data: cached };
  }

  const settings = getRawSettings();
  const programs = getFormattedPrograms();
  const branches = getCachedSheetObjects(SHEET_NAMES.BRANCHES)
    .map(b => ({
      id: b.id,
      name: b.name,
      totalRukon: Number(b.totalRukon) || 0,
      totalKormi: Number(b.totalKormi) || 0,
      totalUnit: Number(b.totalUnit) || 0
    }))

  const result = { settings, programs, branches };
  cacheSetJSON(cacheKey, result);
  return { success: true, data: result };
}

// ==================== কর্মসূচী ====================

function calcTotalDays(startDate, endDate, extendedDate) {
  const start = new Date(startDate);
  let end = extendedDate ? new Date(extendedDate) : new Date(endDate);
  const diff = Math.round((end - start) / (1000 * 60 * 60 * 24)) + 1;
  return diff > 0 ? diff : 0;
}

function getPrograms() {
  return withCache('programs', [SHEET_NAMES.PROGRAMS], 'x', () => getFormattedPrograms());
}

function saveProgram(params) {
  const sh = getSS().getSheetByName(SHEET_NAMES.PROGRAMS);
  const totalDays = calcTotalDays(params.startDate, params.endDate, params.extendedDate);
  const year = Number(params.year) || new Date().getFullYear();

  if (params.id) {
    const rowIdx = findRowIndexById(SHEET_NAMES.PROGRAMS, params.id);
    if (rowIdx === -1) return { success: false, message: 'কর্মসূচী পাওয়া যায়নি' };
    sh.getRange(rowIdx, 2, 1, 6).setValues([[
      params.name, year, params.startDate, params.endDate, params.extendedDate || '', totalDays
    ]]);
    bumpVersion(SHEET_NAMES.PROGRAMS);
    return { success: true, message: 'কর্মসূচী আপডেট হয়েছে' };
  } else {
    const id = generateId();
    sh.appendRow([id, params.name, year, params.startDate, params.endDate, params.extendedDate || '', totalDays, new Date()]);
    bumpVersion(SHEET_NAMES.PROGRAMS);
    return { success: true, message: 'কর্মসূচী যোগ হয়েছে', data: { id } };
  }
}

function deleteProgram(params) {
  const rowIdx = findRowIndexById(SHEET_NAMES.PROGRAMS, params.id);
  if (rowIdx === -1) return { success: false, message: 'কর্মসূচী পাওয়া যায়নি' };
  getSS().getSheetByName(SHEET_NAMES.PROGRAMS).deleteRow(rowIdx);
  bumpVersion(SHEET_NAMES.PROGRAMS);
  return { success: true, message: 'কর্মসূচী মুছে ফেলা হয়েছে' };
}

// ==================== শাখা ====================

function getBranches() {
  return withCache('branches', [SHEET_NAMES.BRANCHES], 'x', () => {
    return getCachedSheetObjects(SHEET_NAMES.BRANCHES)
      .map(b => ({
        id: b.id,
        name: b.name,
        totalRukon: Number(b.totalRukon) || 0,
        totalKormi: Number(b.totalKormi) || 0,
        totalUnit: Number(b.totalUnit) || 0
      }));
  });
}

function saveBranch(params) {
  const sh = getSS().getSheetByName(SHEET_NAMES.BRANCHES);
  const totalRukon = Number(params.totalRukon) || 0;
  const totalKormi = Number(params.totalKormi) || 0;
  const totalUnit = Number(params.totalUnit) || 0;

  if (params.id) {
    const rowIdx = findRowIndexById(SHEET_NAMES.BRANCHES, params.id);
    if (rowIdx === -1) return { success: false, message: 'শাখা পাওয়া যায়নি' };
    sh.getRange(rowIdx, 2, 1, 4).setValues([[params.name, totalRukon, totalKormi, totalUnit]]);
    bumpVersion(SHEET_NAMES.BRANCHES);
    return { success: true, message: 'শাখা আপডেট হয়েছে' };
  } else {
    const id = generateId();
    sh.appendRow([id, params.name, totalRukon, totalKormi, totalUnit, new Date()]);
    bumpVersion(SHEET_NAMES.BRANCHES);
    return { success: true, message: 'শাখা যোগ হয়েছে', data: { id } };
  }
}

function deleteBranch(params) {
  const rowIdx = findRowIndexById(SHEET_NAMES.BRANCHES, params.id);
  if (rowIdx === -1) return { success: false, message: 'শাখা পাওয়া যায়নি' };
  getSS().getSheetByName(SHEET_NAMES.BRANCHES).deleteRow(rowIdx);
  bumpVersion(SHEET_NAMES.BRANCHES);
  return { success: true, message: 'শাখা মুছে ফেলা হয়েছে' };
}

// ==================== রিপোর্ট ====================

function getReports(params) {
  const keySuffix = [
    params.branchId || '', params.programId || '', params.date || '',
    params.fromDate || '', params.toDate || ''
  ].join('|');

  return withCache('reports', [SHEET_NAMES.REPORTS], keySuffix, () => {
    let reports = getCachedSheetObjects(SHEET_NAMES.REPORTS).map(r => ({
      ...r,
      date: formatDateStr(r.date)
    }));

    if (params.branchId) {
      reports = reports.filter(r => String(r.branchId) === String(params.branchId));
    }
    if (params.programId) {
      reports = reports.filter(r => String(r.programId) === String(params.programId));
    }
    if (params.date) {
      reports = reports.filter(r => r.date === params.date);
    }
    if (params.fromDate) {
      reports = reports.filter(r => r.date >= params.fromDate);
    }
    if (params.toDate) {
      reports = reports.filter(r => r.date <= params.toDate);
    }

    // শীটে রো যে ক্রমেই থাকুক (এলোমেলো হলেও), ওয়েবসাইটে সবসময় তারিখ অনুযায়ী সঠিক
    // ক্রমেই (১,২,৩...) দেখানো হয়। একই তারিখের একাধিক এন্ট্রি থাকলে সেগুলো কোন আগে
    // সাবমিট হয়েছে (createdAt) সেই ক্রমে দেখানো হয়।
    reports.sort((a, b) => {
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      return String(a.createdAt).localeCompare(String(b.createdAt));
    });

    return reports;
  });
}

function saveReport(params) {
  const sh = getSS().getSheetByName(SHEET_NAMES.REPORTS);

  const branches = getCachedSheetObjects(SHEET_NAMES.BRANCHES);
  const branchIds = JSON.parse(params.branchIds || '[]');
  const existingReports = getCachedSheetObjects(SHEET_NAMES.REPORTS);

  const rukonParticipated = Number(params.rukonParticipated) || 0;
  const kormiParticipated = Number(params.kormiParticipated) || 0;
  const unitParticipated = Number(params.unitParticipated) || 0;
  const groupCount = Number(params.groupCount) || 0;
  const dawatCount = Number(params.dawatCount) || 0;
  const shohojogiCount = Number(params.shohojogiCount) || 0;

  // ===== এডিট মোড: params.id দেওয়া থাকলে নতুন রো না বানিয়ে পুরনো রিপোর্টটাই আপডেট করা হয় =====
  if (params.id) {
    const rowIdx = findRowIndexById(SHEET_NAMES.REPORTS, params.id);
    if (rowIdx === -1) return { success: false, message: 'রিপোর্ট পাওয়া যায়নি' };

    const branchId = branchIds[0];
    const branch = branches.find(b => String(b.id) === String(branchId));

    const dup = existingReports.find(r =>
      String(r.id) !== String(params.id) &&
      String(r.programId) === String(params.programId) &&
      String(r.branchId) === String(branchId) &&
      formatDateStr(r.date) === params.date
    );
    if (dup) {
      return {
        success: false,
        message: (branch ? branch.name : 'এই শাখার') + ' জন্য এই তারিখে ইতিমধ্যে আরেকটি রিপোর্ট আছে — একই শাখার একই দিনে দুটি রিপোর্ট রাখা যাবে না'
      };
    }

    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const rowRange = sh.getRange(rowIdx, 1, 1, lastCol);
    const rowValues = rowRange.getValues()[0]; // id, createdAt ইত্যাদি অপরিবর্তিত রাখতে আগে থেকে লোড করা হলো

    const updates = {
      date: params.date,
      programId: params.programId,
      branchId: branchId,
      branchName: branch ? branch.name : '',
      rukonParticipated, kormiParticipated, unitParticipated,
      groupCount, dawatCount, shohojogiCount
    };
    headers.forEach((h, i) => {
      if (Object.prototype.hasOwnProperty.call(updates, h)) rowValues[i] = updates[h];
    });
    rowRange.setValues([rowValues]);

    bumpVersion(SHEET_NAMES.REPORTS);
    return { success: true, message: 'রিপোর্ট আপডেট হয়েছে', data: { ids: [params.id] } };
  }

  // ===== নতুন এন্ট্রি (একই সাথে একাধিক শাখার জন্যও হতে পারে) =====
  // একই তারিখ + একই শাখা + একই কর্মসূচীতে আগে থেকে রিপোর্ট থাকলে ডাবল এন্ট্রি আটকানো হয়
  const dupBranchId = branchIds.find(branchId =>
    existingReports.some(r =>
      String(r.programId) === String(params.programId) &&
      String(r.branchId) === String(branchId) &&
      formatDateStr(r.date) === params.date
    )
  );
  if (dupBranchId) {
    const branch = branches.find(b => String(b.id) === String(dupBranchId));
    return {
      success: false,
      message: (branch ? branch.name : 'এই শাখার') + ' জন্য এই তারিখে রিপোর্ট আগে থেকেই জমা আছে — নতুন করে যোগ না করে সেটা এডিট করুন'
    };
  }

  const results = [];
  branchIds.forEach(branchId => {
    const branch = branches.find(b => String(b.id) === String(branchId));
    const id = generateId();
    sh.appendRow([
      id, params.date, params.programId, branchId,
      branch ? branch.name : '',
      rukonParticipated,
      kormiParticipated,
      unitParticipated,
      groupCount,
      dawatCount,
      shohojogiCount,
      new Date()
    ]);
    results.push(id);
  });

  bumpVersion(SHEET_NAMES.REPORTS);
  return { success: true, message: 'রিপোর্ট সংরক্ষণ হয়েছে', data: { ids: results } };
}

function deleteReport(params) {
  const rowIdx = findRowIndexById(SHEET_NAMES.REPORTS, params.id);
  if (rowIdx === -1) return { success: false, message: 'রিপোর্ট পাওয়া যায়নি' };
  getSS().getSheetByName(SHEET_NAMES.REPORTS).deleteRow(rowIdx);
  bumpVersion(SHEET_NAMES.REPORTS);
  return { success: true, message: 'রিপোর্ট মুছে ফেলা হয়েছে' };
}

// ==================== ড্যাশবোর্ড ====================

function getDashboard(params) {
  const programId = params.programId || '';
  const cacheKey = versionedKey('dash', [SHEET_NAMES.REPORTS, SHEET_NAMES.PROGRAMS, SHEET_NAMES.BRANCHES, SHEET_NAMES.SETTINGS], programId || 'latest');
  const cached = cacheGetJSON(cacheKey);
  if (cached) {
    return { success: true, data: cached };
  }

  const settings = getRawSettings();
  const programs = getFormattedPrograms();
  const branches = getCachedSheetObjects(SHEET_NAMES.BRANCHES);
  let reports = getCachedSheetObjects(SHEET_NAMES.REPORTS).map(r => ({ ...r, date: formatDateStr(r.date) }));

  let activeProgram = null;
  if (programId) {
    activeProgram = programs.find(p => String(p.id) === String(programId));
    reports = reports.filter(r => String(r.programId) === String(programId));
  } else if (programs.length > 0) {
    activeProgram = programs[programs.length - 1];
    reports = reports.filter(r => String(r.programId) === String(activeProgram.id));
  }

  const today = Utilities.formatDate(
  new Date(),
  APP_TIMEZONE,
  'yyyy-MM-dd'
);
  const todayReports = reports.filter(r => r.date === today);

  const reportedBranchIds = [...new Set(reports.map(r => String(r.branchId)))];
  const totalBranches = branches.length;
  const reportedCount = reportedBranchIds.length;
  const notReportedCount = Math.max(0, totalBranches - reportedCount);

  function sumField(arr, field) {
    return arr.reduce((sum, r) => sum + (Number(r[field]) || 0), 0);
  }

  const overall = {
    reportedBranches: reportedCount,
    notReportedBranches: notReportedCount,
    totalUnit: sumField(reports, 'unitParticipated'),
    totalGroup: sumField(reports, 'groupCount'),
    totalDawat: sumField(reports, 'dawatCount'),
    totalRukon: sumField(reports, 'rukonParticipated'),
    totalKormi: sumField(reports, 'kormiParticipated'),
    totalShohojogi: sumField(reports, 'shohojogiCount')
  };

  const todayBranchIds = [...new Set(todayReports.map(r => String(r.branchId)))];
  const todaySummary = {
    branchCount: todayBranchIds.length,
    branchNotReportedCount: Math.max(0, totalBranches - todayBranchIds.length),
    totalUnit: sumField(todayReports, 'unitParticipated'),
    totalGroup: sumField(todayReports, 'groupCount'),
    totalDawat: sumField(todayReports, 'dawatCount'),
    totalRukon: sumField(todayReports, 'rukonParticipated'),
    totalKormi: sumField(todayReports, 'kormiParticipated'),
    totalShohojogi: sumField(todayReports, 'shohojogiCount')
  };

  const branchWise = branches.map(b => {
    const branchReports = reports.filter(r => String(r.branchId) === String(b.id));
    return {
      branchId: b.id,
      branchName: b.name,
      profileRukon: Number(b.totalRukon) || 0,
      profileKormi: Number(b.totalKormi) || 0,
      profileUnit: Number(b.totalUnit) || 0,
      unitParticipated: sumField(branchReports, 'unitParticipated'),
      groupCount: sumField(branchReports, 'groupCount'),
      rukonParticipated: sumField(branchReports, 'rukonParticipated'),
      kormiParticipated: sumField(branchReports, 'kormiParticipated'),
      dawatCount: sumField(branchReports, 'dawatCount'),
      shohojogiCount: sumField(branchReports, 'shohojogiCount')
    };
  });

  const result = { settings, activeProgram, overall, today: todaySummary, branchWise };
  cacheSetJSON(cacheKey, result);
  return { success: true, data: result };
}

function getBranchReportSummary(params) {
  const branchId = params.branchId;
  const programId = params.programId;

  return withCache('branchSummary', [SHEET_NAMES.REPORTS, SHEET_NAMES.PROGRAMS, SHEET_NAMES.BRANCHES], branchId + '|' + programId, () => {
    let reports = getCachedSheetObjects(SHEET_NAMES.REPORTS).map(r => ({ ...r, date: formatDateStr(r.date) }));
    reports = reports.filter(r => String(r.branchId) === String(branchId) && String(r.programId) === String(programId));

    const programs = getFormattedPrograms();
    const program = programs.find(p => String(p.id) === String(programId));
    const branches = getCachedSheetObjects(SHEET_NAMES.BRANCHES);
    const branch = branches.find(b => String(b.id) === String(branchId));

    function sumField(arr, field) {
      return arr.reduce((sum, r) => sum + (Number(r[field]) || 0), 0);
    }

    reports.sort((a, b) => a.date.localeCompare(b.date));

    const daily = reports.map(r => ({
      date: r.date,
      unitParticipated: r.unitParticipated,
      groupCount: r.groupCount,
      dawatCount: r.dawatCount,
      rukonParticipated: r.rukonParticipated,
      kormiParticipated: r.kormiParticipated,
      shohojogiCount: r.shohojogiCount
    }));

    const total = {
      totalDays: [...new Set(reports.map(r => r.date))].length,
      totalUnit: sumField(reports, 'unitParticipated'),
      totalGroup: sumField(reports, 'groupCount'),
      totalDawat: sumField(reports, 'dawatCount'),
      totalRukon: sumField(reports, 'rukonParticipated'),
      totalKormi: sumField(reports, 'kormiParticipated'),
      totalShohojogi: sumField(reports, 'shohojogiCount')
    };

    return { branch, program, daily, total };
  });
}

function getMissingBranches(params) {
  const programId = params.programId;
  const date = params.date;

  return withCache('missingBranches', [SHEET_NAMES.REPORTS, SHEET_NAMES.BRANCHES], programId + '|' + (date || ''), () => {
    const branches = getCachedSheetObjects(SHEET_NAMES.BRANCHES);
    let reports = getCachedSheetObjects(SHEET_NAMES.REPORTS).map(r => ({ ...r, date: formatDateStr(r.date) }));
    reports = reports.filter(r => String(r.programId) === String(programId));

    if (date) {
      reports = reports.filter(r => String(r.date).trim() === String(date).trim());
    }

    const reportedIds = new Set(reports.map(r => String(r.branchId)));
    const missing = branches.filter(b => !reportedIds.has(String(b.id)));
    const reported = branches.filter(b => reportedIds.has(String(b.id)));

    return { missing, reported };
  });
}

// ==================== প্রিন্ট রিপোর্ট (অফিস ফরম্যাট) ====================

function sumFieldGeneric(arr, field) {
  return arr.reduce((sum, r) => sum + (Number(r[field]) || 0), 0);
}

// শাখা ভিত্তিক — দৈনিক (নির্দিষ্ট শাখার নির্দিষ্ট তারিখের রিপোর্ট)
function getBranchDateReport(params) {
  const branchId = params.branchId;
  const date = String(params.date || '').trim();
  const programId = params.programId;

  return withCache('branchDate', [SHEET_NAMES.REPORTS, SHEET_NAMES.BRANCHES, SHEET_NAMES.PROGRAMS], [branchId, date, programId].join('|'), () => {
    const branchesRaw = getCachedSheetObjects(SHEET_NAMES.BRANCHES);
    const branchRaw = branchesRaw.find(b => String(b.id) === String(branchId));

    const branch = branchRaw ? {
      id: branchRaw.id,
      name: branchRaw.name,
      totalRukon: Number(branchRaw.totalRukon) || 0,
      totalKormi: Number(branchRaw.totalKormi) || 0,
      totalUnit: Number(branchRaw.totalUnit) || 0
    } : null;

    const programs = getFormattedPrograms();
    const program = programs.find(p => String(p.id) === String(programId));

    let reports = getCachedSheetObjects(SHEET_NAMES.REPORTS)
      .map(r => ({
        ...r,
        date: formatDateStr(r.date)
      }));

reports = reports.filter(r =>
  String(r.branchId) === String(branchId) &&
  String(r.date).trim() === date
);

    if (programId) {
      reports = reports.filter(r =>
        String(r.programId) === String(programId)
      );
    }

    const totals = {
      unitParticipated: sumFieldGeneric(reports, 'unitParticipated'),
      groupCount: sumFieldGeneric(reports, 'groupCount'),
      dawatCount: sumFieldGeneric(reports, 'dawatCount'),
      rukonParticipated: sumFieldGeneric(reports, 'rukonParticipated'),
      kormiParticipated: sumFieldGeneric(reports, 'kormiParticipated'),
      shohojogiCount: sumFieldGeneric(reports, 'shohojogiCount')
    };

    return {
      branch,
      program,
      date: date,
      totals,
      hasReport: reports.length > 0
    };
  });
}

// মূল শাখা — দৈনিক (নির্দিষ্ট তারিখের সকল শাখার রিপোর্ট)
function getMainBranchDateReport(params) {
  const date = String(params.date || '').trim();
  const programId = params.programId;

  return withCache('mainDate', [SHEET_NAMES.REPORTS, SHEET_NAMES.BRANCHES, SHEET_NAMES.PROGRAMS, SHEET_NAMES.SETTINGS], [date, programId].join('|'), () => {
    const settings = getRawSettings();
    const programs = getFormattedPrograms();
    const program = programs.find(p => String(p.id) === String(programId));
    const branches = getCachedSheetObjects(SHEET_NAMES.BRANCHES);

    let reports = getCachedSheetObjects(SHEET_NAMES.REPORTS).map(r => ({ ...r, date: formatDateStr(r.date) }));
    reports = reports.filter(r => String(r.date).trim() === date);
    if (programId) {
      reports = reports.filter(r => String(r.programId) === String(programId));
    }

    const branchRows = branches.map(b => {
      const bReports = reports.filter(r => String(r.branchId) === String(b.id));
      return {
        branchId: b.id,
        branchName: b.name,
        unitParticipated: sumFieldGeneric(bReports, 'unitParticipated'),
        groupCount: sumFieldGeneric(bReports, 'groupCount'),
        dawatCount: sumFieldGeneric(bReports, 'dawatCount'),
        rukonParticipated: sumFieldGeneric(bReports, 'rukonParticipated'),
        kormiParticipated: sumFieldGeneric(bReports, 'kormiParticipated'),
        shohojogiCount: sumFieldGeneric(bReports, 'shohojogiCount')
      };
    });

    const total = {
      unitParticipated: sumFieldGeneric(reports, 'unitParticipated'),
      groupCount: sumFieldGeneric(reports, 'groupCount'),
      dawatCount: sumFieldGeneric(reports, 'dawatCount'),
      rukonParticipated: sumFieldGeneric(reports, 'rukonParticipated'),
      kormiParticipated: sumFieldGeneric(reports, 'kormiParticipated'),
      shohojogiCount: sumFieldGeneric(reports, 'shohojogiCount')
    };

    return { settings, program, date, branches: branchRows, total };
  });
}

// মূল শাখা — পক্ষ (পুরো কর্মসূচীর সময়কালে সকল শাখার সমষ্টি)
function getMainBranchPeriodReport(params) {
  const programId = params.programId;

  return withCache('mainPeriod', [SHEET_NAMES.REPORTS, SHEET_NAMES.BRANCHES, SHEET_NAMES.PROGRAMS, SHEET_NAMES.SETTINGS], programId, () => {
    const settings = getRawSettings();
    const programs = getFormattedPrograms();
    const program = programs.find(p => String(p.id) === String(programId));
    const branches = getCachedSheetObjects(SHEET_NAMES.BRANCHES);

    let reports = getCachedSheetObjects(SHEET_NAMES.REPORTS).map(r => ({ ...r, date: formatDateStr(r.date) }));
    reports = reports.filter(r => String(r.programId) === String(programId));

    const branchRows = branches.map(b => {
      const bReports = reports.filter(r => String(r.branchId) === String(b.id));
      return {
        branchId: b.id,
        branchName: b.name,
        unitParticipated: sumFieldGeneric(bReports, 'unitParticipated'),
        groupCount: sumFieldGeneric(bReports, 'groupCount'),
        dawatCount: sumFieldGeneric(bReports, 'dawatCount'),
        rukonParticipated: sumFieldGeneric(bReports, 'rukonParticipated'),
        kormiParticipated: sumFieldGeneric(bReports, 'kormiParticipated'),
        shohojogiCount: sumFieldGeneric(bReports, 'shohojogiCount')
      };
    });

    const total = {
      unitParticipated: sumFieldGeneric(reports, 'unitParticipated'),
      groupCount: sumFieldGeneric(reports, 'groupCount'),
      dawatCount: sumFieldGeneric(reports, 'dawatCount'),
      rukonParticipated: sumFieldGeneric(reports, 'rukonParticipated'),
      kormiParticipated: sumFieldGeneric(reports, 'kormiParticipated'),
      shohojogiCount: sumFieldGeneric(reports, 'shohojogiCount')
    };

    return { settings, program, branches: branchRows, total };
  });
}
