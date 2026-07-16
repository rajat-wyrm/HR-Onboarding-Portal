var MASTER_SHEET = "Intern Data";
var PROJECT_SHEET = "Project Details";
var PROGRAM_SHEET = "Program Details";
var DRIVE_FOLDER_ID = "1Gj-y3-im1z6r8EChMhOG3uLffxRYkfJY";
var HR_USERID = "admin";
var HR_PASSKEY = "admin";
var API_TOKEN = "hrtn_2026_x";

var CONFIG = {
  ORGANIZATION_NAME: "UptoSkills",
  HR_EMAIL: "hr@uptoskills.com",
  ADDRESS: "UptoSkills Corporate Office, India",
  WEBSITE: "https://www.uptoskills.com",
  LINKEDIN: "https://www.linkedin.com/company/uptoskills",
  INSTAGRAM: "https://www.instagram.com/uptoskills",
  YOUTUBE: "https://www.youtube.com/@uptoskills",
  WHATSAPP_CHANNEL: "https://whatsapp.com/channel/uptoskills",
  WEBAPP_URL: "https://script.google.com/macros/s/<<DEPLOY_ID>>/exec"
};

function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

var DOMAINS_LIST = [
  "Web Development", "App Development", "Data Science", "Machine Learning",
  "Artificial Intelligence", "Cloud Computing", "DevOps", "Cybersecurity",
  "Blockchain", "UI/UX Design", "Graphic Design", "Digital Marketing",
  "Content Writing", "SEO", "Social Media Marketing", "Business Development",
  "Human Resources", "Finance", "Accounting", "Data Analytics",
  "Python Development", "Java Development", "Full Stack Development",
  "Frontend Development", "Backend Development", "Database Administration",
  "Network Engineering", "Quality Assurance", "Technical Writing",
  "Video Editing", "Animation", "Game Development",
  "Internet of Things (IoT)", "Robotics", "AR/VR Development",
  "Project Management", "Campus Ambassador"
];

function ensureSheetsExist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(MASTER_SHEET);
  if (!master) {
    master = ss.insertSheet(MASTER_SHEET);
    var masterHeaders = [
      "Intern Code", "Onboarding Date", "Name", "Email ID", "Mobile No.",
      "Whatsapp No.", "Current/Last College Name", "Current/Last College Degree Name (Eg. BCA,BBA,BTECH)",
      "Current/Last College Branch Name (Eg. CSE,ECE,AIML)", "Current/Last College Enrollment No.",
      "Domain Name", "Start Date", "End Date", "Current Address", "Permanent Address",
      "Validation Status", "Validation Reason", "Assigned Group Name",
      "Intern Mail Sent", "Intern Mail Timestamp", "TL Mail Sent", "TL Mail Timestamp",
      "Current Status", "Status Reason", "Offer Letter Link", "TL Status Locked"
    ];
    master.getRange(1, 1, 1, masterHeaders.length).setValues([masterHeaders]);
    master.setFrozenRows(1);
  } else {
    ensureMasterColumn(master, "Offer Letter Link");
    ensureMasterColumn(master, "TL Status Locked");
  }
  var proj = ss.getSheetByName(PROJECT_SHEET);
  if (!proj) {
    proj = ss.insertSheet(PROJECT_SHEET);
    var projHeaders = ["Group Name", "Email ID", "WhatsApp Link", "Domain Name"];
    proj.getRange(1, 1, 1, projHeaders.length).setValues([projHeaders]);
    proj.setFrozenRows(1);
  }
  var prog = ss.getSheetByName(PROGRAM_SHEET);
  if (!prog) {
    prog = ss.insertSheet(PROGRAM_SHEET);
    var progHeaders = ["Domain Name", "Project Group Name"];
    prog.getRange(1, 1, 1, progHeaders.length).setValues([progHeaders]);
    prog.setFrozenRows(1);
  }
}

function ensureMasterColumn(sheet, name) {
  var last = sheet.getLastColumn();
  var hdr = sheet.getRange(1, 1, 1, last).getValues()[0];
  for (var i = 0; i < hdr.length; i++) {
    if ((hdr[i] || "").toString().trim() === name) return;
  }
  sheet.getRange(1, last + 1).setValue(name);
}

function makeShareableLink(file) {
  if (!file) return "";
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log("setSharing failed: " + e.toString());
  }
  try { return file.getUrl(); } catch (e) { return ""; }
}

function masterHeaderKeys(sheet) {
  var last = sheet.getLastColumn();
  var hdr = sheet.getRange(1, 1, 1, last).getValues()[0];
  var keys = [];
  for (var i = 0; i < last; i++) {
    keys.push((hdr[i] ? String(hdr[i]) : "Col" + i).replace(/[^a-zA-Z0-9]/g, ""));
  }
  return keys;
}

function _normKey(s) {
  return (s == null ? "" : String(s)).replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
}

function getCol(sheet, key) {
  var last = sheet.getLastColumn();
  var hdr = sheet.getRange(1, 1, 1, last).getValues()[0];
  var nk = _normKey(key);
  for (var i = 0; i < last; i++) {
    if (_normKey(hdr[i]) === nk) return i;
  }
  return -1;
}

function valAt(rowData, sheet, key) {
  var c = getCol(sheet, key);
  return c >= 0 && rowData.length > c ? rowData[c] : "";
}

function setCol(sheet, rowNum, key, value) {
  var c = getCol(sheet, key);
  if (c >= 0) sheet.getRange(rowNum, c + 1).setValue(value);
}

function buildRowArray(sheet, valsObj, existingRow) {
  var last = sheet.getLastColumn();
  var hdr = sheet.getRange(1, 1, 1, last).getValues()[0];
  var arr = [];
  for (var i = 0; i < last; i++) {
    var orig = hdr[i] || "";
    var hk = orig.toString().replace(/[^a-zA-Z0-9]/g, "");
    var matchVal = undefined;
    if (valsObj) {
      if (valsObj[hk] !== undefined) matchVal = valsObj[hk];
      else if (valsObj[orig] !== undefined) matchVal = valsObj[orig];
    }
    if (matchVal !== undefined) arr.push(matchVal);
    else if (existingRow && existingRow[i] !== undefined) arr.push(existingRow[i]);
    else arr.push("");
  }
  return arr;
}

 function resolveWebAppUrl() {
  try {
    var stored = PropertiesService.getScriptProperties().getProperty("WEBAPP_URL");
    if (stored && stored.indexOf("<<") === -1) return stored;
    var url = ScriptApp.getService().getUrl();
    if (url && url.indexOf("<<") === -1) {
      PropertiesService.getScriptProperties().setProperty("WEBAPP_URL", url);
      return url;
    }
  } catch (e) { Logger.log("resolveWebAppUrl: " + e.toString()); }
  return CONFIG.WEBAPP_URL;
 }

 var _DATA_CACHE_KEY = "HRPORTAL_DATA_V1";

 function _getDataCache() {
  try { return CacheService.getScriptCache(); } catch (e) { return null; }
 }

 function _readCache(key) {
  try {
    var c = _getDataCache();
    if (!c) return null;
    var raw = c.get(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) { return null; }
 }

 function _writeCache(key, obj, ttlSec) {
  try {
    var c = _getDataCache();
    if (!c) return;
    c.put(key, JSON.stringify(obj), ttlSec || 300);
  } catch (e) { Logger.log("cache put: " + e.toString()); }
 }

 function clearDataCache() {
  try {
    var c = _getDataCache();
    if (c) c.remove(_DATA_CACHE_KEY + "_records");
  } catch (e) { Logger.log("clearDataCache: " + e.toString()); }
 }

function doGet(e) {
  if (e && e.parameter) {
    var p = e.parameter;
    if (p.data) {
      try { p = JSON.parse(p.data); } catch (_) {}
    }
    if (p && p.action) {
      return jsonResponse(handleApi(p));
    }
  }
  ensureSheetsExist();
  CONFIG.WEBAPP_URL = resolveWebAppUrl();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Intern Onboarding & HR Operations Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    var payload = {};
    if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch (_) {}
    }
    if (e.parameter) {
      var keys = Object.keys(e.parameter);
      for (var k = 0; k < keys.length; k++) {
        if (payload[keys[k]] === undefined) payload[keys[k]] = e.parameter[keys[k]];
      }
    }
    return jsonResponse(handleApi(payload));
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function handleApi(payload) {
  try {
    CONFIG.WEBAPP_URL = resolveWebAppUrl();
    payload = payload || {};
    if (payload.__token !== undefined && payload.__token !== API_TOKEN) {
      return { success: false, error: 'Unauthorized: invalid token' };
    }
    var action = payload.action || null;
    var result;
    switch (action) {
      case 'verifyHRLogin':        result = verifyHRLogin(payload.userId || '', payload.password || ''); break;
      case 'verifyTLLogin':         result = verifyTLLogin(payload.email || '', payload.password || ''); break;
      case 'fetchTLInterns':        result = fetchTLInterns(payload.groupName || ''); break;
      case 'getActiveDomains':     result = getActiveDomains(); break;
      case 'handleInternRegistration': result = handleInternRegistration(payload.formData, payload.pdfFileData); break;
      case 'fetchInternRecords':   result = fetchInternRecords(); break;
      case 'fetchProjectDetails':  result = fetchProjectDetails(); break;
      case 'saveRowModifications': result = saveRowModifications(Number(payload.rowNum), payload.fields); break;
      case 'forceCodeGeneration':  result = forceCodeGeneration(Number(payload.rowNum)); break;
      case 'fireOnboardingMails':  result = fireOnboardingMails(Number(payload.rowNum)); break;
      case 'writeProjectSetting':  result = writeProjectSetting(payload.rowNum, payload.groupName, payload.emailId, payload.whatsappLink, payload.domainName); break;
      case 'updateInternRow':      result = updateInternRow(Number(payload.rowNum), payload.formData, payload.pdfFileData, payload.validationStatus, payload.validationReason, payload.internCode, payload.groupName); break;
      case 'updateInternFullRecord': result = updateInternFullRecord(Number(payload.rowNum), payload.fields); break;
      case 'fetchProgramDetails':  result = fetchProgramDetails(); break;
      case 'writeProgramSetting':  result = writeProgramSetting(Number(payload.rowNum), payload.domainName, payload.groupName); break;
      case 'writeProgramDomain':   result = writeProgramDomain(Number(payload.rowNum), payload.domainName); break;
      case 'deleteProgramSetting': result = deleteProgramSetting(Number(payload.rowNum)); break;
      case 'deleteProjectSetting': result = deleteProgramSetting(Number(payload.rowNum)); break;
      case 'fetchProjectGroupsByDomain': result = fetchProjectGroupsByDomain(payload.domain); break;
      case 'getRegistrationStatus': result = getRegistrationStatus(); break;
      case 'setRegistrationStatus': result = setRegistrationStatus(payload.enabled); break;
      case 'sendTLMail':           result = sendTLMail(Number(payload.rowNum)); break;
      case 'sendBulkEmails':       result = sendBulkEmails(payload.rowNums, payload.type); break;
      case 'testDataConnection':   result = testDataConnection(); break;
      case 'deleteInternRecord':   result = deleteInternRecord(Number(payload.rowNum)); break;
      default:                     result = { success: false, error: 'Unknown action: ' + action };
    }
    return result;
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyHRLogin(userId, password) {
  return userId === HR_USERID && password === HR_PASSKEY;
}

function expectedTLPassword(groupName) {
  if (!groupName) return '';
  var g = groupName.toString().trim();
  if (!g) return '';
  return g.charAt(0).toUpperCase() + g.slice(1) + '@2026';
}

function verifyTLLogin(email, password) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: 'No project groups configured' };
    var data = sheet.getDataRange().getValues();
    var e = (email || '').toString().trim().toLowerCase();
    for (var ri = 1; ri < data.length; ri++) {
      var rowEmail = (data[ri][1] || '').toString().trim().toLowerCase();
      if (rowEmail === e) {
        var groupName = (data[ri][0] || '').toString().trim();
        var expected = expectedTLPassword(groupName);
        if (password === expected) {
          return { success: true, groupName: groupName, email: data[ri][1] || '' };
        }
        return { success: false, error: 'Invalid password for this group email' };
      }
    }
    return { success: false, error: 'Email not found in Project Details' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function fetchTLInterns(groupName) {
  try {
    var base = fetchInternRecords();
    if (!base || !base.success) return base;
    var g = (groupName || '').toString().trim().toLowerCase();
    var filtered = [];
    for (var i = 0; i < base.data.length; i++) {
      var r = base.data[i];
      var ag = ((r['Assigned Group Name'] || r['AssignedGroupName'] || '') + '').toString().trim().toLowerCase();
      if (ag !== g) continue;
      var st = ((r['Current Status'] || '') + '').toString().trim().toLowerCase();
      if (st === 'pending verification' || st === 'pending' || st === '') continue;
      filtered.push(r);
    }
    return { success: true, data: filtered, headers: base.headers, headerKeys: base.headerKeys, count: filtered.length, groupName: groupName };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getActiveDomains() {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var prog = ss.getSheetByName(PROGRAM_SHEET);
    if (prog && prog.getLastRow() > 1) {
      var data = prog.getDataRange().getValues();
      var domains = [];
      for (var di = 1; di < data.length; di++) {
        var name = (data[di][0] || "").toString().trim();
        if (name) domains.push(name);
      }
      return uniq(domains);
    }
  } catch (e) {
    Logger.log("getActiveDomains fallback: " + e.toString());
  }
  return DOMAINS_LIST.slice();
}

function handleInternRegistration(formData, pdfFileData) {
  if (!getRegistrationStatus().enabled) {
    return { success: false, error: "Registration is currently closed by the HR team." };
  }
  try { ensureSheetsExist(); } catch (e) {
    return { success: false, error: "ensureSheetsExist failed: " + e.toString() };
  }
  var ss, masterSheet, projSheet;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    masterSheet = ss.getSheetByName(MASTER_SHEET);
    projSheet = ss.getSheetByName(PROJECT_SHEET);
  } catch (e) {
    return { success: false, error: "Spreadsheet access failed: " + e.toString() };
  }
  if (!masterSheet) return { success: false, error: "Master sheet '" + MASTER_SHEET + "' not found" };

  var today = new Date();
  var tz = Session.getScriptTimeZone();
  var formattedDate = Utilities.formatDate(today, tz, "yyyy-MM-dd");
  var email = (formData && formData.emailId || "").trim().toLowerCase();
  var domain = (formData && formData.domainName || "").trim();

  var validationStatus = "Passed";
  var validationReason = "Verified automatically.";
  var extractedStart = "";
  var extractedEnd = "";
  var uploadedFile = null;

  var requiredFields = ['name','emailId','mobileNo','collegeName','degreeName','branchName','domainName'];
  for (var fi = 0; fi < requiredFields.length; fi++) {
    if (!formData || !formData[requiredFields[fi]] || formData[requiredFields[fi]].toString().trim() === "") {
      return { success: false, error: "Required field '" + requiredFields[fi] + "' is missing" };
    }
  }

  try {
    if (!pdfFileData || !pdfFileData.base64) {
      return { success: false, error: "No PDF file data received" };
    }
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (e) {
    return { success: false, error: "Drive folder access failed: " + e.toString() };
  }

  var decodedBytes;
  try {
    decodedBytes = Utilities.base64Decode(pdfFileData.base64);
  } catch (e) {
    return { success: false, error: "base64Decode failed: " + e.toString() };
  }

  try {
    var blob = Utilities.newBlob(decodedBytes, pdfFileData.type || "application/pdf", pdfFileData.name || "offer.pdf");
    uploadedFile = folder.createFile(blob);
  } catch (e) {
    return { success: false, error: "File creation in Drive failed: " + e.toString() };
  }

  try {
    var docContent = extractTextFromPdf(uploadedFile.getId());
    if (!docContent || docContent.length === 0) {
      validationStatus = "Failed";
      validationReason = "Offer letter PDF text could not be extracted. Enable the Drive API advanced service (Editor > Resources > Advanced Google services > Drive API = ON) and re-upload the offer letter.";
    } else {
      if (docContent.toLowerCase().indexOf(domain.toLowerCase()) === -1) {
        validationStatus = "Failed";
        validationReason = "Offer letter text check mismatch: Target domain text not found inside PDF document.";
      }
      var extractedDates = extractDatesFromText(docContent);
      if (extractedDates && extractedDates.length >= 2) {
        extractedStart = Utilities.formatDate(extractedDates[0], tz, "yyyy-MM-dd");
        extractedEnd = Utilities.formatDate(extractedDates[1], tz, "yyyy-MM-dd");
      }
    }
  } catch (e) {
    Logger.log("PDF text extraction failed: " + e.toString());
    validationStatus = "Failed";
    validationReason = "Offer letter PDF processing error: " + e.toString();
  }

  try {
    if (validationStatus === "Passed") {
      var masterData = masterSheet.getDataRange().getValues();
      var activeCount = 0;
      var hasCampusAmbassador = false;
      var duplicateDomain = false;
      for (var mi = 1; mi < masterData.length; mi++) {
        var dbEmail = (valAt(masterData[mi], masterSheet, "Email ID") || "").toString().trim().toLowerCase();
        var dbDomain = (valAt(masterData[mi], masterSheet, "Domain Name") || "").toString().trim();
        var dbStatus = (valAt(masterData[mi], masterSheet, "Current Status") || "").toString().trim();
        if (dbEmail === email && dbStatus !== "Completed") {
          activeCount++;
          if (dbDomain.toLowerCase() === "campus ambassador") hasCampusAmbassador = true;
          if (dbDomain.toLowerCase() === domain.toLowerCase()) duplicateDomain = true;
        }
      }
      if (duplicateDomain) {
        validationStatus = "Failed";
        validationReason = "Duplicate entry check: Applicant is already active inside this identical domain track.";
      } else if (activeCount >= 1 && !hasCampusAmbassador && domain.toLowerCase() !== "campus ambassador") {
        validationStatus = "Failed";
        validationReason = "Multiple simultaneous allocations require one domain role to explicitly be Campus Ambassador.";
      }
    }
  } catch (e) {
    return { success: false, error: "Duplicate analysis failed: " + e.toString() };
  }

  try {
    var internCode = "";
    var offerLink = "";
    if (uploadedFile) {
      offerLink = makeShareableLink(uploadedFile);
    }
    if (validationStatus === "Passed" && uploadedFile) {
      var yr = Utilities.formatDate(today, tz, "yyyy");
      var dd = Utilities.formatDate(today, tz, "dd");
      var mo = Utilities.formatDate(today, tz, "MM");
      var rand = Math.floor(10000 + Math.random() * 90000);
      internCode = "USINT" + yr + dd + mo + rand;
      uploadedFile.setName(internCode + "_OfferLetter.pdf");
    }
    var groupName = "Unassigned";
    if (projSheet) {
      var projData = projSheet.getDataRange().getValues();
      for (var pi = 1; pi < projData.length; pi++) {
        var hdDomList = (projData[pi][3] || "").toString().toLowerCase().split(',');
        var hdMatch = false;
        for (var hddl = 0; hddl < hdDomList.length; hddl++) { if (hdDomList[hddl].trim() === domain.toLowerCase()) { hdMatch = true; break; } }
        if (hdMatch) {
          groupName = projData[pi][0] || "Unassigned";
          break;
        }
      }
    }
    if (groupName === "Unassigned") {
      try {
        var progSheet = ss.getSheetByName(PROGRAM_SHEET);
        if (progSheet && progSheet.getLastRow() > 1) {
          var progData = progSheet.getDataRange().getValues();
          for (var pgi = 1; pgi < progData.length; pgi++) {
            if ((progData[pgi][0] || "").toString().toLowerCase() === domain.toLowerCase()) {
              groupName = progData[pgi][1] || "Unassigned";
              break;
            }
          }
        }
      } catch (e) {
        Logger.log("Program sheet lookup non-fatal: " + e.toString());
      }
    }
    var currParts = [];
    if (formData.currAddr1) currParts.push(formData.currAddr1);
    if (formData.currAddr2) currParts.push(formData.currAddr2);
    if (formData.currPo) currParts.push(formData.currPo);
    if (formData.currDistrict) currParts.push(formData.currDistrict);
    if (formData.currState) currParts.push(formData.currState);
    if (formData.currCountry) currParts.push(formData.currCountry);
    if (formData.currPin) currParts.push(formData.currPin);
    var currentAddr = currParts.join(", ");

    var permanentAddr;
    if (formData.sameAsCurrent === true || formData.sameAsCurrent === 'true') {
      permanentAddr = currentAddr;
    } else {
      var permParts = [];
      if (formData.permAddr1) permParts.push(formData.permAddr1);
      if (formData.permAddr2) permParts.push(formData.permAddr2);
      if (formData.permPo) permParts.push(formData.permPo);
      if (formData.permDistrict) permParts.push(formData.permDistrict);
      if (formData.permState) permParts.push(formData.permState);
      if (formData.permCountry) permParts.push(formData.permCountry);
      if (formData.permPin) permParts.push(formData.permPin);
      permanentAddr = permParts.join(", ");
    }
    var rowVals = {
      "Intern Code": internCode,
      "Onboarding Date": formattedDate,
      "Name": formData.name || "",
      "Email ID": email,
      "Mobile No.": formData.mobileNo || "",
      "Whatsapp No.": formData.whatsappNo || "",
      "Current/Last College Name": formData.collegeName || "",
      "Current/Last College Degree Name (Eg. BCA,BBA,BTECH)": formData.degreeName || "",
      "Current/Last College Branch Name (Eg. CSE,ECE,AIML)": formData.branchName || "",
      "Current/Last College Enrollment No.": formData.enrollmentNo || "",
      "Domain Name": domain,
      "Start Date": extractedStart,
      "End Date": extractedEnd,
      "Current Address": currentAddr,
      "Permanent Address": permanentAddr,
      "Validation Status": validationStatus,
      "Validation Reason": validationReason,
      "Assigned Group Name": groupName,
      "Intern Mail Sent": "No",
      "Intern Mail Timestamp": "",
      "TL Mail Sent": "No",
      "TL Mail Timestamp": "",
      "Current Status": validationStatus === "Passed" ? "Active" : "Pending Verification",
      "Status Reason": "",
      "Offer Letter Link": offerLink
    };
    var finalRow = buildRowArray(masterSheet, rowVals, null);
    masterSheet.appendRow(finalRow);
    var newRowNum = masterSheet.getLastRow();
    return { success: true, status: validationStatus, reason: validationReason, code: internCode, rowNumber: newRowNum };
  } catch (e) {
    return { success: false, error: "Final save failed: " + e.toString() };
  }
}

function testDataConnection() {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var info = [];
    for (var si = 0; si < sheets.length; si++) {
      info.push(sheets[si].getName() + " (rows: " + sheets[si].getLastRow() + ")");
    }
    var master = ss.getSheetByName(MASTER_SHEET);
    var proj = ss.getSheetByName(PROJECT_SHEET);
    return {
      success: true,
      MASTER_SHEET: MASTER_SHEET,
      masterFound: !!master,
      masterRows: master ? master.getLastRow() : 0,
      projFound: !!proj,
      projRows: proj ? proj.getLastRow() : 0,
      allSheets: info,
      scriptId: ScriptApp.getScriptId(),
      timeZone: Session.getScriptTimeZone()
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

 function fetchInternRecords() {
  try {
    var cached = _readCache(_DATA_CACHE_KEY + "_records");
    if (cached && cached.success) return cached;
  } catch (e) {}
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
    if (!sheet) {
      Logger.log("fetchInternRecords: sheet '" + MASTER_SHEET + "' not found");
      return { success: false, error: "Sheet not found", data: [] };
    }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, count: 0, headers: [], headerKeys: [], data: [] };
    }
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0] ? data[0].slice() : [];
    var headerKeys = [];
    for (var hi = 0; hi < headers.length; hi++) {
      var headerStr = headers[hi] ? String(headers[hi]) : "Col" + hi;
      headerKeys.push(headerStr.replace(/[^a-zA-Z0-9]/g, ""));
    }
    var maxWidth = headers.length;
    for (var ri = 1; ri < data.length; ri++) {
      if (data[ri].length > maxWidth) maxWidth = data[ri].length;
    }
    for (var pi = headers.length; pi < maxWidth; pi++) {
      headers[pi] = "Col" + (pi + 1);
      headerKeys[pi] = "Col" + (pi + 1);
    }
    var records = [];
    for (var ri = 1; ri < data.length; ri++) {
      var record = { rowNumber: ri + 1 };
      for (var hi = 0; hi < maxWidth; hi++) {
        var cellVal = (hi < data[ri].length && data[ri][hi] !== undefined && data[ri][hi] !== null) ? data[ri][hi] : "";
        record[headerKeys[hi]] = cellVal;
        record[headers[hi]] = cellVal;
      }
      records.push(record);
    }
    Logger.log("fetchInternRecords: returning " + records.length + " records, " + maxWidth + " columns");
    var payload = { success: true, count: records.length, headers: headers, headerKeys: headerKeys, data: records };
    _writeCache(_DATA_CACHE_KEY + "_records", payload, 600);
    return payload;
  } catch (e) {
    Logger.log("fetchInternRecords error: " + e.toString());
    return { success: false, error: e.toString(), data: [] };
  }
 }

function fetchProjectDetails() {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getDataRange().getValues();
    var records = [];
    for (var ri = 1; ri < data.length; ri++) {
      records.push({ rowNumber: ri + 1, groupName: data[ri][0] || "", emailId: data[ri][1] || "", whatsappLink: data[ri][2] || "", domainName: data[ri][3] || "" });
    }
    return records;
  } catch (e) {
    Logger.log("fetchProjectDetails error: " + e.toString());
    return [];
  }
}

 function getDashboardInit() {
  try {
    ensureSheetsExist();
    var rec = fetchInternRecords();
    var prog = fetchProgramDetails();
    var proj = fetchProjectDetails();
    var domains = [];
    if (prog && prog.length) {
      for (var di = 0; di < prog.length; di++) {
        var dn = (prog[di].domainName || "").toString().trim();
        if (dn && domains.indexOf(dn) === -1) domains.push(dn);
      }
    }
    if (!domains.length) domains = getActiveDomains() || [];
    var reg = getRegistrationStatus();
    return {
      success: true,
      records: rec.success ? rec.data : [],
      headers: rec.success ? rec.headers : [],
      headerKeys: rec.success ? rec.headerKeys : [],
      count: rec.success ? rec.count : 0,
      program: prog || [],
      project: proj || [],
      domains: domains,
      regEnabled: reg && reg.enabled
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function saveRowModifications(rowNum, fields) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
    if (!sheet) return { success: false, error: "Master sheet not found" };
    if (fields && fields.Name !== undefined) setCol(sheet, rowNum, "Name", fields.Name);
    if (fields && fields.EmailID !== undefined) setCol(sheet, rowNum, "Email ID", fields.EmailID);
    if (fields && fields.DomainName !== undefined) setCol(sheet, rowNum, "Domain Name", fields.DomainName);
    if (fields && fields.CurrentStatus !== undefined) setCol(sheet, rowNum, "Current Status", fields.CurrentStatus);
    if (fields && fields.StatusReason !== undefined) setCol(sheet, rowNum, "Status Reason", fields.StatusReason);
    clearDataCache();
    return { success: true };
  } catch (e) {
    Logger.log("saveRowModifications error: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function forceCodeGeneration(rowNum) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
    if (!sheet) return { success: false, error: "Master sheet not found" };
    var today = new Date();
    var tz = Session.getScriptTimeZone();
    var yr = Utilities.formatDate(today, tz, "yyyy");
    var dd = Utilities.formatDate(today, tz, "dd");
    var mo = Utilities.formatDate(today, tz, "MM");
    var rand = Math.floor(10000 + Math.random() * 90000);
    var code = "USINT" + yr + dd + mo + rand;
    setCol(sheet, rowNum, "Intern Code", code);
    setCol(sheet, rowNum, "Validation Status", "Passed");
    setCol(sheet, rowNum, "Validation Reason", "Manually verified and bypassed by HR Admin.");
    setCol(sheet, rowNum, "Current Status", "Active");
    clearDataCache();
    return { success: true, code: code };
  } catch (e) {
    Logger.log("forceCodeGeneration error: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function resolveGroupInfo(rowData, projData) {
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
  var domain = (valAt(rowData, masterSheet, "Domain Name") || "").toString().toLowerCase();
  var assigned = (valAt(rowData, masterSheet, "Assigned Group Name") || "").toString().trim();
  var groupName = assigned || "Assigned Team Track";
  var tlEmail = "";
  var whatsappLink = "";
  if (projData && projData.length) {
    if (assigned) {
      for (var pi = 1; pi < projData.length; pi++) {
        if ((projData[pi][0] || "").toString().trim().toLowerCase() === assigned.toLowerCase()) {
          groupName = projData[pi][0] || assigned;
          tlEmail = projData[pi][1] || "";
          whatsappLink = projData[pi][2] || "";
          break;
        }
      }
    }
    if (!tlEmail) {
      for (var pj = 1; pj < projData.length; pj++) {
        var domList = (projData[pj][3] || "").toString().toLowerCase().split(',');
        var matched = false;
        for (var dl = 0; dl < domList.length; dl++) { if (domList[dl].trim() === domain) { matched = true; break; } }
        if (matched) {
          if (!assigned) groupName = projData[pj][0] || groupName;
          tlEmail = projData[pj][1] || tlEmail;
          whatsappLink = projData[pj][2] || whatsappLink;
          break;
        }
      }
    }
  }
  return { groupName: groupName, tlEmail: tlEmail, whatsappLink: whatsappLink };
}

function fmtDate(v) {
  if (v === undefined || v === null || v === '') return '';
  var d;
  if (v instanceof Date) {
    d = v;
  } else {
    var s = String(v).trim();
    if (!s) return '';
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      d = new Date(s + (s.length === 10 ? 'T00:00:00' : ''));
    } else if (/^\d{1,2}[-/]\d{1,2}[-/]\d{4}/.test(s)) {
      var p = s.split(/[-/]/); d = new Date(p[2], p[1] - 1, p[0]);
    } else {
      d = new Date(s);
    }
  }
  if (isNaN(d.getTime())) return String(v);
  var dd = ('0' + d.getDate()).slice(-2);
  var mm = ('0' + (d.getMonth() + 1)).slice(-2);
  var yyyy = d.getFullYear();
  return dd + '-' + mm + '-' + yyyy;
}

function srTlRowHtml(rowData, tz) {
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
  var d = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var onboarding = fmtDate(valAt(rowData, masterSheet, "Onboarding Date") || d);
  var code = valAt(rowData, masterSheet, "Intern Code") || "";
  var name = valAt(rowData, masterSheet, "Name") || "";
  var email = valAt(rowData, masterSheet, "Email ID") || "";
  var mobile = valAt(rowData, masterSheet, "Mobile No.") || "";
  var whatsapp = valAt(rowData, masterSheet, "Whatsapp No.") || "";
  var domain = valAt(rowData, masterSheet, "Domain Name") || "";
  var start = fmtDate(valAt(rowData, masterSheet, "Start Date") || "");
  var end = fmtDate(valAt(rowData, masterSheet, "End Date") || "");
  return "<tr>"
    + '<td style="padding:10px; border-bottom:1px solid #e2e8f0; border-right:1px solid #e2e8f0;">' + escapeHtml(onboarding) + "</td>"
    + '<td style="padding:10px; border-bottom:1px solid #e2e8f0; border-right:1px solid #e2e8f0;">' + escapeHtml(code) + "</td>"
    + '<td style="padding:10px; border-bottom:1px solid #e2e8f0; border-right:1px solid #e2e8f0;">' + escapeHtml(name) + "</td>"
    + '<td style="padding:10px; border-bottom:1px solid #e2e8f0; border-right:1px solid #e2e8f0;">' + escapeHtml(email) + "</td>"
    + '<td style="padding:10px; border-bottom:1px solid #e2e8f0; border-right:1px solid #e2e8f0;">' + escapeHtml(mobile) + "</td>"
    + '<td style="padding:10px; border-bottom:1px solid #e2e8f0; border-right:1px solid #e2e8f0;">' + escapeHtml(whatsapp) + "</td>"
    + '<td style="padding:10px; border-bottom:1px solid #e2e8f0; border-right:1px solid #e2e8f0;">' + escapeHtml(domain) + "</td>"
    + '<td style="padding:10px; border-bottom:1px solid #e2e8f0; border-right:1px solid #e2e8f0;">' + escapeHtml(start) + "</td>"
    + '<td style="padding:10px; border-bottom:1px solid #e2e8f0;">' + escapeHtml(end) + "</td>"
    + "</tr>";
}

function fireOnboardingMails(rowNum) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET);
    var projSheet = ss.getSheetByName(PROJECT_SHEET);
    if (!masterSheet) return { success: false, error: "Master sheet not found" };
    var tz = Session.getScriptTimeZone();
    var rowData = masterSheet.getRange(rowNum, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    var internCode = valAt(rowData, masterSheet, "Intern Code");
    var internName = valAt(rowData, masterSheet, "Name");
    var internEmail = valAt(rowData, masterSheet, "Email ID");
    var mobileNum = valAt(rowData, masterSheet, "Mobile No.");
    var domain = valAt(rowData, masterSheet, "Domain Name");
    var onboardingDate = valAt(rowData, masterSheet, "Onboarding Date");
    var projData = projSheet ? projSheet.getDataRange().getValues() : [];
    var info = resolveGroupInfo(rowData, projData);
    var groupInfo = { groupName: info.groupName, email: info.tlEmail, whatsapp: info.whatsappLink };
    var timestamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    if (internEmail) {
      var internSubject = "Welcome Aboard! Your Internship Onboarding Details";
      var internHtml = internWelcomeTemplate({ name: internName, code: internCode }, groupInfo);
      var internPlain = "Hello " + internName + ",\n\nYour onboarding is verified.\nIntern Code: " + internCode
        + "\nGroup: " + info.groupName + "\nTeam Community: " + info.whatsappLink
        + "\nOnboarding Date: " + fmtDate(onboardingDate) + "\n\nBest Regards,\nHR Operations Team";
      MailApp.sendEmail({ to: internEmail, subject: internSubject, body: internPlain, htmlBody: internHtml });
      setCol(masterSheet, rowNum, "Intern Mail Sent", "Yes");
      setCol(masterSheet, rowNum, "Intern Mail Timestamp", timestamp);
    }
    if (info.tlEmail) {
      var tlSubject = "New Intern Assignment - " + internName + " (" + domain + ")";
      var tlHtml = srTlSummaryTemplate(info.groupName, srTlRowHtml(rowData, tz), 1);
      var tlPlain = "Hello,\n\nA new intern has been assigned to your team.\nName: " + internName
        + "\nEmail: " + internEmail + "\nMobile: " + mobileNum + "\nDomain: " + domain + "\n\nBest Regards,\nHR Operations Team";
      MailApp.sendEmail({ to: info.tlEmail, subject: tlSubject, body: tlPlain, htmlBody: tlHtml });
      setCol(masterSheet, rowNum, "TL Mail Sent", "Yes");
      setCol(masterSheet, rowNum, "TL Mail Timestamp", timestamp);
    }
    clearDataCache();
    return { success: true };
  } catch (e) {
    Logger.log("fireOnboardingMails error: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function writeProjectSetting(rowNum, groupName, emailId, whatsappLink, domainName) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_SHEET);
    if (!sheet) return { success: false, error: "Project sheet not found" };
    var savedRow = null;
    if (rowNum === 0 || rowNum === "0" || rowNum === "") {
      sheet.appendRow([groupName, emailId, whatsappLink, domainName]);
      savedRow = sheet.getLastRow();
    } else {
      sheet.getRange(rowNum, 1).setValue(groupName);
      sheet.getRange(rowNum, 2).setValue(emailId);
      sheet.getRange(rowNum, 3).setValue(whatsappLink);
      sheet.getRange(rowNum, 4).setValue(domainName);
      savedRow = rowNum;
    }
    clearDataCache();
    return { success: true, rowNumber: savedRow };
  } catch (e) {
    Logger.log("writeProjectSetting error: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function deleteInternRecord(rowNum) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
    if (!sheet) return { success: false, error: "Master sheet not found" };
    sheet.deleteRow(rowNum);
    clearDataCache();
    return { success: true };
  } catch (e) {
    Logger.log("deleteInternRecord error: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function updateInternRow(rowNum, formData, pdfFileData, validationStatus, validationReason, internCode, groupName) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET);
    var projSheet = ss.getSheetByName(PROJECT_SHEET);
    if (!masterSheet) return { success: false, error: "Master sheet not found" };
    var today = new Date();
    var tz = Session.getScriptTimeZone();
    var email = (formData && formData.emailId || "").trim().toLowerCase();
    var domain = (formData && formData.domainName || "").trim();
    var hasPdf = !!(pdfFileData && pdfFileData.base64);

    var finalStatus = "Passed";
    var finalReason = "Verified automatically.";
    var extractedStart = "";
    var extractedEnd = "";
    var uploadedFile = null;
    var existingRow = masterSheet.getRange(rowNum, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    var code = (internCode || (existingRow[0] !== undefined ? existingRow[0] : "")) || "";

    if (hasPdf) {
      try { var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID); }
      catch (e) { return { success: false, error: "Drive folder access failed: " + e.toString() }; }
      var decodedBytes;
      try { decodedBytes = Utilities.base64Decode(pdfFileData.base64); }
      catch (e) { return { success: false, error: "base64Decode failed: " + e.toString() }; }
      try {
        var blob = Utilities.newBlob(decodedBytes, pdfFileData.type || "application/pdf", pdfFileData.name || "offer.pdf");
        uploadedFile = folder.createFile(blob);
      } catch (e) { return { success: false, error: "File creation in Drive failed: " + e.toString() }; }
       try {
         var docContent = extractTextFromPdf(uploadedFile.getId());
         if (!docContent || docContent.length === 0) {
           finalStatus = "Failed";
           finalReason = "Offer letter PDF text could not be extracted. Enable the Drive API advanced service (Editor > Resources > Advanced Google services > Drive API = ON) and re-upload.";
         } else {
           if (docContent.toLowerCase().indexOf(domain.toLowerCase()) === -1) {
             finalStatus = "Failed";
             finalReason = "Offer letter text check mismatch: Target domain text not found inside PDF document.";
           }
           var extractedDates = extractDatesFromText(docContent);
           if (extractedDates && extractedDates.length >= 2) {
             extractedStart = Utilities.formatDate(extractedDates[0], tz, "yyyy-MM-dd");
             extractedEnd = Utilities.formatDate(extractedDates[1], tz, "yyyy-MM-dd");
           }
         }
       } catch (e) {
         Logger.log("PDF extraction failed: " + e.toString());
         finalStatus = "Failed";
         finalReason = "Offer letter PDF processing error: " + e.toString();
       }
      try {
        if (finalStatus === "Passed") {
          var masterData = masterSheet.getDataRange().getValues();
          var activeCount = 0, hasCA = false, dupDomain = false;
          for (var mi = 1; mi < masterData.length; mi++) {
            var dbEmail = (valAt(masterData[mi], masterSheet, "Email ID") || "").toString().trim().toLowerCase();
            var dbDomain = (valAt(masterData[mi], masterSheet, "Domain Name") || "").toString().trim();
            var dbStatus = (valAt(masterData[mi], masterSheet, "Current Status") || "").toString().trim();
            if (dbEmail === email && dbStatus !== "Completed" && mi + 1 !== rowNum) {
              activeCount++;
              if (dbDomain.toLowerCase() === "campus ambassador") hasCA = true;
              if (dbDomain.toLowerCase() === domain.toLowerCase()) dupDomain = true;
            }
          }
          if (dupDomain) {
            finalStatus = "Failed";
            finalReason = "Duplicate entry check: Applicant is already active inside this identical domain track.";
          } else if (activeCount >= 1 && !hasCA && domain.toLowerCase() !== "campus ambassador") {
            finalStatus = "Failed";
            finalReason = "Multiple simultaneous allocations require one domain role to explicitly be Campus Ambassador.";
          }
        }
      } catch (e) { return { success: false, error: "Duplicate analysis failed: " + e.toString() }; }
       var offerLink = undefined;
       if (uploadedFile) {
         offerLink = makeShareableLink(uploadedFile);
       }
       if (finalStatus === "Passed" && uploadedFile) {
        var yr = Utilities.formatDate(today, tz, "yyyy");
        var dd = Utilities.formatDate(today, tz, "dd");
        var mo = Utilities.formatDate(today, tz, "MM");
        var rand = Math.floor(10000 + Math.random() * 90000);
        code = "USINT" + yr + dd + mo + rand;
        uploadedFile.setName(code + "_OfferLetter.pdf");
      }
    } else {
      finalStatus = (typeof validationStatus === "string" && validationStatus !== "") ? validationStatus : "Passed";
      finalReason = (typeof validationReason === "string" && validationReason !== "") ? validationReason : "Updated by HR.";
    }

    var finalGroupName = groupName || "Unassigned";
    if (finalGroupName === "Unassigned" && projSheet) {
      var projData = projSheet.getDataRange().getValues();
      for (var pi = 1; pi < projData.length; pi++) {
        var pjDomList = (projData[pi][3] || "").toString().toLowerCase().split(',');
        var pjMatch = false;
        for (var pjdl = 0; pjdl < pjDomList.length; pjdl++) { if (pjDomList[pjdl].trim() === domain.toLowerCase()) { pjMatch = true; break; } }
        if (pjMatch) {
          finalGroupName = projData[pi][0] || "Unassigned";
          break;
        }
      }
    }
    if (finalGroupName === "Unassigned") {
      try {
        var progSheet = ss.getSheetByName(PROGRAM_SHEET);
        if (progSheet && progSheet.getLastRow() > 1) {
          var progData = progSheet.getDataRange().getValues();
          for (var pgi = 1; pgi < progData.length; pgi++) {
            if ((progData[pgi][0] || "").toString().toLowerCase() === domain.toLowerCase()) {
              finalGroupName = progData[pgi][1] || "Unassigned";
              break;
            }
          }
        }
      } catch (e) { Logger.log("Program sheet lookup non-fatal: " + e.toString()); }
    }

    var currParts = [];
    if (formData.currAddr1) currParts.push(formData.currAddr1);
    if (formData.currAddr2) currParts.push(formData.currAddr2);
    if (formData.currPo) currParts.push(formData.currPo);
    if (formData.currDistrict) currParts.push(formData.currDistrict);
    if (formData.currState) currParts.push(formData.currState);
    if (formData.currCountry) currParts.push(formData.currCountry);
    if (formData.currPin) currParts.push(formData.currPin);
    var currentAddr = currParts.join(", ");
    var permanentAddr;
    if (formData.sameAsCurrent === true || formData.sameAsCurrent === 'true') {
      permanentAddr = currentAddr;
    } else {
      var permParts = [];
      if (formData.permAddr1) permParts.push(formData.permAddr1);
      if (formData.permAddr2) permParts.push(formData.permAddr2);
      if (formData.permPo) permParts.push(formData.permPo);
      if (formData.permDistrict) permParts.push(formData.permDistrict);
      if (formData.permState) permParts.push(formData.permState);
      if (formData.permCountry) permParts.push(formData.permCountry);
      if (formData.permPin) permParts.push(formData.permPin);
      permanentAddr = permParts.join(", ");
    }
    var rowVals = {
      "Intern Code": code,
      "Onboarding Date": formData.onboardingDate || Utilities.formatDate(today, tz, "yyyy-MM-dd"),
      "Name": formData.name || "",
      "Email ID": email,
      "Mobile No.": formData.mobileNo || "",
      "Whatsapp No.": formData.whatsappNo || "",
      "Current/Last College Name": formData.collegeName || "",
      "Current/Last College Degree Name (Eg. BCA,BBA,BTECH)": formData.degreeName || "",
      "Current/Last College Branch Name (Eg. CSE,ECE,AIML)": formData.branchName || "",
      "Current/Last College Enrollment No.": formData.enrollmentNo || "",
      "Domain Name": domain,
      "Start Date": extractedStart,
      "End Date": extractedEnd,
      "Current Address": currentAddr,
      "Permanent Address": permanentAddr,
      "Validation Status": finalStatus,
      "Validation Reason": finalReason,
      "Assigned Group Name": finalGroupName,
      "Intern Mail Sent": "No",
      "Intern Mail Timestamp": "",
      "TL Mail Sent": "No",
      "TL Mail Timestamp": "",
      "Current Status": finalStatus === "Passed" ? "Active" : "Pending Verification",
      "Status Reason": "",
      "Offer Letter Link": offerLink
    };
    var finalRow = buildRowArray(masterSheet, rowVals, existingRow);
    masterSheet.getRange(rowNum, 1, 1, finalRow.length).setValues([finalRow]);
    try {
      if (finalStatus === "Failed" && email) {
        var webUrl = resolveWebAppUrl();
        var vHtml = validationFailedTemplate(formData.name || "", finalReason, webUrl);
        var vPlain = "Dear " + (formData.name || "") + ",\n\n" + finalReason + "\n\nPlease correct your records: " + webUrl;
        MailApp.sendEmail({ to: email, subject: "Action Required: Onboarding Data Verification", body: vPlain, htmlBody: vHtml });
      }
    } catch (me) { Logger.log("validation mail error: " + me.toString()); }
    clearDataCache();
    return { success: true, status: finalStatus, reason: finalReason, code: code, groupName: finalGroupName };
  } catch (e) {
    return { success: false, error: "Update failed: " + e.toString() };
  }
}

function updateInternFullRecord(rowNum, fields) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
    if (!sheet) return { success: false, error: "Master sheet not found" };
    var lastCol = sheet.getLastColumn();
    var headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var existingRow = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var row = [];
    for (var ci = 0; ci < headerVals.length; ci++) {
      var k = (headerVals[ci] ? String(headerVals[ci]) : "Col" + ci).replace(/[^a-zA-Z0-9]/g, "");
      if (fields && fields[k] !== undefined) row.push(fields[k]);
      else if (existingRow[ci] !== undefined) row.push(existingRow[ci]);
      else row.push("");
    }
    sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
    clearDataCache();
    return { success: true };
  } catch (e) {
    return { success: false, error: "Update failed: " + e.toString() };
  }
}

function addInternRecord(fields) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
    if (!sheet) return { success: false, error: "Master sheet not found" };
    var lastCol = sheet.getLastColumn();
    var headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var row = [];
    for (var ci = 0; ci < headerVals.length; ci++) {
      var k = (headerVals[ci] ? String(headerVals[ci]) : "Col" + ci).replace(/[^a-zA-Z0-9]/g, "");
      row.push((fields && fields[k] !== undefined) ? fields[k] : "");
    }
    sheet.appendRow(row);
    clearDataCache();
    return { success: true, rowNumber: sheet.getLastRow() };
  } catch (e) {
    return { success: false, error: "Add failed: " + e.toString() };
  }
}

function fetchProgramDetails() {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROGRAM_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < data.length; i++) {
      out.push({ rowNumber: i + 1, domainName: data[i][0] || "", groupName: data[i][1] || "" });
    }
    return out;
  } catch (e) {
    Logger.log("fetchProgramDetails error: " + e.toString());
    return [];
  }
}

function writeProgramSetting(rowNum, domainName, groupName) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROGRAM_SHEET);
    if (!sheet) return { success: false, error: "Program sheet not found" };
    var savedRow = null;
    if (!rowNum || rowNum === "0" || rowNum === 0) {
      sheet.appendRow([domainName, groupName]);
      savedRow = sheet.getLastRow();
    } else {
      sheet.getRange(rowNum, 1).setValue(domainName);
      sheet.getRange(rowNum, 2).setValue(groupName);
      savedRow = rowNum;
    }
    clearDataCache();
    return { success: true, rowNumber: savedRow };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function writeProgramDomain(rowNum, domainName) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROGRAM_SHEET);
    if (!sheet) return { success: false, error: "Program sheet not found" };
    var savedRow = null;
    if (!rowNum || rowNum === "0" || rowNum === 0) {
      sheet.appendRow([domainName, ""]);
      savedRow = sheet.getLastRow();
    } else {
      sheet.getRange(rowNum, 1).setValue(domainName);
      savedRow = rowNum;
    }
    clearDataCache();
    return { success: true, rowNumber: savedRow };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteProgramSetting(rowNum) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROGRAM_SHEET);
    if (!sheet) return { success: false, error: "Program sheet not found" };
    sheet.deleteRow(rowNum);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteProjectSetting(rowNum) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_SHEET);
    if (!sheet) return { success: false, error: "Project sheet not found" };
    sheet.deleteRow(rowNum);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function fetchProjectGroupsByDomain(domain) {
  try {
    ensureSheetsExist();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getDataRange().getValues();
    var out = [];
    var d = (domain || "").toString().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      var ddList = (data[i][3] || "").toString().toLowerCase().split(',');
      var ddMatch = false;
      for (var ddl = 0; ddl < ddList.length; ddl++) { if (ddList[ddl].trim() === d) { ddMatch = true; break; } }
      if (ddMatch) {
        var g = data[i][0] || "";
        if (g && out.indexOf(g) === -1) out.push(g);
      }
    }
    return out;
  } catch (e) {
    Logger.log("fetchProjectGroupsByDomain error: " + e.toString());
    return [];
  }
}

function getRegistrationStatus() {
  try {
    var props = PropertiesService.getScriptProperties();
    var v = props.getProperty("REG_OPEN");
    return { success: true, enabled: v === null ? true : (v === "true") };
  } catch (e) {
    return { success: false, error: e.toString(), enabled: true };
  }
}

function setRegistrationStatus(enabled) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty("REG_OPEN", enabled ? "true" : "false");
    return { success: true, enabled: !!enabled };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function sendTLMail(rowNum) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET);
    var projSheet = ss.getSheetByName(PROJECT_SHEET);
    if (!masterSheet) return { success: false, error: "Master sheet not found" };
    var tz = Session.getScriptTimeZone();
    var lastCol = masterSheet.getLastColumn();
    var rowData = masterSheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var internName = valAt(rowData, masterSheet, "Name") || "";
    var internEmail = valAt(rowData, masterSheet, "Email ID") || "";
    var mobileNum = valAt(rowData, masterSheet, "Mobile No.") || "";
    var domain = valAt(rowData, masterSheet, "Domain Name") || "";
    var onboardingDate = valAt(rowData, masterSheet, "Onboarding Date") || "";
    var projData = projSheet ? projSheet.getDataRange().getValues() : [];
    var info = resolveGroupInfo(rowData, projData);
    if (!info.tlEmail) return { success: false, error: "No Team Lead email configured for group: " + info.groupName };
    var ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    var tlSubject = "New Intern Assignment - " + internName + " (" + domain + ")";
    var tlHtml = srTlSummaryTemplate(info.groupName, srTlRowHtml(rowData, tz), 1);
    var tlPlain = "Hello,\n\nA new intern has been assigned to your team.\nName: " + internName
      + "\nEmail: " + internEmail + "\nMobile: " + mobileNum + "\nDomain: " + domain + "\n\nBest Regards,\nHR Operations Team";
    MailApp.sendEmail({ to: info.tlEmail, subject: tlSubject, body: tlPlain, htmlBody: tlHtml });
    setCol(masterSheet, rowNum, "TL Mail Sent", "Yes");
    setCol(masterSheet, rowNum, "TL Mail Timestamp", ts);
    clearDataCache();
    return { success: true };
  } catch (e) {
    Logger.log("sendTLMail error: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function sendBulkEmails(rowNums, type) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET);
    var projSheet = ss.getSheetByName(PROJECT_SHEET);
    if (!masterSheet) return { success: false, error: "Master sheet not found" };
    var tz = Session.getScriptTimeZone();
    var ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    var projData = projSheet ? projSheet.getDataRange().getValues() : [];
    var sent = 0;
    if (type === 'tl') {
      var byGroup = {};
      for (var i = 0; i < (rowNums || []).length; i++) {
        var rn = Number(rowNums[i]);
        var rd = masterSheet.getRange(rn, 1, 1, masterSheet.getLastColumn()).getValues()[0];
        var inf = resolveGroupInfo(rd, projData);
        if (!inf.tlEmail) continue;
        var key = inf.tlEmail.toLowerCase();
        if (!byGroup[key]) byGroup[key] = { group: inf.groupName, email: inf.tlEmail, rows: "", total: 0 };
        byGroup[key].rows += srTlRowHtml(rd, tz);
        byGroup[key].total += 1;
      }
      for (var ek in byGroup) {
        var g = byGroup[ek];
        var subject = "Cohort Deployment Log - " + g.group;
        var html = srTlSummaryTemplate(g.group, g.rows, g.total);
        var plain = "Dear Sr. Team Leader,\n\nCohort deployment log for " + g.group + " (" + g.total + " interns).\n\nBest Regards,\nHR Operations Team";
        MailApp.sendEmail({ to: g.email, subject: subject, body: plain, htmlBody: html });
        sent++;
      }
      for (var j = 0; j < (rowNums || []).length; j++) {
        var rj = Number(rowNums[j]);
        var rdj = masterSheet.getRange(rj, 1, 1, masterSheet.getLastColumn()).getValues()[0];
        var infj = resolveGroupInfo(rdj, projData);
        if (infj.tlEmail) { setCol(masterSheet, rj, "TL Mail Sent", "Yes"); setCol(masterSheet, rj, "TL Mail Timestamp", ts); }
      }
    } else {
      for (var k = 0; k < (rowNums || []).length; k++) {
        var rk = Number(rowNums[k]);
        var rowD = masterSheet.getRange(rk, 1, 1, masterSheet.getLastColumn()).getValues()[0];
        var internEmail = valAt(rowD, masterSheet, "Email ID");
        if (!internEmail) continue;
        var info = resolveGroupInfo(rowD, projData);
        var gi = { groupName: info.groupName, email: info.tlEmail, whatsapp: info.whatsappLink };
        var internNameV = valAt(rowD, masterSheet, "Name");
        var internCodeV = valAt(rowD, masterSheet, "Intern Code");
        var iSubject = "Welcome Aboard! Your Internship Onboarding Details";
        var iHtml = internWelcomeTemplate({ name: internNameV, code: internCodeV }, gi);
        var iPlain = "Hello " + internNameV + ",\n\nYour onboarding is verified.\nIntern Code: " + internCodeV
          + "\nGroup: " + info.groupName + "\nTeam Community: " + info.whatsappLink + "\n\nBest Regards,\nHR Operations Team";
        MailApp.sendEmail({ to: internEmail, subject: iSubject, body: iPlain, htmlBody: iHtml });
        setCol(masterSheet, rk, "Intern Mail Sent", "Yes");
        setCol(masterSheet, rk, "Intern Mail Timestamp", ts);
        sent++;
      }
    }
    clearDataCache();
    return { success: true, sent: sent };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function uniq(arr) {
  var seen = {};
  var result = [];
  for (var i = 0; i < arr.length; i++) {
    var lower = arr[i].toLowerCase();
    if (seen[lower]) continue;
    seen[lower] = true;
    result.push(arr[i]);
  }
  return result;
}
