/*****************************************************************
 * UptoSkills Onboarding Automation
 * File : PdfParser.gs
 * Version : 1.9
 *****************************************************************/

function extractTextFromPdf(fileId) {
  // Try 1: copy to a Google Doc with OCR enabled (handles text + image PDFs).
  try {
    var fileMetadata = Drive.Files.get(fileId);

    // NATIVE GOOGLE DOC BYPASS:
    if (fileMetadata.mimeType === "application/vnd.google-apps.document") {
      Logger.log("Document is already a native Google Doc. Executing direct text pull.");
      return DocumentApp.openById(fileId).getBody().getText();
    }

    var resource = {
      title: "OCR_Temp_" + fileMetadata.title,
      mimeType: "application/vnd.google-apps.document"
    };
    var tempDoc = Drive.Files.copy(resource, fileId, { ocr: true, ocrLanguage: "en" });
    var rawText = DocumentApp.openById(tempDoc.id).getBody().getText();
    Drive.Files.remove(tempDoc.id);
    if (rawText && rawText.length) {
      Logger.log("extractTextFromPdf: copy method extracted " + rawText.length + " chars");
      return rawText;
    }
  } catch (e) {
    Logger.log("PDF OCR copy method failed: " + e.toString());
  }

  // Try 2: legacy insert-via-blob convert path (reliable for text-based PDFs).
  try {
    var file = DriveApp.getFileById(fileId);
    var tempDoc2 = Drive.Files.insert(
      { title: "OCR_Temp2_" + file.getName() },
      file.getBlob(),
      { convert: true }
    );
    var rawText2 = DocumentApp.openById(tempDoc2.id).getBody().getText();
    Drive.Files.remove(tempDoc2.id);
    if (rawText2 && rawText2.length) {
      Logger.log("extractTextFromPdf: insert method extracted " + rawText2.length + " chars");
      return rawText2;
    }
  } catch (e2) {
    Logger.log("PDF OCR insert method failed: " + e2.toString());
  }

  Logger.log("PDF OCR Critical Engine Failure Log: no text extracted");
  return "";
}

function extractDatesFromText(text) {
  var results = [];
  var seen = {};
  var monthsShort = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
  var monthsLong = { january:0, february:1, march:2, april:3, may:4, june:5, july:6, august:7, september:8, october:9, november:10, december:11 };

  function pushDate(d) {
    if (Object.prototype.toString.call(d) !== "[object Date]" || isNaN(d.getTime())) return;
    if (seen[d.getTime()]) return;
    seen[d.getTime()] = true;
    results.push(d);
  }
  function monthIndex(name) {
    name = (name || "").toLowerCase();
    if (monthsShort[name] !== undefined) return monthsShort[name];
    if (monthsLong[name] !== undefined) return monthsLong[name];
    return -1;
  }

  var m;

  // 1) ISO yyyy-MM-dd / yyyy/MM/dd
  var reIso = /\b(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\b/g;
  while ((m = reIso.exec(text)) !== null) {
    pushDate(new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  }

  // 2) dd Mon/Month yyyy  OR  Month dd, yyyy
  var reNamed = /\b(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})\b|\b([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})\b/g;
  while ((m = reNamed.exec(text)) !== null) {
    if (m[1]) {
      var mo = monthIndex(m[2]);
      if (mo !== -1) pushDate(new Date(Number(m[3]), mo, Number(m[1])));
    } else {
      var mo2 = monthIndex(m[4]);
      if (mo2 !== -1) pushDate(new Date(Number(m[6]), mo2, Number(m[5])));
    }
  }

  // 3) dd/mm/yyyy or dd-mm-yyyy (day first)
  var reDmy = /\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\b/g;
  while ((m = reDmy.exec(text)) !== null) {
    var day = Number(m[1]), mon = Number(m[2]) - 1, yr = Number(m[3]);
    if (mon <= 11 && day <= 31) pushDate(new Date(yr, mon, day));
  }

  // 4) dd/mm/yy or dd-mm-yy (2-digit year)
  var reDmy2 = /\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})\b/g;
  while ((m = reDmy2.exec(text)) !== null) {
    var day2 = Number(m[1]), mon2 = Number(m[2]) - 1, yr2 = Number(m[3]) + 2000;
    if (mon2 <= 11 && day2 <= 31) pushDate(new Date(yr2, mon2, day2));
  }

  Logger.log("extractDatesFromText: found " + results.length + " date(s)");
  return results;
}
