/**
 * ==============================================================================
 * CREATOR CRAFT MEDIA — WORKSHOP REGISTRATION SPREADSHEET BACKEND
 * WITH AUTOMATIC VIP DELEGATE PASS PDF GENERATOR & EXCEL DOWNLOAD LINKS
 * ==============================================================================
 * 
 * INSTRUCTIONS TO GENERATE ALL PDFS RIGHT NOW:
 * 1. In the Apps Script toolbar at the top, click the dropdown next to "Debug"
 *    where it currently says "doPost" (or "onOpen").
 * 2. SELECT: "generateAllPassPdfs"
 * 3. Click "▷ Run".
 * 4. Google will immediately show: "Authorization required".
 *    - Click "Review permissions" -> Select your Google Account
 *    - Click "Advanced" (small link at bottom) -> "Go to Untitled project (unsafe)"
 *    - Click "Allow"
 * 5. It will generate PDFs for all attendees in your sheet and add the
 *    "Official Pass PDF (Drive Link)" column with clickable download links!
 * ==============================================================================
 */

/**
 * [FUNCTION 1] RUN THIS FUNCTION TO GENERATE ALL PASS PDFS & POPULATE EXCEL:
 * Selected by default at the top of the file.
 */
function generateAllPassPdfs() {
  // Explicitly touch DriveApp & SpreadsheetApp so Google triggers the authorization dialog
  var testDrive = DriveApp.getRootFolder();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Ensure the PDF column exists in your sheet
  var pdfCol = ensurePassPdfColumnExists(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    try {
      SpreadsheetApp.getUi().alert("No participant records found in this sheet.");
    } catch (e) {
      Logger.log("No participant records found in this sheet.");
    }
    return;
  }

  var dataRange = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
  var allData = dataRange.getValues();
  var headers = allData[0];

  // Map header column indices dynamically
  var cols = mapHeaderIndices(headers);

  var count = 0;
  for (var r = 1; r < allData.length; r++) {
    var row = allData[r];
    var currentPdfVal = String(row[pdfCol - 1] || '');

    // Generate if missing or not an active URL
    if (!currentPdfVal || currentPdfVal.indexOf('http') === -1) {
      var regId = (cols.regId >= 0 && row[cols.regId]) ? String(row[cols.regId]).trim() : ('CCM-TPT-' + (10000 + r));
      var fullName = (cols.fullName >= 0 && row[cols.fullName]) ? String(row[cols.fullName]).trim() : 'Participant';
      var mobile = (cols.mobile >= 0 && row[cols.mobile]) ? String(row[cols.mobile]).trim() : '-';
      var email = (cols.email >= 0 && row[cols.email]) ? String(row[cols.email]).trim() : '-';
      var category = (cols.category >= 0 && row[cols.category]) ? String(row[cols.category]).trim() : 'Content Creator';
      var location = (cols.location >= 0 && row[cols.location]) ? String(row[cols.location]).trim() : 'Tirupati';
      var utr = (cols.utr >= 0 && row[cols.utr]) ? String(row[cols.utr]).trim() : 'Payment Completed';
      var status = (cols.status >= 0 && row[cols.status]) ? String(row[cols.status]).trim() : 'Confirmed - Payment Submitted';

      var pdfUrl = generateAndSavePassPdf({
        regId: regId,
        fullName: fullName,
        mobile: mobile,
        email: email,
        category: category,
        location: location,
        status: status,
        utr: utr,
        fee: 'Rs. 2,999 (Lunch & Snacks Included)'
      });

      if (pdfUrl && pdfUrl.indexOf('http') !== -1) {
        var cell = sheet.getRange(r + 1, pdfCol);
        cell.setFormula('=HYPERLINK("' + pdfUrl + '", "📄 Download Pass PDF")');
        count++;
      }
    }
  }

  try {
    SpreadsheetApp.getUi().alert('Success! Generated VIP Delegate Pass PDFs for ' + count + ' registration(s) and added download links to your sheet!');
  } catch (uiErr) {
    Logger.log('Generated PDFs for ' + count + ' participants.');
  }
}

/**
 * Custom Menu inside Google Sheets:
 * Appears automatically when you open your spreadsheet.
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('🎟️ CCM Workshop')
      .addItem('📄 Generate Pass PDFs for All Existing Rows', 'generateAllPassPdfs')
      .addItem('🔄 Setup / Fix Sheet Columns', 'ensurePassPdfColumnExists')
      .addToUi();
  } catch (e) {}
}

/**
 * Helper: Ensures "Official Pass PDF (Drive Link)" column exists in your sheet.
 * If missing, it adds it right next to UTR / Transaction ID (or before Screenshots).
 */
function ensurePassPdfColumnExists(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  var headers = headerRange.getValues()[0];

  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase();
    if (h.indexOf('pass pdf') !== -1 || (h.indexOf('pass') !== -1 && h.indexOf('pdf') !== -1) || h.indexOf('official pass') !== -1) {
      return i + 1; // 1-based column index
    }
  }

  // Not found: find best position to insert (after UTR or before Screenshots)
  var insertAt = lastCol + 1;
  for (var j = 0; j < headers.length; j++) {
    var hj = String(headers[j]).toLowerCase();
    if (hj.indexOf('receipt') !== -1 || hj.indexOf('screenshot') !== -1) {
      insertAt = j + 1;
      break;
    }
  }

  sheet.insertColumnBefore(insertAt);
  var newCell = sheet.getRange(1, insertAt);
  newCell.setValue("Official Pass PDF (Drive Link)");
  newCell.setBackground("#0284C7");
  newCell.setFontColor("#FFFFFF");
  newCell.setFontWeight("bold");

  return insertAt;
}

/**
 * Handle incoming registration submissions from register.html
 */
function doPost(e) {
  // If run manually from Apps Script editor with Run button
  if (!e || !e.postData) {
    return generateAllPassPdfs();
  }

  var lock = LockService.getScriptLock();
  lock.tryLock(15000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Auto-create headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      var defaultHeaders = [
        "Timestamp",
        "Registration ID",
        "Status",
        "Full Name",
        "Mobile",
        "Email",
        "Category / Role",
        "Knowledge Level",
        "Location / City",
        "Freelance Interest",
        "Workshop Fee",
        "UTR / Transaction ID",
        "Official Pass PDF (Drive Link)",
        "Receipt 1 Filename",
        "Receipt 1 File (Drive Link)",
        "Receipt 2 Filename",
        "Receipt 2 File (Drive Link)"
      ];
      sheet.appendRow(defaultHeaders);
      var headerRange = sheet.getRange(1, 1, 1, defaultHeaders.length);
      headerRange.setBackground("#0F172A");
      headerRange.setFontColor("#F8FAFC");
      headerRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // Ensure PDF column exists in existing sheet
    ensurePassPdfColumnExists(sheet);

    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = e.parameter || {};
    }

    var regId = data.regId || ("CCM-TPT-" + Math.floor(10000 + Math.random() * 90000));
    var isConfirmed = data.status && data.status.indexOf("Confirmed") !== -1;
    var status = data.status || (isConfirmed ? "Confirmed - Payment Submitted" : "Step 1 - Payment Pending");

    // 1. Generate VIP Delegate Pass PDF & save to Google Drive
    var pdfDriveUrl = "Pending Payment";
    try {
      pdfDriveUrl = generateAndSavePassPdf({
        regId: regId,
        fullName: data.fullName || "Participant",
        mobile: data.mobile || "-",
        email: data.email || "-",
        category: data.category || "Creator / Marketer",
        knowledge: data.knowledge || "Beginner",
        location: data.location || "Tirupati",
        fee: data.fee || "Rs. 2,999 (Lunch & Snacks Included)",
        status: status,
        utr: data.utr || (isConfirmed ? "Payment Completed" : "Pending Verification")
      });
    } catch (pdfErr) {
      pdfDriveUrl = "PDF Error: " + pdfErr.toString();
    }

    // 2. Save Screenshot 1 to Google Drive if image data is passed
    var fileUrl1 = "Not uploaded";
    var receipt1Filename = data.receiptFilename || data.receipt1Filename || "None";
    var receipt1Base64 = data.receiptBase64 || data.receipt1Base64;
    if (receipt1Base64 && receipt1Filename && receipt1Filename !== "None" && receipt1Filename !== "Pending") {
      try {
        var contentType1 = receipt1Base64.substring(5, receipt1Base64.indexOf(";"));
        var bytes1 = Utilities.base64Decode(receipt1Base64.split(",")[1]);
        var blob1 = Utilities.newBlob(bytes1, contentType1, regId + "_Screenshot1_" + receipt1Filename);
        var file1 = DriveApp.createFile(blob1);
        file1.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl1 = file1.getUrl();
      } catch (driveErr1) {
        fileUrl1 = "Saved with error: " + driveErr1.toString();
      }
    }

    // 3. Save Screenshot 2 to Google Drive if image data is passed
    var fileUrl2 = "Not uploaded";
    var receipt2Filename = data.receipt2Filename || "None";
    var receipt2Base64 = data.receipt2Base64;
    if (receipt2Base64 && receipt2Filename && receipt2Filename !== "None" && receipt2Filename !== "Pending") {
      try {
        var contentType2 = receipt2Base64.substring(5, receipt2Base64.indexOf(";"));
        var bytes2 = Utilities.base64Decode(receipt2Base64.split(",")[1]);
        var blob2 = Utilities.newBlob(bytes2, contentType2, regId + "_Screenshot2_" + receipt2Filename);
        var file2 = DriveApp.createFile(blob2);
        file2.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl2 = file2.getUrl();
      } catch (driveErr2) {
        fileUrl2 = "Saved with error: " + driveErr2.toString();
      }
    }

    // 4. Read current header row to map fields to the EXACT columns in user's sheet
    var lastCol = sheet.getLastColumn();
    var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var pdfCellFormula = (pdfDriveUrl.indexOf('http') !== -1) ? ('=HYPERLINK("' + pdfDriveUrl + '", "📄 Download Pass PDF")') : pdfDriveUrl;

    var row = currentHeaders.map(function (h) {
      var low = String(h).toLowerCase();
      if (low.indexOf('time') !== -1 || low.indexOf('date') !== -1) return new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });
      if (low.indexOf('reg') !== -1 && low.indexOf('id') !== -1) return regId;
      if (low.indexOf('status') !== -1) return status;
      if (low.indexOf('pass') !== -1 || low.indexOf('pdf') !== -1) return pdfCellFormula;
      
      // Screenshots & receipts (check before generic "name" to avoid matching Filename as Name)
      var isReceipt = low.indexOf('receipt') !== -1 || low.indexOf('screenshot') !== -1;
      var isSecond = low.indexOf('2') !== -1;
      if (isReceipt && !isSecond && low.indexOf('filename') !== -1) return receipt1Filename;
      if (isReceipt && !isSecond && (low.indexOf('link') !== -1 || low.indexOf('drive') !== -1 || low.indexOf('file') !== -1 || low.indexOf('url') !== -1)) return fileUrl1;
      if (isReceipt && isSecond && low.indexOf('filename') !== -1) return receipt2Filename;
      if (isReceipt && isSecond && (low.indexOf('link') !== -1 || low.indexOf('drive') !== -1 || low.indexOf('file') !== -1 || low.indexOf('url') !== -1)) return fileUrl2;

      if (low.indexOf('mobile') !== -1 || low.indexOf('phone') !== -1) return data.mobile || "";
      if (low.indexOf('email') !== -1) return data.email || "";
      if (low.indexOf('category') !== -1 || low.indexOf('role') !== -1) return data.category || "";
      if (low.indexOf('knowledge') !== -1 || low.indexOf('level') !== -1) return data.knowledge || "";
      if (low.indexOf('location') !== -1 || low.indexOf('city') !== -1) return data.location || "";
      if (low.indexOf('freelance') !== -1) return data.freelance || "";
      if (low.indexOf('fee') !== -1) return data.fee || "Rs. 2,999 (Lunch & Snacks Included)";
      if (low.indexOf('utr') !== -1 || low.indexOf('trans') !== -1) return data.utr || "";
      if (low.indexOf('name') !== -1) return data.fullName || "";

      // Leave any custom columns (like T-Shirt) empty without corrupting row alignment
      return "";
    });

    // 5. Intelligent Duplicate Prevention:
    // If participant already completed Step 1 and now submits Step 2, update their row!
    var lastRow = sheet.getLastRow();
    var existingRow = -1;

    if (lastRow > 1) {
      var allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      var colIndices = mapHeaderIndices(currentHeaders);

      for (var i = 0; i < allData.length; i++) {
        var rowRegId = colIndices.regId >= 0 ? String(allData[i][colIndices.regId]).trim() : "";
        var rowMobile = colIndices.mobile >= 0 ? String(allData[i][colIndices.mobile]).replace(/\D/g, "") : "";
        var cleanMobile = String(data.mobile || "").replace(/\D/g, "");

        if ((data.regId && rowRegId === data.regId) ||
            (cleanMobile.length >= 10 && rowMobile.length >= 10 && rowMobile.slice(-10) === cleanMobile.slice(-10))) {
          existingRow = i + 2;
          break;
        }
      }
    }

    if (existingRow > 1) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "success",
        regId: regId,
        pdfUrl: pdfDriveUrl,
        message: "Registration and VIP Delegate Pass recorded successfully in Google Sheet"
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Generate PDF document and save to Google Drive with public view access
 */
function generateAndSavePassPdf(data) {
  var passHtml = createPassHtmlDocument(data);
  var htmlOutput = HtmlService.createHtmlOutput(passHtml);
  var pdfBlob = htmlOutput.getAs("application/pdf");
  pdfBlob.setName((data.regId || "CCM-TPT") + "_VIP_Delegate_Pass.pdf");

  var pdfFile = DriveApp.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return pdfFile.getUrl();
}

/**
 * Map header names to column array indices dynamically
 */
function mapHeaderIndices(headers) {
  var cols = {
    regId: -1,
    fullName: -1,
    mobile: -1,
    email: -1,
    category: -1,
    knowledge: -1,
    location: -1,
    status: -1,
    utr: -1,
    pdf: -1
  };

  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).toLowerCase();
    if (h.indexOf('reg') !== -1 && h.indexOf('id') !== -1) cols.regId = c;
    else if (h.indexOf('mobile') !== -1 || h.indexOf('phone') !== -1) cols.mobile = c;
    else if (h.indexOf('email') !== -1) cols.email = c;
    else if (h.indexOf('category') !== -1 || h.indexOf('role') !== -1) cols.category = c;
    else if (h.indexOf('knowledge') !== -1 || h.indexOf('level') !== -1) cols.knowledge = c;
    else if (h.indexOf('location') !== -1 || h.indexOf('city') !== -1) cols.location = c;
    else if (h.indexOf('status') !== -1) cols.status = c;
    else if (h.indexOf('utr') !== -1 || h.indexOf('trans') !== -1) cols.utr = c;
    else if (h.indexOf('pass') !== -1 || h.indexOf('pdf') !== -1) cols.pdf = c;
    else if (h.indexOf('name') !== -1 && h.indexOf('file') === -1 && h.indexOf('receipt') === -1) cols.fullName = c;
  }
  return cols;
}

/**
 * GET Endpoint for browser health check & direct pass download
 */
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.regId && (e.parameter.download === "pdf" || e.parameter.view === "pdf")) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var cols = mapHeaderIndices(headers);

        for (var i = 0; i < data.length; i++) {
          var rowRegId = cols.regId >= 0 ? String(data[i][cols.regId]).trim() : "";
          if (rowRegId === e.parameter.regId.trim()) {
            var passHtml = createPassHtmlDocument({
              regId: rowRegId,
              fullName: cols.fullName >= 0 ? data[i][cols.fullName] : "Participant",
              mobile: cols.mobile >= 0 ? data[i][cols.mobile] : "-",
              email: cols.email >= 0 ? data[i][cols.email] : "-",
              category: cols.category >= 0 ? data[i][cols.category] : "Creator / Marketer",
              location: cols.location >= 0 ? data[i][cols.location] : "Tirupati",
              status: cols.status >= 0 ? data[i][cols.status] : "Confirmed",
              utr: cols.utr >= 0 ? data[i][cols.utr] : "Completed",
              fee: "Rs. 2,999 (Lunch & Snacks Included)"
            });
            var pdfBlob = HtmlService.createHtmlOutput(passHtml).getAs("application/pdf");
            pdfBlob.setName(e.parameter.regId.trim() + "_VIP_Delegate_Pass.pdf");
            return pdfBlob;
          }
        }
      }
    }
  } catch (err) {}

  return ContentService
    .createTextOutput(JSON.stringify({
      status: "active",
      service: "Creator Craft Media Workshop Registration & PDF API",
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * HTML Template for Official VIP Delegate Pass PDF
 */
function createPassHtmlDocument(data) {
  var regId = data.regId || "CCM-TPT-10036";
  var name = data.fullName || "Participant";
  var mobile = data.mobile || "-";
  var email = data.email || "-";
  var category = data.category || "Creator / Marketer";
  var location = data.location || "Tirupati";
  var fee = data.fee || "Rs. 2,999 (Lunch & Snacks Included)";
  var isConfirmed = data.status && data.status.indexOf("Confirmed") !== -1;
  var statusText = isConfirmed ? "Payment Submitted" : "Payment Pending";
  var utr = data.utr || (isConfirmed ? "Payment Completed (Screenshot Attached)" : "Pending Verification");

  return '<!DOCTYPE html>' +
'<html>' +
'<head>' +
'<meta charset="UTF-8">' +
'<title>Official VIP Delegate Pass - ' + regId + '</title>' +
'<style>' +
'  @page { size: A4 portrait; margin: 8mm; }' +
'  * { box-sizing: border-box; margin: 0; padding: 0; }' +
'  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; background: #FFFFFF; color: #0F172A; line-height: 1.4; padding: 10px; }' +
'  .pass-wrapper { border: 2px solid #0284C7; border-radius: 8px; overflow: hidden; max-width: 740px; margin: 0 auto; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }' +
'  .pass-header { background: #060911; color: #FFFFFF; padding: 18px 24px; border-bottom: 3px solid #0284C7; }' +
'  .header-table { width: 100%; border-collapse: collapse; }' +
'  .brand-title { font-size: 16px; font-weight: bold; color: #FFFFFF; letter-spacing: 0.05em; text-transform: uppercase; }' +
'  .brand-sub { font-size: 9px; color: #94A3B8; letter-spacing: 0.12em; text-transform: uppercase; margin-top: 3px; }' +
'  .header-right { text-align: right; }' +
'  .official-tag { font-size: 10px; font-weight: bold; color: #38BDF8; letter-spacing: 0.1em; text-transform: uppercase; }' +
'  .reg-code { font-family: Courier, monospace; font-size: 20px; font-weight: bold; color: #38BDF8; letter-spacing: 0.05em; margin: 2px 0; }' +
'  .event-date { font-size: 11px; color: #CBD5E1; }' +
'  .status-badge { display: inline-block; padding: 3px 12px; border-radius: 999px; font-size: 9.5px; font-weight: bold; text-transform: uppercase; margin-top: 5px; background: ' + (isConfirmed ? '#047857' : '#B45309') + '; color: #FFFFFF; }' +
'  .pass-body { padding: 18px 22px; background: #FFFFFF; }' +
'  .section-label { font-size: 11.5px; font-weight: bold; color: #0F172A; letter-spacing: 0.05em; text-transform: uppercase; margin-bottom: 8px; margin-top: 6px; }' +
'  .cred-table { width: 100%; border-collapse: collapse; border: 1px solid #CBD5E1; border-left: 6px solid #0284C7; border-radius: 6px; margin-bottom: 16px; font-size: 11.5px; }' +
'  .cred-table td { padding: 9px 12px; border-bottom: 1px solid #E2E8F0; }' +
'  .cred-table tr:last-child td { border-bottom: none; }' +
'  .cred-label { color: #64748B; font-weight: 600; width: 22%; }' +
'  .cred-val { color: #0F172A; font-weight: bold; width: 28%; }' +
'  .modules-table { width: 100%; border-collapse: separate; border-spacing: 10px 0; margin-bottom: 14px; }' +
'  .mod-card { border-radius: 6px; padding: 12px 14px; font-size: 10.5px; vertical-align: top; width: 50%; }' +
'  .mod-1 { background: #F0F9FF; border: 1px solid #BAE6FD; }' +
'  .mod-2 { background: #FAF5FF; border: 1px solid #E9D5FF; }' +
'  .mod-title-1 { font-size: 11px; font-weight: bold; color: #0284C7; text-transform: uppercase; margin-bottom: 6px; }' +
'  .mod-title-2 { font-size: 11px; font-weight: bold; color: #7E22CE; text-transform: uppercase; margin-bottom: 6px; }' +
'  .mod-item { margin-bottom: 3px; color: #334155; line-height: 1.35; }' +
'  .instructions-box { background: #FEFCE8; border: 1px solid #FEF08A; border-radius: 6px; padding: 10px 14px; margin-bottom: 14px; font-size: 10px; color: #451A03; }' +
'  .instructions-title { font-weight: bold; color: #854D0E; text-transform: uppercase; margin-bottom: 4px; font-size: 10.5px; }' +
'  .pass-footer { background: #060911; color: #FFFFFF; padding: 14px 22px; border-top: 3px solid #0284C7; }' +
'  .footer-table { width: 100%; border-collapse: collapse; font-size: 10.5px; }' +
'  .desk-title { font-size: 12px; font-weight: bold; color: #38BDF8; margin-bottom: 3px; }' +
'  .desk-sub { color: #94A3B8; line-height: 1.45; }' +
'  .verif-box { border: 2px solid #0284C7; border-radius: 6px; padding: 6px 12px; text-align: center; background: #0c1c2e; }' +
'  .verif-title { font-size: 11px; font-weight: bold; color: #38BDF8; }' +
'  .verif-code { font-family: Courier, monospace; font-size: 11px; font-weight: bold; color: #FFFFFF; margin-top: 2px; }' +
'</style>' +
'</head>' +
'<body>' +
'  <div class="pass-wrapper">' +
'    <div class="pass-header">' +
'      <table class="header-table">' +
'        <tr>' +
'          <td>' +
'            <div class="brand-title">CREATOR CRAFT MEDIA</div>' +
'            <div class="brand-sub">SHOOT &bull; CINEMATIC &bull; EDIT &bull; ICONIC</div>' +
'          </td>' +
'          <td class="header-right">' +
'            <div class="official-tag">OFFICIAL VIP DELEGATE PASS</div>' +
'            <div class="reg-code">' + regId + '</div>' +
'            <div class="event-date">Nov 1, 2026 &bull; Tirupati, AP</div>' +
'            <div><span class="status-badge">' + statusText + '</span></div>' +
'          </td>' +
'        </tr>' +
'      </table>' +
'    </div>' +
'    <div class="pass-body">' +
'      <div class="section-label">DELEGATE CREDENTIAL &amp; RESERVATION DETAILS</div>' +
'      <table class="cred-table">' +
'        <tr>' +
'          <td class="cred-label">Participant Name:</td><td class="cred-val">' + name + '</td>' +
'          <td class="cred-label">Mobile Number:</td><td class="cred-val">' + mobile + '</td>' +
'        </tr>' +
'        <tr style="background:#F8FAFC;">' +
'          <td class="cred-label">Email Address:</td><td class="cred-val">' + email + '</td>' +
'          <td class="cred-label">Category / Role:</td><td class="cred-val" style="color:#4F46E5;">' + category + '</td>' +
'        </tr>' +
'        <tr>' +
'          <td class="cred-label">Workshop Fee:</td><td class="cred-val">' + fee + '</td>' +
'          <td class="cred-label">Payment Status:</td><td class="cred-val" style="color:' + (isConfirmed ? '#047857' : '#B45309') + ';">' + statusText + '</td>' +
'        </tr>' +
'        <tr style="background:#F8FAFC;">' +
'          <td class="cred-label">UTR Reference:</td><td class="cred-val">' + utr + '</td>' +
'          <td class="cred-label">Location / City:</td><td class="cred-val">' + location + '</td>' +
'        </tr>' +
'      </table>' +
'      <div class="section-label">OFFICIAL WORKSHOP PRACTICAL MODULES</div>' +
'      <table class="modules-table">' +
'        <tr>' +
'          <td class="mod-card mod-1">' +
'            <div class="mod-title-1">DAY 01 &bull; NOV 01: META ADS LIVE WORKSHOP</div>' +
'            <div class="mod-item">&bull; High-ROAS Facebook &amp; Instagram Ad Architecture</div>' +
'            <div class="mod-item">&bull; Laser Audience Targeting &amp; Lookalike Segments</div>' +
'            <div class="mod-item">&bull; Winning Ad Creatives, Angles &amp; Copy Strategy</div>' +
'            <div class="mod-item">&bull; Budget Optimization, Scaling &amp; Tracking</div>' +
'            <div class="mod-item">&bull; Sound Design, Trending Audio &amp; Viral Pacing</div>' +
'          </td>' +
'          <td class="mod-card mod-2">' +
'            <div class="mod-title-2">DAY 02 &bull; NOV 01: AI FILMS &amp; MONETIZATION</div>' +
'            <div class="mod-item">&bull; Generative AI Video, Scripting &amp; Storyboards</div>' +
'            <div class="mod-item">&bull; Google Flow, Higgsfield, Seedance &amp; Kling</div>' +
'            <div class="mod-item">&bull; AI Voice Cloning &amp; Lip-Sync Filmmaking</div>' +
'            <div class="mod-item">&bull; CapCut Pro Video Editing Masterclass</div>' +
'            <div class="mod-item">&bull; Client Acquisition, Pricing &amp; Blueprint</div>' +
'          </td>' +
'        </tr>' +
'      </table>' +
'      <div class="instructions-box">' +
'        <div class="instructions-title">IMPORTANT WORKSHOP INSTRUCTIONS FOR ATTENDEES</div>' +
'        <div>1. <strong>Reporting:</strong> Arrive at venue by 8:45 AM for registration kit collection.</div>' +
'        <div>2. <strong>Equipment:</strong> Bring smartphone with 15GB free storage &amp; power bank (Laptop optional).</div>' +
'        <div>3. <strong>Food &amp; Refreshments:</strong> Delicious lunch &amp; high-tea snacks provided on both days.</div>' +
'        <div>4. <strong>Certification:</strong> Official Creator Craft Media Workshop Certificate awarded at closing.</div>' +
'      </div>' +
'    </div>' +
'    <div class="pass-footer">' +
'      <table class="footer-table">' +
'        <tr>' +
'          <td>' +
'            <div class="desk-title">CREATOR CRAFT MEDIA &bull; TIRUPATI DESK</div>' +
'            <div class="desk-sub">Helpline / WhatsApp: +91 9121796816 | Instagram: @creator_craft_media<br>UPI Merchant ID: alphaassetfinserv@sbi (SBI Business Account)<br>Please present this verified pass or Registration ID at venue desk.</div>' +
'          </td>' +
'          <td style="width: 170px; text-align: right;">' +
'            <div class="verif-box">' +
'              <div class="verif-title">VERIFIED VIP PASS</div>' +
'              <div class="verif-code">' + regId + '</div>' +
'              <div style="font-size: 8.5px; color: #94A3B8; margin-top: 2px;">Valid for Nov 1, 2026</div>' +
'            </div>' +
'          </td>' +
'        </tr>' +
'      </table>' +
'    </div>' +
'  </div>' +
'</body>' +
'</html>';
}
