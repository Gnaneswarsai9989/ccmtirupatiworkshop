/**
 * ==============================================================================
 * CREATOR CRAFT MEDIA — WORKSHOP REGISTRATION SPREADSHEET BACKEND
 * ==============================================================================
 * 
 * HOW TO SETUP IN 2 MINUTES:
 * 1. Open Google Sheets (https://sheets.google.com) and create a New Spreadsheet.
 *    Name it: "CCM Tirupati Workshop Registrations 2026".
 * 
 * 2. In Google Sheets menu, click:
 *    Extensions > Apps Script
 * 
 * 3. Delete any existing code in the editor, and PASTE THIS ENTIRE FILE.
 * 
 * 4. Click "Deploy" (blue button at top right) > "New deployment".
 *    - Click the Gear icon next to "Select type" and select: "Web app"
 *    - Description: "Workshop Registration API"
 *    - Execute as: "Me" (your email)
 *    - Who has access: "Anyone"  <--- (IMPORTANT! Must select "Anyone")
 * 
 * 5. Click "Deploy", approve permissions if prompted, and COPY the "Web app URL"
 *    (It looks like: https://script.google.com/macros/s/AKfycb.../exec).
 * 
 * 6. Deployed Web App URL:
 *    https://script.google.com/macros/s/AKfycbxnpKsBm2SU4DWR3KoD5DP1NTpyNcJRKoXpUavF4xshLOOG53AlTDCNxTNGFr-MH6ihfQ/exec
 * 
 * THAT'S IT! All submissions, UTR numbers, and receipt details will instantly
 * populate your Google Sheet in real-time!
 * ==============================================================================
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Auto-create headers if the sheet is empty
    if (sheet.getLastRow() === 0) {
      var headers = [
        "Timestamp",
        "Registration ID",
        "Full Name",
        "Mobile",
        "Email",
        "T-Shirt",
        "Category",
        "Knowledge Level",
        "Location",
        "Freelancer Interest",
        "UTR / Transaction ID",
        "Receipt Screenshot Filename",
        "Receipt Screenshot File (Drive Link)"
      ];
      sheet.appendRow(headers);
      
      // Format header row (Dark luxury style)
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#0F172A");
      headerRange.setFontColor("#F8FAFC");
      headerRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = e.parameter;
    }

    // Save receipt screenshot to Google Drive if image data is passed
    var fileUrl = "Not uploaded";
    if (data.receiptBase64 && data.receiptFilename) {
      try {
        var contentType = data.receiptBase64.substring(5, data.receiptBase64.indexOf(';'));
        var bytes = Utilities.base64Decode(data.receiptBase64.split(',')[1]);
        var blob = Utilities.newBlob(bytes, contentType, (data.regId || "REC") + "_" + data.receiptFilename);
        var file = DriveApp.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = file.getUrl();
      } catch (driveErr) {
        fileUrl = "Saved with error: " + driveErr.toString();
      }
    }

    var row = [
      new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }),
      data.regId || "CCM-TRP-" + Math.floor(1000 + Math.random() * 9000),
      data.fullName || "",
      data.mobile || "",
      data.email || "",
      data.tshirt || "",
      data.category || "",
      data.knowledge || "",
      data.location || "",
      data.freelance || "",
      data.utr || "",
      data.receiptFilename || "None",
      fileUrl
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", regId: data.regId, message: "Registration recorded successfully" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "active", message: "CCM Workshop Registration API is live" }))
    .setMimeType(ContentService.MimeType.JSON);
}
