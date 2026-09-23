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
        "Receipt 1 Screenshot Filename",
        "Receipt 1 Screenshot File (Drive Link)",
        "Receipt 2 Screenshot Filename",
        "Receipt 2 Screenshot File (Drive Link)"
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

    // Save Screenshot 1 to Google Drive if image data is passed
    var fileUrl1 = "Not uploaded";
    var receipt1Filename = data.receiptFilename || data.receipt1Filename || "None";
    var receipt1Base64 = data.receiptBase64 || data.receipt1Base64;
    if (receipt1Base64 && receipt1Filename && receipt1Filename !== "None") {
      try {
        var contentType1 = receipt1Base64.substring(5, receipt1Base64.indexOf(';'));
        var bytes1 = Utilities.base64Decode(receipt1Base64.split(',')[1]);
        var blob1 = Utilities.newBlob(bytes1, contentType1, (data.regId || "REC") + "_Screenshot1_" + receipt1Filename);
        var file1 = DriveApp.createFile(blob1);
        file1.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl1 = file1.getUrl();
      } catch (driveErr1) {
        fileUrl1 = "Saved with error: " + driveErr1.toString();
      }
    }

    // Save Screenshot 2 to Google Drive if image data is passed
    var fileUrl2 = "Not uploaded";
    var receipt2Filename = data.receipt2Filename || "None";
    var receipt2Base64 = data.receipt2Base64;
    if (receipt2Base64 && receipt2Filename && receipt2Filename !== "None") {
      try {
        var contentType2 = receipt2Base64.substring(5, receipt2Base64.indexOf(';'));
        var bytes2 = Utilities.base64Decode(receipt2Base64.split(',')[1]);
        var blob2 = Utilities.newBlob(bytes2, contentType2, (data.regId || "REC") + "_Screenshot2_" + receipt2Filename);
        var file2 = DriveApp.createFile(blob2);
        file2.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl2 = file2.getUrl();
      } catch (driveErr2) {
        fileUrl2 = "Saved with error: " + driveErr2.toString();
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
      receipt1Filename,
      fileUrl1,
      receipt2Filename,
      fileUrl2
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
