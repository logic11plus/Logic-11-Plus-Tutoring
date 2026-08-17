/**
 * ==============================================================================
 * LOGIC 11+ GOOGLE APPS SCRIPT: AUTO GOOGLE SHEET + AUTO EMAIL ON "YES"
 * ==============================================================================
 * 
 * STEP 1: CREATE THE GOOGLE SHEET
 * 1. Open Google Sheets (sheets.google.com) signed in as logic11plus@gmail.com
 * 2. In Row 1, add these exact column headers:
 *    Col A: Timestamp
 *    Col B: Parent Name
 *    Col C: Child Name
 *    Col D: Child School
 *    Col E: Parent Email
 *    Col F: Target Year
 *    Col G: Additional Notes
 *    Col H: Send Further Email
 *    Col I: Email Status
 * 
 * STEP 2: PASTE THIS CODE INTO APPS SCRIPT
 * 1. In your Google Sheet menu, click: Extensions > Apps Script
 * 2. Delete whatever is in the editor, and paste THIS ENTIRE FILE.
 * 3. Click the Save icon (floppy disk).
 * 
 * STEP 3: DEPLOY AS WEB APP (So your website can insert rows automatically)
 * 1. Click the blue "Deploy" button (top right) -> "New deployment".
 * 2. Select type: "Web app" (click the gear icon next to Select type).
 * 3. Set:
 *    - Description: Logic 11+ Form Submissions
 *    - Execute as: "Me (logic11plus@gmail.com)"
 *    - Who has access: "Anyone" (allows website form submissions without login)
 * 4. Click "Deploy" and authorize permissions.
 * 5. Copy the generated "Web App URL" (looks like https://script.google.com/macros/s/.../exec).
 * 6. Paste that URL into script.js in your website folder (line 17: GOOGLE_SHEET_WEBAPP_URL).
 * 
 * STEP 4: ENABLE AUTO EMAIL TRIGGER ON "YES"
 * 1. In Apps Script, click the Triggers icon (clock icon on the left sidebar).
 * 2. Click "+ Add Trigger" (bottom right):
 *    - Choose which function to run: onEditTrigger
 *    - Select event source: From spreadsheet
 *    - Select event type: On edit
 * 3. Click Save.
 * 
 * DONE! 
 * - When someone submits the form on the website -> Row is added to Google Sheet instantly.
 * - When you type "Yes" in Column H -> An email is automatically sent to the parent from logic11plus@gmail.com!
 */

// --- 1. RECEIVE FORM SUBMISSION FROM WEBSITE & INSERT INTO SHEET ---
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var parentName = data.parentName || "";
    var childName = data.childName || "";
    var childSchool = data.childSchool || "";
    var parentEmail = data.parentEmail || "";
    var targetYear = data.targetYear || "";
    var notes = data.notes || "";
    var sendFurtherEmail = "No"; // Defaults to No
    var emailStatus = "Pending Review";

    // Append new row
    sheet.appendRow([
      timestamp,
      parentName,
      childName,
      childSchool,
      parentEmail,
      targetYear,
      notes,
      sendFurtherEmail,
      emailStatus
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", row: sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Support GET for testing
function doGet(e) {
  return ContentService.createTextOutput("Logic 11+ Google Sheet Webhook is active!");
}

// --- 2. AUTOMATED EMAIL WHEN YOU TYPE "YES" IN COLUMN H ---
function onEditTrigger(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  // Column H is Column 8 ("Send Further Email")
  if (row > 1 && col === 8) {
    var val = range.getValue().toString().trim().toLowerCase();
    
    if (val === "yes") {
      var emailStatusCell = sheet.getRange(row, 9); // Column I
      var currentStatus = emailStatusCell.getValue().toString();
      
      // Prevent duplicate sends
      if (currentStatus.indexOf("Sent") !== -1) {
        SpreadsheetApp.getUi().alert("Email has already been sent for this lead!");
        return;
      }
      
      var parentName = sheet.getRange(row, 2).getValue();
      var childName = sheet.getRange(row, 3).getValue();
      var childSchool = sheet.getRange(row, 4).getValue();
      var parentEmail = sheet.getRange(row, 5).getValue();
      var targetYear = sheet.getRange(row, 6).getValue();
      
      if (!parentEmail || parentEmail.indexOf("@") === -1) {
        SpreadsheetApp.getUi().alert("Invalid or missing email address in Column E (Row " + row + ").");
        return;
      }
      
      var subject = "Logic 11+ Mathematics Tuition: Enrollment Confirmation for " + childName;
      
      var body = "Dear " + parentName + ",\n\n" +
        "Thank you for registering " + childName + " (" + childSchool + ") for Logic 11+ Mathematics Tuition.\n\n" +
        "We are pleased to confirm your place in our online group tuition cohort (" + targetYear + ").\n\n" +
        "Key Course Details:\n" +
        "• Format: 100% Online Interactive Live Group Classroom\n" +
        "• Focus: Pure 11+ Mathematics & Reasoning Mastery\n" +
        "• Contact: logic11plus@gmail.com\n\n" +
        "Your upcoming session schedule and digital classroom link will be shared in our next follow-up.\n\n" +
        "If you have any questions, simply reply directly to this email.\n\n" +
        "Warm regards,\n" +
        "The Logic 11+ Team\n" +
        "logic11plus@gmail.com";
      
      try {
        MailApp.sendEmail({
          to: parentEmail,
          subject: subject,
          body: body,
          replyTo: "logic11plus@gmail.com",
          name: "Logic 11+ Tuition"
        });
        
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        emailStatusCell.setValue("Sent on " + timestamp);
        sheet.getRange(row, 8).setValue("Yes");
      } catch (err) {
        emailStatusCell.setValue("Error: " + err.message);
      }
    }
  }
}
