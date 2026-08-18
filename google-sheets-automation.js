/**
 * ==============================================================================
 * LOGIC 11+ GOOGLE APPS SCRIPT: CUSTOM EMAIL DRAFTS & SENDER
 * ==============================================================================
 * 
 * NEW FEATURES:
 * 1. Column H: "Send Further Email" -> Set to "Yes" or "Draft"
 *    - If you type "Draft": Creates a pre-filled Gmail Draft in your logic11plus@gmail.com 
 *      inbox so you can easily review, tweak, and send it with one click!
 *    - If you type "Yes": Sends the customized email directly.
 * 2. Column G: "Custom Email Note" -> Any special note you type here will automatically
 *    appear in that parent's email (e.g. "We have placed Leo into the Tuesday 5pm cohort").
 * 
 * GOOGLE SHEET COLUMN HEADERS (Row 1):
 * Col A: Timestamp
 * Col B: Parent Name
 * Col C: Child Name
 * Col D: Child School
 * Col E: Parent Email
 * Col F: Target Year
 * Col G: Custom Email Note / Class Time
 * Col H: Send Email (Yes / Draft / No)
 * Col I: Email Status
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
    var customNote = data.notes || "";
    var sendFurtherEmail = "No"; // Defaults to No
    var emailStatus = "Pending Review";

    sheet.appendRow([
      timestamp,
      parentName,
      childName,
      childSchool,
      parentEmail,
      targetYear,
      customNote,
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

function doGet(e) {
  return ContentService.createTextOutput("Logic 11+ Google Sheet Webhook is active!");
}

// --- 2. ONE-TIME PERMISSION AUTHORIZER FUNCTION ---
// If drafts or emails fail with "Exception: Action not allowed", select this function in Apps Script and click "Run" once to grant Gmail permissions.
function authorizeGmailPermissions() {
  var testDraft = GmailApp.createDraft("logic11plus@gmail.com", "Permissions Test", "Testing GmailApp permissions.");
  testDraft.deleteDraft();
  Logger.log("GmailApp authorized successfully!");
}

// --- 3. AUTOMATED EMAIL / DRAFT CREATOR WHEN YOU TYPE 'YES' OR 'DRAFT' ---
function onEditTrigger(e) {
  if (!e || !e.source || !e.range) {
    Logger.log("onEditTrigger must be triggered by editing the spreadsheet.");
    return;
  }

  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  // Column H is Column 8 ("Send Further Email")
  if (row > 1 && col === 8) {
    var rawVal = range.getValue();
    if (!rawVal) return;
    var val = rawVal.toString().trim().toLowerCase();
    
    if (val === "yes" || val === "draft") {
      var emailStatusCell = sheet.getRange(row, 9); // Column I
      var currentStatus = emailStatusCell.getValue().toString();
      
      // Prevent duplicate sends
      if (currentStatus.indexOf("Sent") !== -1 && val === "yes") {
        SpreadsheetApp.getUi().alert("Email has already been sent for this lead!");
        return;
      }
      
      var parentName = sheet.getRange(row, 2).getValue();
      var childName = sheet.getRange(row, 3).getValue();
      var childSchool = sheet.getRange(row, 4).getValue();
      var parentEmail = sheet.getRange(row, 5).getValue();
      var targetYear = sheet.getRange(row, 6).getValue();
      var customNote = sheet.getRange(row, 7).getValue(); // Column G custom note
      
      if (!parentEmail || parentEmail.indexOf("@") === -1) {
        SpreadsheetApp.getUi().alert("Invalid or missing email address in Column E (Row " + row + ").");
        emailStatusCell.setValue("Error: Missing Email");
        return;
      }
      
      var subject = "Logic 11+ Mathematics Tuition: Enrollment Details for " + (childName || "Your Child");
      
      var customNoteSection = "";
      if (customNote && customNote.toString().trim() !== "" && customNote.toString().trim() !== "None") {
        customNoteSection = "\nSpecial Note for Your Placement:\n" + customNote + "\n";
      }

      var body = "Dear " + (parentName || "Parent") + ",\n\n" +
        "Thank you for registering " + (childName || "your child") + (childSchool ? " (" + childSchool + ")" : "") + " for Logic 11+ Mathematics Tuition.\n\n" +
        "We are pleased to confirm your place in our online group tuition cohort (" + (targetYear || "Year 4/5") + ").\n\n" +
        "Key Course Details:\n" +
        "• Format: 100% Online Interactive Live Group Classroom\n" +
        "• Focus: Pure 11+ Mathematics & Reasoning Mastery\n" +
        "• Contact: logic11plus@gmail.com\n" +
        customNoteSection + "\n" +
        "Your upcoming session schedule and digital classroom link will be shared in our next follow-up.\n\n" +
        "If you have any questions, simply reply directly to this email.\n\n" +
        "Warm regards,\n" +
        "The Logic 11+ Team\n" +
        "logic11plus@gmail.com";
      
      try {
        if (val === "draft") {
          // Creates an editable draft in your Gmail Drafts folder!
          GmailApp.createDraft(parentEmail, subject, body);
          emailStatusCell.setValue("Draft Created in Gmail (" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm") + ")");
        } else {
          // Sends email directly
          MailApp.sendEmail({
            to: parentEmail,
            subject: subject,
            body: body,
            replyTo: "logic11plus@gmail.com",
            name: "Logic 11+ Tuition"
          });
          
          var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
          emailStatusCell.setValue("Sent on " + timestamp);
        }
      } catch (err) {
        emailStatusCell.setValue("Error: " + err.message);
      }
    }
  }
}
