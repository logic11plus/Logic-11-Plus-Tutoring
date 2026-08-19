/**
 * ==============================================================================
 * LOGIC 11+ GOOGLE APPS SCRIPT: 100% RELIABLE DRAFTS & EMAILS
 * ==============================================================================
 * 
 * WHY THIS WORKS EVERY TIME:
 * Google blocks `GmailApp` inside simple edit triggers due to security restrictions.
 * Instead, this script adds a custom "🎓 Logic 11+" menu directly into your Google Sheet!
 * 
 * HOW TO USE IT:
 * 1. Open your Google Sheet. You will see a new menu at the top: "🎓 Logic 11+".
 * 2. Click on the row of the parent you want to email.
 * 3. Click "🎓 Logic 11+" -> "📝 Create Gmail Draft for Selected Row".
 * 4. Google will immediately create the draft in your logic11plus@gmail.com inbox!
 * 
 * You can also still type "Yes" or "Draft" in Column H if you install an Installable Trigger.
 */

// --- 1. ADD CUSTOM MENU TO GOOGLE SHEET TOP BAR ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🎓 Logic 11+')
    .addItem('📝 Create Gmail Draft for Selected Row', 'createDraftForActiveRow')
    .addItem('✉️ Send Email Directly for Selected Row', 'sendEmailForActiveRow')
    .addSeparator()
    .addItem('🔑 Authorize Permissions', 'authorizePermissions')
    .addToUi();
}

// --- 2. ONE-CLICK DRAFT CREATOR FOR THE ROW YOU HAVE CLICKED ON ---
function createDraftForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("Please click on a row with a parent's details (Row 2 or below).");
    return;
  }
  
  processRowAction(sheet, row, "draft");
}

// --- 3. ONE-CLICK EMAIL SENDER FOR THE ROW YOU HAVE CLICKED ON ---
function sendEmailForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("Please click on a row with a parent's details (Row 2 or below).");
    return;
  }
  
  processRowAction(sheet, row, "yes");
}

// --- 4. CORE ACTION PROCESSOR (CREATES DRAFT OR SENDS EMAIL) ---
function processRowAction(sheet, row, actionType) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var childSchool = sheet.getRange(row, 4).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var targetYear = sheet.getRange(row, 6).getValue();
  var customNote = sheet.getRange(row, 7).getValue(); // Column G custom note
  var emailStatusCell = sheet.getRange(row, 9); // Column I
  
  if (!parentEmail || parentEmail.toString().indexOf("@") === -1) {
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
    if (actionType === "draft") {
      GmailApp.createDraft(parentEmail, subject, body);
      emailStatusCell.setValue("Draft Created in Gmail (" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm") + ")");
      sheet.getRange(row, 8).setValue("Draft");
      SpreadsheetApp.getUi().alert("✅ Draft created successfully in Gmail Drafts for " + parentEmail + "!");
    } else {
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
      SpreadsheetApp.getUi().alert("✅ Email sent successfully to " + parentEmail + "!");
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    SpreadsheetApp.getUi().alert("❌ Error: " + err.message);
  }
}

// --- 5. RECEIVE FORM SUBMISSION FROM WEBSITE & INSERT INTO SHEET ---
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
    var sendFurtherEmail = "No";
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

// --- 6. TRIGGER FOR AUTOMATIC TYPING IN COLUMN H ---
function onEditTrigger(e) {
  if (!e || !e.source || !e.range) return;
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  if (row > 1 && col === 8) {
    var rawVal = range.getValue();
    if (!rawVal) return;
    var val = rawVal.toString().trim().toLowerCase();
    
    if (val === "yes") {
      processRowAction(sheet, row, "yes");
    } else if (val === "draft") {
      processRowAction(sheet, row, "draft");
    }
  }
}

// One-time authorization helper
function authorizePermissions() {
  var testDraft = GmailApp.createDraft("logic11plus@gmail.com", "Test", "Test");
  testDraft.deleteDraft();
  SpreadsheetApp.getUi().alert("✅ Permissions successfully authorized!");
}
