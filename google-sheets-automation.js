/**
 * ==============================================================================
 * LOGIC 11+ GOOGLE APPS SCRIPT: 100% RELIABLE DRAFTS & EMAILS
 * ==============================================================================
 * 
 * FEATURES IN THE "🎓 Logic 11+" MENU (Next to Help):
 * 1. 📝 Create Confirmation Draft: Creates an enrollment draft with their selected slot.
 * 2. ✉️ Send Confirmation Email: Sends the enrollment email directly.
 * 3. ⏳ Send "No Spaces / Waitlist" Email: Informs parent class is full and they are on the priority waitlist.
 * 4. 📋 Create Waitlist Available Draft: Pre-fills an invitation draft with slot placeholders to fill in.
 * 5. 🔑 Authorize Permissions: One-click one-time setup.
 */

// --- 1. ADD CUSTOM MENU TO GOOGLE SHEET TOP BAR ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🎓 Logic 11+')
    .addItem('📝 Create Confirmation Draft for Selected Row', 'createConfirmationDraftForActiveRow')
    .addItem('✉️ Send Confirmation Email for Selected Row', 'sendConfirmationEmailForActiveRow')
    .addSeparator()
    .addItem('⏳ Send "No Spaces Available" Email for Selected Row', 'sendNoSpacesEmailForActiveRow')
    .addItem('📋 Create "Waitlist Space Available" Draft for Selected Row', 'createWaitlistAvailableDraftForActiveRow')
    .addSeparator()
    .addItem('✅ Mark Selected Row as Sent (Clear Yellow/Red Highlight)', 'markSentAndClearHighlightForActiveRow')
    .addSeparator()
    .addItem('🔑 Authorize Permissions', 'authorizePermissions')
    .addToUi();
}

// --- 2. MENU ACTION TRIGGERS ---
function markSentAndClearHighlightForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  sheet.getRange(row, 8).setValue("Yes");
  sheet.getRange(row, 9).setValue("Sent on " + timestamp);
  setRowHighlight(sheet, row, null); // Clear background color completely
  showAlertSafely("✅ Row marked as Sent and highlight cleared!");
}

// --- 2. MENU ACTION TRIGGERS ---
function createConfirmationDraftForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  processConfirmationAction(sheet, row, "draft");
}

function sendConfirmationEmailForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  processConfirmationAction(sheet, row, "yes");
}

function sendNoSpacesEmailForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  processNoSpacesAction(sheet, row);
}

function createWaitlistAvailableDraftForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  processWaitlistAvailableDraft(sheet, row);
}

// Helper to set background color for an entire row (Columns A through I)
function setRowHighlight(sheet, row, colorCode) {
  try {
    var rowRange = sheet.getRange(row, 1, 1, 9);
    if (colorCode) {
      rowRange.setBackground(colorCode);
    } else {
      rowRange.setBackground(null); // Clear background color
    }
  } catch (e) {
    Logger.log("Highlight error: " + e.message);
  }
}

// --- 3. CORE ACTION 1: CONFIRMATION / PLACEMENT EMAIL & DRAFT ---
function processConfirmationAction(sheet, row, actionType) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var childSchool = sheet.getRange(row, 4).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var targetYear = sheet.getRange(row, 6).getValue();
  var customNote = sheet.getRange(row, 7).getValue();
  var emailStatusCell = sheet.getRange(row, 9);
  
  if (!parentEmail || parentEmail.toString().indexOf("@") === -1) {
    showAlertSafely("Invalid or missing email address in Column E (Row " + row + ").");
    emailStatusCell.setValue("Error: Missing Email");
    return;
  }
  
  var subject = "Logic 11+ Mathematics Tuition: Enrollment Confirmation for " + (childName || "Your Child");
  
  var customNoteSection = "";
  if (customNote && customNote.toString().trim() !== "" && customNote.toString().trim() !== "None") {
    customNoteSection = "\nSpecial Note for Your Placement:\n" + customNote + "\n";
  }

  var body = "Dear " + (parentName || "Parent") + ",\n\n" +
    "Thank you for registering " + (childName || "your child") + (childSchool ? " (" + childSchool + ")" : "") + " for Logic 11+ Mathematics Tuition.\n\n" +
    "We are pleased to confirm your place in our online group tuition cohort:\n\n" +
    "📌 Registered Session Details:\n" +
    "• Cohort & Time: " + (targetYear || "Online Group Session") + "\n" +
    "• Duration: 1 Hour (Weekly)\n" +
    "• Fee Rate: £15 per 1-hour session\n" +
    "• Format: 100% Online Interactive Live Group Classroom\n" +
    "• Tutor Contact: logic11plus@gmail.com\n" +
    customNoteSection + "\n" +
    "Your child's digital classroom link and secure payment details will be shared in our next follow-up before the first class.\n\n" +
    "If you have any questions, simply reply directly to this email.\n\n" +
    "Warm regards,\n" +
    "The Logic 11+ Team\n" +
    "logic11plus@gmail.com";
  
  try {
    if (actionType === "draft") {
      GmailApp.createDraft(parentEmail, subject, body);
      emailStatusCell.setValue("Draft Created in Gmail (" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm") + ")");
      sheet.getRange(row, 8).setValue("Draft");
      setRowHighlight(sheet, row, "#fff2b2"); // Soft pastel yellow for draft created
      showAlertSafely("✅ Confirmation draft created in Gmail Drafts for " + parentEmail + "!");
    } else {
      MailApp.sendEmail({
        to: parentEmail,
        subject: subject,
        body: body,
        replyTo: "logic11plus@gmail.com",
        name: "Logic 11+ Tuition"
      });
      
      var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      emailStatusCell.setValue("Sent Confirmation on " + timestamp);
      sheet.getRange(row, 8).setValue("Yes");
      setRowHighlight(sheet, row, null); // Clear background colors once email is sent off
      showAlertSafely("✅ Confirmation email sent successfully to " + parentEmail + "!");
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 4. CORE ACTION 2: NO SPACES AVAILABLE / WAITLIST EMAIL ---
function processNoSpacesAction(sheet, row) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var childSchool = sheet.getRange(row, 4).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var targetYear = sheet.getRange(row, 6).getValue();
  var emailStatusCell = sheet.getRange(row, 9);
  
  if (!parentEmail || parentEmail.toString().indexOf("@") === -1) {
    showAlertSafely("Invalid or missing email address in Column E (Row " + row + ").");
    emailStatusCell.setValue("Error: Missing Email");
    return;
  }
  
  var subject = "Logic 11+ Mathematics Tuition: Status Update for " + (childName || "Your Child");

  var body = "Dear " + (parentName || "Parent") + ",\n\n" +
    "Thank you for registering your interest in Logic 11+ Mathematics Tuition for " + (childName || "your child") + ".\n\n" +
    "Due to high demand and our strict small-group capacity limits, we currently have no open spaces available in the requested session:\n" +
    "• Requested Slot: " + (targetYear || "Online Group Session") + "\n\n" +
    "We have automatically placed " + (childName || "your child") + " onto our Priority Waitlist. As soon as a space opens or an additional cohort is scheduled, you will be contacted immediately with first priority.\n\n" +
    "If your schedule has flexibility or you would like to enquire about other days/times, please reply directly to this email.\n\n" +
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
    emailStatusCell.setValue("Waitlist Email Sent on " + timestamp);
    sheet.getRange(row, 8).setValue("Waitlisted");
    setRowHighlight(sheet, row, "#ffd1d1"); // Soft pastel red for no spaces available
    showAlertSafely("✅ 'No Spaces / Waitlist' email sent successfully to " + parentEmail + "!");
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 5. CORE ACTION 3: WAITLIST SPACE AVAILABLE DRAFT ---
function processWaitlistAvailableDraft(sheet, row) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var childSchool = sheet.getRange(row, 4).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var rawTargetYear = sheet.getRange(row, 6).getValue().toString();
  var emailStatusCell = sheet.getRange(row, 9);
  
  if (!parentEmail || parentEmail.toString().indexOf("@") === -1) {
    showAlertSafely("Invalid or missing email address in Column E (Row " + row + ").");
    emailStatusCell.setValue("Error: Missing Email");
    return;
  }

  // Extract clean Year Group (e.g. "Year 4" or "Year 5") without slot times
  var cleanCohortYear = "Year 4 / Year 5";
  if (rawTargetYear.indexOf("Year 4") !== -1) {
    cleanCohortYear = "Year 4";
  } else if (rawTargetYear.indexOf("Year 5") !== -1) {
    cleanCohortYear = "Year 5";
  }
  
  var subject = "Logic 11+ Mathematics Tuition: Space Now Available for " + (childName || "Your Child");

  var body = "Dear " + (parentName || "Parent") + ",\n\n" +
    "Great news! A space has now opened up in our Logic 11+ Mathematics online group tuition for " + (childName || "your child") + ".\n\n" +
    "Available Session Details:\n" +
    "• Day & Time: [INSERT TIME SLOT HERE - e.g. Saturday 9:00 AM – 10:00 AM / Thursday 6:00 PM – 7:00 PM]\n" +
    "• Cohort: " + cleanCohortYear + "\n" +
    "• Fee: £15 per 1-hour session\n" +
    "• Format: 100% Online Interactive Live Group Classroom\n\n" +
    "Please reply to this email at your earliest convenience to confirm if you would like to accept this place before it is offered to the next family on our waitlist.\n\n" +
    "Warm regards,\n" +
    "The Logic 11+ Team\n" +
    "logic11plus@gmail.com";
  
  try {
    GmailApp.createDraft(parentEmail, subject, body);
    emailStatusCell.setValue("Space Open Draft Created (" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm") + ")");
    setRowHighlight(sheet, row, "#fff2b2"); // Soft pastel yellow for draft created
    showAlertSafely("✅ 'Space Available' draft created in Gmail Drafts for " + parentEmail + "! You can now open Gmail, fill in the time slot, and send.");
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 6. RECEIVE FORM SUBMISSION FROM WEBSITE & INSERT INTO SHEET ---
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

// --- 7. TRIGGER FOR AUTOMATIC TYPING IN COLUMN H ---
function onEditTrigger(e) {
  if (!e || !e.source || !e.range) return;
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  if (row > 1 && col === 8) {
    var rawVal = range.getValue();
    if (!rawVal) {
      setRowHighlight(sheet, row, null); // If cell is cleared, clear color
      return;
    }
    var val = rawVal.toString().trim().toLowerCase();
    
    if (val === "yes" || val === "sent") {
      processConfirmationAction(sheet, row, "yes");
      setRowHighlight(sheet, row, null);
    } else if (val === "draft") {
      processConfirmationAction(sheet, row, "draft");
    } else if (val === "waitlist" || val === "no space" || val === "no spaces") {
      processNoSpacesAction(sheet, row);
    } else if (val === "clear" || val === "done") {
      setRowHighlight(sheet, row, null);
    }
  }
}

// Safe UI alert helper that works both in spreadsheet UI and Apps Script editor
function showAlertSafely(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (err) {
    Logger.log(message);
  }
}

// One-time authorization helper (safe to run directly from Apps Script editor)
function authorizePermissions() {
  var testDraft = GmailApp.createDraft("logic11plus@gmail.com", "Permissions Test", "Permissions authorized successfully.");
  testDraft.deleteDraft();
  Logger.log("✅ GmailApp and MailApp permissions authorized successfully!");
  showAlertSafely("✅ Permissions successfully authorized!");
}

