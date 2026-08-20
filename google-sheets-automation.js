/**
 * ==============================================================================
 * LOGIC 11+ GOOGLE APPS SCRIPT: COMPLETE PAYMENTS, CAPACITY & WEEKLY AUTOMATION
 * ==============================================================================
 * 
 * COLOR CODE RULES:
 * 🔵 Blue (#cce5ff)   : Yet to pay / Payment Due for upcoming week.
 * 🟢 Green (#d1fae5)  : Paid & Confirmed (Post-payment Zoom email sent).
 * 🟡 Yellow (#fff2b2) : Draft created in Gmail awaiting your review.
 * 🔴 Red (#ffd1d1)    : Waitlisted / No spaces available.
 * 
 * SPREADSHEET HEADERS (Row 1):
 * Col A: Timestamp
 * Col B: Parent Name
 * Col C: Child Name
 * Col D: Child School
 * Col E: Parent Email
 * Col F: Target Slot / Year
 * Col G: Custom Email Note
 * Col H: Send Email (Awaiting Payment / Paid / Draft / Waitlisted)
 * Col I: Email Status
 * Col J: Current Week Number (e.g. 1, 2, 3...)
 */

var STRIPE_PAYMENT_LINK = "https://buy.stripe.com/test_5kQcN55Fm3A5fBW75F9R600";
var MAX_CAPACITY_PER_SLOT = 20;

var ZOOM_LINKS = {
  "Year 4 — Saturday 9:00 AM": {
    topic: "Year 4 Logic 11+ Mathematics (Saturday 9:00 AM – 10:00 AM)",
    url: "https://zoom.us/j/YOUR_Y4_SAT_MEETING_ID?pwd=YOUR_PASSWORD",
    meetingId: "111 222 3333",
    passcode: "LOGIC11"
  },
  "Year 5 — Saturday 10:00 AM": {
    topic: "Year 5 Logic 11+ Mathematics (Saturday 10:00 AM – 11:00 AM)",
    url: "https://zoom.us/j/YOUR_Y5_SAT_MEETING_ID?pwd=YOUR_PASSWORD",
    meetingId: "444 555 6666",
    passcode: "LOGIC11"
  },
  "Year 5 — Thursday 5:00 PM": {
    topic: "Year 5 Logic 11+ Mathematics (Thursday 5:00 PM – 6:00 PM)",
    url: "https://zoom.us/j/YOUR_Y5_THU_MEETING_ID?pwd=YOUR_PASSWORD",
    meetingId: "777 888 9999",
    passcode: "LOGIC11"
  },
  "Year 4 — Thursday 6:00 PM": {
    topic: "Year 4 Logic 11+ Mathematics (Thursday 6:00 PM – 7:00 PM)",
    url: "https://zoom.us/j/YOUR_Y4_THU_MEETING_ID?pwd=YOUR_PASSWORD",
    meetingId: "000 111 2222",
    passcode: "LOGIC11"
  }
};

// --- 1. ADD CUSTOM MENU TO GOOGLE SHEET TOP BAR ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🎓 Logic 11+')
    // Initial Registrations (Week 1)
    .addItem('📝 Create Initial Confirmation Draft (Week 1)', 'createConfirmationDraftForActiveRow')
    .addItem('✉️ Send Initial Confirmation Email (Week 1)', 'sendConfirmationEmailForActiveRow')
    .addSeparator()
    // Post-Payment (Zoom Link Dispatch ➔ Turns Green & Increments Week)
    .addItem('🚀 Send Post-Payment Zoom Email (Turns Green & Confirms Place)', 'sendPostPaymentEmailForActiveRow')
    .addItem('📝 Create Post-Payment Zoom Draft (For Selected Row)', 'createPostPaymentDraftForActiveRow')
    .addSeparator()
    // Continuing Students Weekly Workflow (Week 2, 3, 4...)
    .addItem('🔁 Send Continuing Student Re-Enrollment Email (Next Week)', 'sendContinuingEmailForActiveRow')
    .addItem('📝 Create Continuing Student Re-Enrollment Draft (Next Week)', 'createContinuingDraftForActiveRow')
    .addItem('🔄 Run Weekly Reset (Turn Active Sessions Blue/Due for Next Week)', 'runWeeklySessionReset')
    .addSeparator()
    // Waitlist Management
    .addItem('⏳ Send "No Spaces Available / Waitlist" Email', 'sendNoSpacesEmailForActiveRow')
    .addItem('📋 Create "Waitlist Space Available" Draft', 'createWaitlistAvailableDraftForActiveRow')
    .addSeparator()
    .addItem('📊 Refresh Capacity Tracker Tab (20 Students/Slot)', 'generateCapacityTrackerTab')
    .addItem('🔑 Authorize Permissions', 'authorizePermissions')
    .addToUi();
}

// --- 2. COLOR HIGHLIGHT HELPER (COLUMNS A THROUGH J) ---
function setRowHighlight(sheet, row, colorCode) {
  try {
    var rowRange = sheet.getRange(row, 1, 1, 10);
    if (colorCode) {
      rowRange.setBackground(colorCode);
    } else {
      rowRange.setBackground(null);
    }
  } catch (e) {
    Logger.log("Highlight error: " + e.message);
  }
}

function getRowHighlightColor(sheet, row) {
  try {
    return sheet.getRange(row, 1).getBackground();
  } catch (e) {
    return "#ffffff";
  }
}

// Helper to extract clean Zoom information based on target slot string
function getZoomInfoForSlot(slotString) {
  var str = slotString ? slotString.toString() : "";
  if (str.indexOf("Saturday 9:00 AM") !== -1) return ZOOM_LINKS["Year 4 — Saturday 9:00 AM"];
  if (str.indexOf("Saturday 10:00 AM") !== -1) return ZOOM_LINKS["Year 5 — Saturday 10:00 AM"];
  if (str.indexOf("Thursday 5:00 PM") !== -1) return ZOOM_LINKS["Year 5 — Thursday 5:00 PM"];
  if (str.indexOf("Thursday 6:00 PM") !== -1) return ZOOM_LINKS["Year 4 — Thursday 6:00 PM"];
  
  return {
    topic: "Logic 11+ Online Mathematics Tuition",
    url: "https://zoom.us",
    meetingId: "Provided by Tutor",
    passcode: "LOGIC11"
  };
}

// Helper to get or default the current student week from Column J
function getStudentWeekNumber(sheet, row) {
  var val = sheet.getRange(row, 10).getValue();
  var num = parseInt(val, 10);
  return isNaN(num) || num < 1 ? 1 : num;
}

// --- 3. MENU TRIGGER WRAPPERS ---
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

function sendPostPaymentEmailForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  processPostPaymentAction(sheet, row, "yes");
}

function createPostPaymentDraftForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  processPostPaymentAction(sheet, row, "draft");
}

function sendContinuingEmailForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  processContinuingCustomerAction(sheet, row, "yes");
}

function createContinuingDraftForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  processContinuingCustomerAction(sheet, row, "draft");
}

// --- 4. ACTION 1: INITIAL CONFIRMATION EMAIL / DRAFT (WEEK 1) ---
function processConfirmationAction(sheet, row, actionType) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var childSchool = sheet.getRange(row, 4).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var targetYear = sheet.getRange(row, 6).getValue();
  var customNote = sheet.getRange(row, 7).getValue();
  var emailStatusCell = sheet.getRange(row, 9);
  var weekCell = sheet.getRange(row, 10);
  
  if (!parentEmail || parentEmail.toString().indexOf("@") === -1) {
    showAlertSafely("Invalid or missing email address in Column E (Row " + row + ").");
    emailStatusCell.setValue("Error: Missing Email");
    return;
  }
  
  // Initialize Week 1 in Column J if empty
  if (!weekCell.getValue()) {
    weekCell.setValue(1);
  }

  var subject = "Logic 11+ Mathematics: Enrollment Details & Payment for " + (childName || "Your Child") + " (Week 1)";
  
  var customNoteSection = "";
  if (customNote && customNote.toString().trim() !== "" && customNote.toString().trim() !== "None") {
    customNoteSection = "\nSpecial Note for Your Placement:\n" + customNote + "\n";
  }

  var body = "Dear " + (parentName || "Parent") + ",\n\n" +
    "Thank you for registering " + (childName || "your child") + (childSchool ? " (" + childSchool + ")" : "") + " for Logic 11+ Mathematics Tuition.\n\n" +
    "We are pleased to offer you a place in our online interactive group cohort:\n\n" +
    "📌 Session Details:\n" +
    "• Cohort & Slot: " + (targetYear || "Online Group Session") + "\n" +
    "• Duration: 1 Hour\n" +
    "• Tuition Fee: £15 per weekly session\n" +
    "• Format: 100% Online Live Interactive Classroom (Zoom)\n" +
    customNoteSection + "\n" +
    "💳 How to Complete Payment & Secure Your Place (Week 1):\n" +
    "Please use our secure Stripe payment link to complete this week's £15 tuition fee:\n" +
    STRIPE_PAYMENT_LINK + "\n\n" +
    "⚠️ CRITICAL PAYMENT INSTRUCTIONS:\n" +
    "1. In the checkout form, please enter your child's exact name (" + (childName || "as registered") + ") so the payment automatically links to their attendance record.\n" +
    "2. If you are paying for more than 1 child, please complete a separate payment for each child using their registered name.\n" +
    "3. Important Deadline: Payment must be completed before the day of the scheduled session. If payment is not received by the day of the tuition, your child will be considered as not wanting to proceed, and the reserved slot will be released to the next family on our waitlist.\n\n" +
    "Once your payment is received, you will automatically receive your Post-Payment Confirmation email containing your child's live Zoom classroom link and joining passcode.\n\n" +
    "If you have any questions, feel free to reply directly to this email.\n\n" +
    "Warm regards,\n" +
    "The Logic 11+ Team\n" +
    "logic11plus@gmail.com";
  
  try {
    if (actionType === "draft") {
      GmailApp.createDraft(parentEmail, subject, body);
      emailStatusCell.setValue("Confirmation Draft Created (Week 1) (" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm") + ")");
      sheet.getRange(row, 8).setValue("Draft");
      setRowHighlight(sheet, row, "#fff2b2"); // Yellow
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
      emailStatusCell.setValue("Sent Confirmation (Week 1) on " + timestamp);
      sheet.getRange(row, 8).setValue("Awaiting Payment");
      setRowHighlight(sheet, row, "#cce5ff"); // Turns Blue (Yet to pay)
      showAlertSafely("✅ Confirmation email sent to " + parentEmail + " (Row turned Blue - Awaiting Payment).");
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 5. ACTION 2: POST-PAYMENT ZOOM DETAILS (TURNS GREEN, COUNTS TO CAPACITY) ---
function processPostPaymentAction(sheet, row, actionType) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var targetYear = sheet.getRange(row, 6).getValue().toString();
  var emailStatusCell = sheet.getRange(row, 9);
  var currentWeek = getStudentWeekNumber(sheet, row);
  
  if (!parentEmail || parentEmail.toString().indexOf("@") === -1) {
    showAlertSafely("Invalid or missing email address in Column E (Row " + row + ").");
    return;
  }

  // Safety Check: Verify that the row is Blue (Yet to pay) or Awaiting Payment
  var currentColor = getRowHighlightColor(sheet, row).toLowerCase();
  var isBlue = (currentColor === "#cce5ff" || currentColor === "rgb(204, 229, 255)");
  var isAwaiting = sheet.getRange(row, 8).getValue().toString().toLowerCase().indexOf("awaiting") !== -1 || sheet.getRange(row, 8).getValue().toString().toLowerCase().indexOf("due") !== -1;

  if (actionType === "yes" && !isBlue && !isAwaiting) {
    var proceed = SpreadsheetApp.getUi().alert(
      "⚠️ Payment Status Notice",
      "This row is not currently marked as Blue / Awaiting Payment. Has payment been confirmed in Stripe for " + childName + " (Week " + currentWeek + ")?\n\nClick OK to proceed anyway, or Cancel to stop.",
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (proceed !== SpreadsheetApp.getUi().Button.OK) {
      return;
    }
  }

  var zoomInfo = getZoomInfoForSlot(targetYear);
  var subject = "Logic 11+ Mathematics: Week " + currentWeek + " Payment Received & Zoom Link for " + (childName || "Your Child");

  var body = "Dear " + (parentName || "Parent") + ",\n\n" +
    "Thank you! We have received your payment of £15 for " + (childName || "your child") + " (Week " + currentWeek + ").\n\n" +
    "Your place is 100% confirmed for this week's live session.\n\n" +
    "📹 LIVE DIGITAL CLASSROOM DETAILS (WEEK " + currentWeek + "):\n" +
    "• Session: " + zoomInfo.topic + "\n" +
    "• Zoom Join Link: " + zoomInfo.url + "\n" +
    "• Meeting ID: " + zoomInfo.meetingId + "\n" +
    "• Passcode: " + zoomInfo.passcode + "\n\n" +
    "Student Guidelines:\n" +
    "1. Please ensure " + (childName || "your child") + " joins 5 minutes before the start time with a notebook, pencil, and ruler.\n" +
    "2. Microphones and cameras should be working for active participation.\n\n" +
    "We look forward to a great lesson!\n\n" +
    "Warm regards,\n" +
    "The Logic 11+ Team\n" +
    "logic11plus@gmail.com";

  try {
    if (actionType === "draft") {
      GmailApp.createDraft(parentEmail, subject, body);
      emailStatusCell.setValue("Post-Payment Draft Created (Week " + currentWeek + ") (" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm") + ")");
      showAlertSafely("✅ Post-Payment Zoom draft created in Gmail for " + parentEmail + "!");
    } else {
      MailApp.sendEmail({
        to: parentEmail,
        subject: subject,
        body: body,
        replyTo: "logic11plus@gmail.com",
        name: "Logic 11+ Tuition"
      });
      
      var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      emailStatusCell.setValue("Zoom Link Sent on " + timestamp + " (Week " + currentWeek + ")");
      sheet.getRange(row, 8).setValue("Paid");
      setRowHighlight(sheet, row, "#d1fae5"); // Turns Green (Paid & Counted)
      showAlertSafely("✅ Post-Payment Zoom details for Week " + currentWeek + " sent to " + parentEmail + "! (Row turned Green - Space Confirmed).");
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 6. ACTION 3: CONTINUING CUSTOMER RE-ENROLLMENT EMAIL / DRAFT (NEXT WEEK) ---
function processContinuingCustomerAction(sheet, row, actionType) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var targetYear = sheet.getRange(row, 6).getValue().toString();
  var emailStatusCell = sheet.getRange(row, 9);
  var currentWeek = getStudentWeekNumber(sheet, row);
  var nextWeek = currentWeek + 1;
  
  if (!parentEmail || parentEmail.toString().indexOf("@") === -1) {
    showAlertSafely("Invalid or missing email address in Column E (Row " + row + ").");
    return;
  }

  var subject = "Logic 11+ Mathematics: Week " + nextWeek + " Re-enrollment for " + (childName || "Your Child");

  var body = "Dear " + (parentName || "Parent") + ",\n\n" +
    "We hope " + (childName || "your child") + " enjoyed their recent Logic 11+ Mathematics session (Week " + currentWeek + ")!\n\n" +
    "To confirm your child's attendance for WEEK " + nextWeek + "'s session (" + (targetYear || "Online Group Session") + "), please complete your weekly £15 tuition payment via the link below:\n\n" +
    "💳 Week " + nextWeek + " Re-enrollment Link:\n" +
    STRIPE_PAYMENT_LINK + "\n\n" +
    "⚠️ IMPORTANT PAYMENT & ATTENDANCE POLICY:\n" +
    "1. Please enter your child's exact registered name (" + (childName || "as registered") + ") at checkout.\n" +
    "2. If you have multiple children, please submit a separate payment for each child.\n" +
    "3. Deadline: Payment must be completed prior to the day of the session. If payment is not completed before the session day, we will assume you do not wish to continue for this week, and the place will be opened to our waitlist.\n\n" +
    "As soon as payment is confirmed, your Week " + nextWeek + " live classroom Zoom link will be dispatched.\n\n" +
    "Warm regards,\n" +
    "The Logic 11+ Team\n" +
    "logic11plus@gmail.com";

  try {
    if (actionType === "draft") {
      GmailApp.createDraft(parentEmail, subject, body);
      emailStatusCell.setValue("Continuing Draft Created (Week " + nextWeek + ") (" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm") + ")");
      sheet.getRange(row, 8).setValue("Draft");
      setRowHighlight(sheet, row, "#fff2b2"); // Yellow
      showAlertSafely("✅ Continuing Student draft created in Gmail for Week " + nextWeek + "!");
    } else {
      MailApp.sendEmail({
        to: parentEmail,
        subject: subject,
        body: body,
        replyTo: "logic11plus@gmail.com",
        name: "Logic 11+ Tuition"
      });
      
      var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      emailStatusCell.setValue("Continuing Notice Sent on " + timestamp + " (Week " + nextWeek + ")");
      sheet.getRange(row, 8).setValue("Payment Due");
      sheet.getRange(row, 10).setValue(nextWeek); // Advance to next week in Column J
      setRowHighlight(sheet, row, "#cce5ff"); // Turns Blue (Due for next week)
      showAlertSafely("✅ Continuing Student notice for Week " + nextWeek + " sent to " + parentEmail + " (Row turned Blue).");
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 7. ACTION 4: WEEKLY AUTOMATIC RESET (TURNS PASSED SESSIONS BLUE) ---
function runWeeklySessionReset() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Submissions") || ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    showAlertSafely("No student rows found to reset.");
    return;
  }

  var count = 0;
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 2;
    var currentSendVal = data[i][7].toString().toLowerCase();
    var currentWeek = parseInt(data[i][9], 10);
    if (isNaN(currentWeek) || currentWeek < 1) currentWeek = 1;

    // Reset rows that were previously paid / active and not waitlisted
    if (currentSendVal !== "waitlisted" && currentSendVal !== "no space" && currentSendVal !== "no spaces") {
      setRowHighlight(sheet, rowNum, "#cce5ff"); // Turn Blue (Payment Due for next session)
      sheet.getRange(rowNum, 8).setValue("Payment Due");
      sheet.getRange(rowNum, 10).setValue(currentWeek + 1); // Increment week number
      count++;
    }
  }

  showAlertSafely("🔄 Weekly Reset Complete: " + count + " active student rows advanced to their next week and turned Blue (Payment Due).");
}

// --- 8. ACTION 5: CAPACITY TRACKER TAB GENERATOR (20 STUDENTS PER SLOT) ---
function generateCapacityTrackerTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var capacitySheet = ss.getSheetByName("Capacity Tracker");
  
  if (!capacitySheet) {
    capacitySheet = ss.insertSheet("Capacity Tracker");
  }

  capacitySheet.clear();

  // Set Headers
  capacitySheet.getRange("A1:E1").setValues([[
    "Slot / Cohort", "Max Limit", "Active Paid Spaces (Green / Zoom Sent)", "Spaces Remaining", "Status"
  ]]).setFontWeight("bold").setBackground("#183153").setFontColor("#ffffff");

  var slots = [
    ["Year 4 — Saturday 9:00 AM – 10:00 AM", MAX_CAPACITY_PER_SLOT, '=COUNTIFS(Submissions!F:F, "*Saturday 9:00 AM*", Submissions!I:I, "*Zoom Link Sent*")', '=B2-C2', '=IF(D2<=0, "FULL", D2&" SPACES LEFT")'],
    ["Year 5 — Saturday 10:00 AM – 11:00 AM", MAX_CAPACITY_PER_SLOT, '=COUNTIFS(Submissions!F:F, "*Saturday 10:00 AM*", Submissions!I:I, "*Zoom Link Sent*")', '=B3-C3', '=IF(D3<=0, "FULL", D3&" SPACES LEFT")'],
    ["Year 5 — Thursday 5:00 PM – 6:00 PM", MAX_CAPACITY_PER_SLOT, '=COUNTIFS(Submissions!F:F, "*Thursday 5:00 PM*", Submissions!I:I, "*Zoom Link Sent*")', '=B4-C4', '=IF(D4<=0, "FULL", D4&" SPACES LEFT")'],
    ["Year 4 — Thursday 6:00 PM – 7:00 PM", MAX_CAPACITY_PER_SLOT, '=COUNTIFS(Submissions!F:F, "*Thursday 6:00 PM*", Submissions!I:I, "*Zoom Link Sent*")', '=B5-C5', '=IF(D5<=0, "FULL", D5&" SPACES LEFT")']
  ];

  capacitySheet.getRange(2, 1, slots.length, 5).setValues(slots);
  capacitySheet.autoResizeColumns(1, 5);

  showAlertSafely("✅ 'Capacity Tracker' tab updated! (Limit: 20 students per slot).");
}

// --- 9. ACTION 6: NO SPACES / WAITLIST EMAIL ---
function processNoSpacesAction(sheet, row) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
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
    setRowHighlight(sheet, row, "#ffd1d1"); // Red
    showAlertSafely("✅ 'No Spaces / Waitlist' email sent successfully to " + parentEmail + "!");
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 10. ACTION 7: WAITLIST SPACE AVAILABLE DRAFT ---
function processWaitlistAvailableDraft(sheet, row) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var rawTargetYear = sheet.getRange(row, 6).getValue().toString();
  var emailStatusCell = sheet.getRange(row, 9);
  
  if (!parentEmail || parentEmail.toString().indexOf("@") === -1) {
    showAlertSafely("Invalid or missing email address in Column E (Row " + row + ").");
    emailStatusCell.setValue("Error: Missing Email");
    return;
  }

  var cleanCohortYear = "Year 4 / Year 5";
  if (rawTargetYear.indexOf("Year 4") !== -1) {
    cleanCohortYear = "Year 4";
  } else if (rawTargetYear.indexOf("Year 5") !== -1) {
    cleanCohortYear = "Year 5";
  }
  
  var subject = "Logic 11+ Mathematics: Space Now Available for " + (childName || "Your Child");

  var body = "Dear " + (parentName || "Parent") + ",\n\n" +
    "Great news! A space has opened up in our Logic 11+ Mathematics online group tuition for " + (childName || "your child") + ".\n\n" +
    "Available Session Details:\n" +
    "• Day & Time: [INSERT TIME SLOT HERE - e.g. Saturday 9:00 AM – 10:00 AM / Thursday 6:00 PM – 7:00 PM]\n" +
    "• Cohort: " + cleanCohortYear + "\n" +
    "• Fee: £15 per 1-hour session\n" +
    "• Format: 100% Online Interactive Live Group Classroom\n\n" +
    "💳 Payment Link to Secure Place:\n" +
    STRIPE_PAYMENT_LINK + "\n\n" +
    "⚠️ Critical Payment Note:\n" +
    "Please enter your child's exact registered name (" + (childName || "as registered") + ") at checkout. If you have multiple children, please submit a separate payment for each child.\n" +
    "Payment must be completed before the session day to confirm the place; otherwise, the space will be offered to the next family on our waitlist.\n\n" +
    "Once payment is received, your child's Zoom classroom link will be emailed to you.\n\n" +
    "Warm regards,\n" +
    "The Logic 11+ Team\n" +
    "logic11plus@gmail.com";
  
  try {
    GmailApp.createDraft(parentEmail, subject, body);
    emailStatusCell.setValue("Space Open Draft Created (" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm") + ")");
    setRowHighlight(sheet, row, "#fff2b2"); // Yellow
    showAlertSafely("✅ 'Space Available' draft created in Gmail Drafts for " + parentEmail + "! Fill in the day/time and send.");
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 11. WEBHOOK: RECEIVE FORM SUBMISSION FROM WEBSITE ---
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
    var weekNumber = 1;

    sheet.appendRow([
      timestamp,
      parentName,
      childName,
      childSchool,
      parentEmail,
      targetYear,
      customNote,
      sendFurtherEmail,
      emailStatus,
      weekNumber
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

// --- 12. ON-EDIT TRIGGER ---
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
      processConfirmationAction(sheet, row, "yes");
    } else if (val === "draft") {
      processConfirmationAction(sheet, row, "draft");
    } else if (val === "paid" || val === "zoom") {
      processPostPaymentAction(sheet, row, "yes");
    } else if (val === "waitlist" || val === "no space" || val === "no spaces") {
      processNoSpacesAction(sheet, row);
    }
  }
}

// Safe UI Alert helper
function showAlertSafely(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (err) {
    Logger.log(message);
  }
}

// One-time authorization helper
function authorizePermissions() {
  var testDraft = GmailApp.createDraft("logic11plus@gmail.com", "Permissions Test", "Permissions authorized successfully.");
  testDraft.deleteDraft();
  Logger.log("✅ GmailApp and MailApp permissions authorized successfully!");
  showAlertSafely("✅ Permissions successfully authorized!");
}

