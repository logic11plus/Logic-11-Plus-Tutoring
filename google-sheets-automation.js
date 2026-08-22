/**
 * ==============================================================================
 * LOGIC 11+ GOOGLE APPS SCRIPT: DYNAMIC WEEKLY TRACKING & AUTO-PAYMENT
 * ==============================================================================
 * 
 * VIVID COLOR LIFECYCLE:
 * 🔴 VIVID RED (#ff7b72)    : NEW SUBMISSION (No initial confirmation email sent yet)
 * 🟡 VIVID YELLOW (#ffe066) : INITIAL CONFIRMATION SENT (Week 1 Registered, awaiting payment)
 * 🟣 VIVID PURPLE (#c084fc) : SUNDAY RESET (Returning student ready for next week)
 * 🔵 VIVID BLUE (#60a5fa)   : RE-ENROLLMENT LINK SENT (Week 2+ payment email dispatched)
 * 🟢 VIVID GREEN (#4ade80)  : PAID & ZOOM SENT (Confirmed for this week, active session space)
 * 
 * SPREADSHEET HEADERS (Row 1):
 * Col A: Timestamp
 * Col B: Parent Name
 * Col C: Child Name
 * Col D: Child School
 * Col E: Parent Email
 * Col F: Target Slot / Year
 * Col G: Custom Email Note
 * Col H: Send Email / Action
 * Col I: Email Status & Delivery Log
 * Col J: Current Week Number (1, 2, 3...)
 * Col K: Week 1 (Paid / Not Paid / Cancelled)
 * Col L: Week 2 (Paid / Not Paid / Cancelled)
 * Col M: Week 3 ... (Added dynamically as students advance)
 */

var STRIPE_PAYMENT_LINK = "https://buy.stripe.com/test_5kQcN55Fm3A5fBW75F9R600";
var MAX_CAPACITY_PER_SLOT = 20;

// Vivid Color Hex Palette
var COLOR_RED_NEW       = "#ff7b72"; // Red: New lead, no email yet
var COLOR_YELLOW_CONF   = "#ffe066"; // Yellow: Initial confirmation sent, awaiting Week 1 payment
var COLOR_PURPLE_RESET  = "#c084fc"; // Purple: Sunday reset, returning student ready for re-enrollment
var COLOR_BLUE_RE_ENROLL= "#60a5fa"; // Blue: Re-enrollment payment link sent to parent
var COLOR_GREEN_ZOOM    = "#4ade80"; // Green: Paid & Zoom link sent (active confirmed space)
var COLOR_GRAY_CANCEL   = "#cbd5e1"; // Gray: Week cancelled by parent

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
    // Stage 1: Initial Confirmations (Red ➔ Yellow)
    .addItem('✉️ Send Initial Confirmation Email (Turns 🟡 Yellow)', 'sendConfirmationEmailForActiveRow')
    .addItem('📝 Create Initial Confirmation Draft', 'createConfirmationDraftForActiveRow')
    .addSeparator()
    // Stage 2: Re-enrollments (Purple ➔ Blue)
    .addItem('🔁 Send Re-Enrollment Payment Email (Turns 🔵 Blue)', 'sendContinuingEmailForActiveRow')
    .addItem('📝 Create Re-Enrollment Draft', 'createContinuingDraftForActiveRow')
    .addSeparator()
    // Stage 3: Zoom Link Dispatch (Instant / Manual ➔ Green)
    .addItem('🚀 Send Instant Zoom Link (Turns 🟢 Green)', 'sendPostPaymentEmailForActiveRow')
    .addItem('📝 Create Zoom Link Draft', 'createPostPaymentDraftForActiveRow')
    .addSeparator()
    // Stage 4: Cancellation Management
    .addItem('🚫 Cancel This Week for Student (Skips Week & Emails Parent)', 'cancelCurrentWeekForActiveRow')
    .addSeparator()
    // Stage 5: Sunday Weekly Reset
    .addItem('🔄 Run Sunday Weekly Reset (Active Students Turn 🟣 Purple & Add Next Week)', 'runWeeklySessionReset')
    .addSeparator()
    .addItem('🎨 Apply Red Highlight to Any Uncontacted Rows', 'applyRedToNewSubmissions')
    .addItem('📊 Refresh Capacity Tracker Tab (20 Seats/Slot)', 'generateCapacityTrackerTab')
    .addItem('🔑 Authorize Permissions', 'authorizePermissions')
    .addToUi();
}

// --- 2. ROW HIGHLIGHT HELPER (ACROSS ALL COLUMNS) ---
function setRowHighlight(sheet, row, colorCode) {
  try {
    var maxCols = Math.max(sheet.getLastColumn(), 15);
    var rowRange = sheet.getRange(row, 1, 1, maxCols);
    if (colorCode) {
      rowRange.setBackground(colorCode);
    } else {
      rowRange.setBackground(null);
    }
  } catch (e) {
    Logger.log("Highlight error: " + e.message);
  }
}

// Helper to get or ensure week tracking column (Col K = Week 1, Col L = Week 2, Col M = Week 3...)
function getWeekColumnIndex(sheet, weekNum) {
  var lastCol = Math.max(sheet.getLastColumn(), 15);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 1. Search existing headers for "Week 1", "Week1", "Week 1 Paid", etc.
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c] ? headers[c].toString().trim().toLowerCase() : "";
    // Match "week 1", "week1", "week 1 status", "week 1 paid"
    if (h === ("week " + weekNum) || h === ("week" + weekNum) || h.indexOf("week " + weekNum) === 0 || h.indexOf("week" + weekNum) === 0) {
      return c + 1; // 1-based column index
    }
  }

  // 2. If not found, use Column 10 + weekNum (Col 11 for Week 1, Col 12 for Week 2...)
  var targetCol = 10 + weekNum;
  sheet.getRange(1, targetCol).setValue("Week " + weekNum).setFontWeight("bold").setBackground("#e2e8f0");
  return targetCol;
}

// Helper to set weekly payment status: "Paid", "Not Paid", or "Cancelled"
function setStudentWeekPaymentStatus(sheet, row, weekNum, status) {
  var col = getWeekColumnIndex(sheet, weekNum);
  var cell = sheet.getRange(row, col);
  
  if (status === "Paid") {
    cell.setValue("Paid").setFontColor("#15803d").setFontWeight("bold").setBackground("#d1fae5");
  } else if (status === "Cancelled") {
    cell.setValue("Cancelled").setFontColor("#64748b").setFontWeight("bold").setBackground("#f1f5f9");
  } else {
    cell.setValue("Not Paid").setFontColor("#b91c1c").setFontWeight("normal").setBackground(null);
  }
  SpreadsheetApp.flush();
}

// Helper to get current student week number from Column J
function getStudentWeekNumber(sheet, row) {
  var val = sheet.getRange(row, 10).getValue();
  var num = parseInt(val, 10);
  return isNaN(num) || num < 1 ? 1 : num;
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

// Helper to scan all uncontacted rows and apply Red highlight
function applyRedToNewSubmissions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Submissions") || ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 2;
    var sendStatus = data[i][7].toString().toLowerCase();
    var emailStatus = data[i][8].toString().toLowerCase();
    
    // If no confirmation email has been sent yet, color Red
    if (sendStatus === "no" || emailStatus.indexOf("pending") !== -1 || emailStatus.indexOf("new") !== -1) {
      setRowHighlight(sheet, rowNum, COLOR_RED_NEW);
      setStudentWeekPaymentStatus(sheet, rowNum, 1, "Not Paid");
      count++;
    }
  }
  showAlertSafely("🔴 Checked rows: " + count + " new submission(s) highlighted in Red.");
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

function cancelCurrentWeekForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  processCancelWeekAction(sheet, row);
}

// --- 4. ACTION 1: INITIAL CONFIRMATION EMAIL (RED ➔ YELLOW) ---
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
  
  // Set Week 1 in Column J and Column K
  weekCell.setValue(1);
  setStudentWeekPaymentStatus(sheet, row, 1, "Not Paid");

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
    "3. Important Deadline: Payment must be completed before the day of the scheduled session. If payment is not received by the day of the tuition, your child will be considered as not wanting to proceed, and the reserved slot will be released to another family.\n\n" +
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
      setRowHighlight(sheet, row, COLOR_YELLOW_CONF); // Turns VIVID YELLOW
      showAlertSafely("✅ Confirmation email sent to " + parentEmail + "!\nRow is now 🟡 YELLOW (Awaiting Week 1 Payment).");
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 5. ACTION 2: INSTANT POST-PAYMENT ZOOM DETAILS (TURNS GREEN) ---
function processPostPaymentAction(sheet, row, actionType, isWebhook) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var targetYear = sheet.getRange(row, 6).getValue().toString();
  var emailStatusCell = sheet.getRange(row, 9);
  var currentWeek = getStudentWeekNumber(sheet, row);
  
  if (!parentEmail || parentEmail.toString().indexOf("@") === -1) {
    if (!isWebhook) showAlertSafely("Invalid or missing email address in Column E (Row " + row + ").");
    return;
  }

  // Safety Check: If this week's Zoom link has ALREADY been sent, prevent duplicate sending
  var currentStatus = emailStatusCell.getValue().toString();
  if (actionType !== "draft" && currentStatus.indexOf("Zoom Link Sent") !== -1 && currentStatus.indexOf("Week " + currentWeek) !== -1) {
    Logger.log("Zoom link already dispatched for " + parentEmail + " (Week " + currentWeek + "). Skipping duplicate send.");
    return;
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
      if (!isWebhook) showAlertSafely("✅ Post-Payment Zoom draft created in Gmail for " + parentEmail + "!");
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
      sheet.getRange(row, 8).setValue("Zoom Sent");
      
      // Update Column for this week to Paid (Green bold text)
      setStudentWeekPaymentStatus(sheet, row, currentWeek, "Paid");
      
      // Turn row Vivid Green
      setRowHighlight(sheet, row, COLOR_GREEN_ZOOM);
      SpreadsheetApp.flush(); // Commit all sheet cell updates immediately
      
      if (!isWebhook) {
        showAlertSafely("🚀 Instant Zoom link sent to " + parentEmail + " for Week " + currentWeek + "!\nRow is now 🟢 GREEN (Space Confirmed).");
      }
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    if (!isWebhook) showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 6. ACTION 3: CONTINUING CUSTOMER RE-ENROLLMENT EMAIL (TURNS BLUE) ---
function processContinuingCustomerAction(sheet, row, actionType) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var targetYear = sheet.getRange(row, 6).getValue().toString();
  var emailStatusCell = sheet.getRange(row, 9);
  
  var targetWeek = getStudentWeekNumber(sheet, row);
  var previousWeek = targetWeek > 1 ? targetWeek - 1 : 1;
  
  if (!parentEmail || parentEmail.toString().indexOf("@") === -1) {
    showAlertSafely("Invalid or missing email address in Column E (Row " + row + ").");
    return;
  }

  var subject = "Logic 11+ Mathematics: Week " + targetWeek + " Re-enrollment for " + (childName || "Your Child");

  var body = "Dear " + (parentName || "Parent") + ",\n\n" +
    "We hope " + (childName || "your child") + " enjoyed their recent Logic 11+ Mathematics session (Week " + previousWeek + ")!\n\n" +
    "To confirm your child's attendance for WEEK " + targetWeek + "'s session (" + (targetYear || "Online Group Session") + "), please complete your weekly £15 tuition payment via the link below:\n\n" +
    "💳 Week " + targetWeek + " Re-enrollment Link:\n" +
    STRIPE_PAYMENT_LINK + "\n\n" +
    "⚠️ IMPORTANT PAYMENT & ATTENDANCE POLICY:\n" +
    "1. Please enter your child's exact registered name (" + (childName || "as registered") + ") at checkout.\n" +
    "2. If you have multiple children, please submit a separate payment for each child.\n" +
    "3. Deadline: Payment must be completed prior to the day of the session. If payment is not completed before the session day, we will assume you do not wish to continue for this week, and the place will be released.\n\n" +
    "As soon as payment is confirmed, your Week " + targetWeek + " live classroom Zoom link will be dispatched.\n\n" +
    "Warm regards,\n" +
    "The Logic 11+ Team\n" +
    "logic11plus@gmail.com";

  try {
    if (actionType === "draft") {
      GmailApp.createDraft(parentEmail, subject, body);
      emailStatusCell.setValue("Continuing Draft Created (Week " + targetWeek + ") (" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm") + ")");
      sheet.getRange(row, 8).setValue("Draft");
      showAlertSafely("✅ Continuing Student draft created in Gmail for Week " + targetWeek + "!");
    } else {
      MailApp.sendEmail({
        to: parentEmail,
        subject: subject,
        body: body,
        replyTo: "logic11plus@gmail.com",
        name: "Logic 11+ Tuition"
      });
      
      var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      emailStatusCell.setValue("Re-enrollment Link Sent on " + timestamp + " (Week " + targetWeek + ")");
      sheet.getRange(row, 8).setValue("Payment Due (Week " + targetWeek + ")");
      
      // 🔵 Turns VIVID BLUE: Indicates Re-enrollment payment link has been dispatched
      setRowHighlight(sheet, row, COLOR_BLUE_RE_ENROLL);
      showAlertSafely("✅ Re-enrollment payment link for Week " + targetWeek + " sent to " + parentEmail + "!\nRow is now 🔵 BLUE (Awaiting Re-enrollment Payment).");
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 7. ACTION 4: CANCEL CURRENT WEEK (SKIPS TO NEXT WEEK & EMAILS PARENT) ---
function processCancelWeekAction(sheet, row) {
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

  var confirmCancel = SpreadsheetApp.getUi().alert(
    "Cancel Week " + currentWeek + " Confirmation",
    "Are you sure you want to cancel Week " + currentWeek + " for " + (childName || "this student") + "?\n\nThis will mark Week " + currentWeek + " as Cancelled, advance them to Week " + nextWeek + ", and email the parent.",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (confirmCancel !== SpreadsheetApp.getUi().Button.YES) return;

  var subject = "Logic 11+ Mathematics: Week " + currentWeek + " Session Update for " + (childName || "Your Child");

  var body = "Dear " + (parentName || "Parent") + ",\n\n" +
    "Thank you for letting us know regarding " + (childName || "your child") + "'s attendance.\n\n" +
    "We have noted that " + (childName || "your child") + " will not be attending the Logic 11+ Mathematics tuition session for this week (Week " + currentWeek + ").\n\n" +
    "Your child's profile remains active in our system. We look forward to seeing " + (childName || "them") + " next week for Week " + nextWeek + "'s session (" + (targetYear || "Online Group Session") + ")!\n\n" +
    "Prior to next week's lesson, you will receive your Week " + nextWeek + " re-enrollment payment link.\n\n" +
    "If you have any questions in the meantime, feel free to reply directly to this email.\n\n" +
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
    
    // 1. Mark Current Week column as "Cancelled" (Gray bold text)
    setStudentWeekPaymentStatus(sheet, row, currentWeek, "Cancelled");
    
    // 2. Advance student to Next Week
    sheet.getRange(row, 10).setValue(nextWeek);
    setStudentWeekPaymentStatus(sheet, row, nextWeek, "Not Paid");
    
    // 3. Update Status and Row Highlight to Purple (Ready for next week)
    sheet.getRange(row, 8).setValue("Week " + currentWeek + " Cancelled");
    emailStatusCell.setValue("Week " + currentWeek + " Cancelled Email Sent on " + timestamp);
    setRowHighlight(sheet, row, COLOR_PURPLE_RESET);

    showAlertSafely("✅ Week " + currentWeek + " marked as Cancelled for " + (childName || "student") + "!\nCancellation confirmation emailed to " + parentEmail + ".\nStudent is now set up for Week " + nextWeek + " (🟣 Purple).");
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// Helper to find the main student data sheet (supports "Students", "Submissions", or active tab)
function getStudentSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName("Students") || 
         ss.getSheetByName("Submissions") || 
         ss.getSheetByName("Sheet1") || 
         ss.getSheets()[0];
}

// --- 8. ACTION 5: SUNDAY WEEKLY RESET (ONLY ADVANCES CONFIRMED STUDENTS) ---
function runWeeklySessionReset() {
  var sheet = getStudentSheet();
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

    // Advance students who attended (Green / Zoom Sent / Paid)
    var isConfirmedForCurrentWeek = (currentSendVal.indexOf("zoom") !== -1 || currentSendVal.indexOf("paid") !== -1);

    if (isConfirmedForCurrentWeek) {
      var nextWeek = currentWeek + 1;
      
      // Update Column J to next week
      sheet.getRange(rowNum, 10).setValue(nextWeek);
      
      // Add Week (X+1) column & mark "Not Paid"
      setStudentWeekPaymentStatus(sheet, rowNum, nextWeek, "Not Paid");
      
      // Turn Purple (Returning Student ready for re-enrollment email)
      setRowHighlight(sheet, rowNum, COLOR_PURPLE_RESET);
      sheet.getRange(rowNum, 8).setValue("Waiting for Re-enrollment");
      sheet.getRange(rowNum, 9).setValue("Reset for Week " + nextWeek + " (Waiting Payment)");
      count++;
    }
  }

  showAlertSafely("🔄 Sunday Weekly Reset Complete:\n" + count + " active student(s) who attended advanced to their next week, received their new Week Column (Not Paid), and turned 🟣 PURPLE.");
}

// --- 9. ACTION 6: CAPACITY TRACKER TAB GENERATOR (20 SEATS PER SLOT) ---
function generateCapacityTrackerTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = getStudentSheet();
  var capacitySheet = ss.getSheetByName("Capacity Tracker");
  
  if (!capacitySheet) {
    capacitySheet = ss.insertSheet("Capacity Tracker");
  }

  // Set Headers (without clearing custom column widths or row heights)
  capacitySheet.getRange("A1:E1").setValues([[
    "Slot / Cohort", "Max Limit", "Active Students (🟢 Green / 🟣 Purple / 🔵 Blue)", "Spaces Remaining", "Status"
  ]]).setFontWeight("bold").setBackground("#183153").setFontColor("#ffffff");

  // Calculate live count directly from studentSheet (Students / Submissions tab)
  var lastRow = studentSheet.getLastRow();
  var countY4Sat = 0;
  var countY5Sat = 0;
  var countY5Thu = 0;
  var countY4Thu = 0;

  if (lastRow > 1) {
    var data = studentSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var bgColors = studentSheet.getRange(2, 1, lastRow - 1, 1).getBackgrounds();

    for (var i = 0; i < data.length; i++) {
      var slot = data[i][5] ? data[i][5].toString().toLowerCase() : "";
      var sendStatus = data[i][7] ? data[i][7].toString().toLowerCase() : "";
      var emailStatus = data[i][8] ? data[i][8].toString().toLowerCase() : "";
      var bg = bgColors[i][0] ? bgColors[i][0].toLowerCase() : "";

      // Check if student is active:
      // Background is Green, Purple, Blue OR status contains zoom / paid / waiting / due
      var isGreen = (bg === "#4ade80" || bg === "#d1fae5" || bg === "rgb(74, 222, 128)" || bg === "rgb(209, 250, 229)" || sendStatus.indexOf("zoom") !== -1 || emailStatus.indexOf("zoom") !== -1);
      var isPurple = (bg === "#c084fc" || bg === "rgb(192, 132, 252)" || sendStatus.indexOf("waiting") !== -1 || emailStatus.indexOf("waiting") !== -1);
      var isBlue = (bg === "#60a5fa" || bg === "#cce5ff" || bg === "rgb(96, 165, 250)" || bg === "rgb(204, 229, 255)" || sendStatus.indexOf("due") !== -1 || sendStatus.indexOf("paid") !== -1);
      
      var isActive = (isGreen || isPurple || isBlue);

      if (isActive) {
        if (slot.indexOf("9") !== -1 || (slot.indexOf("sat") !== -1 && slot.indexOf("4") !== -1)) {
          countY4Sat++;
        } else if (slot.indexOf("10") !== -1 || (slot.indexOf("sat") !== -1 && slot.indexOf("5") !== -1)) {
          countY5Sat++;
        } else if (slot.indexOf("5") !== -1 || (slot.indexOf("thu") !== -1 && slot.indexOf("5") !== -1)) {
          countY5Thu++;
        } else if (slot.indexOf("6") !== -1 || (slot.indexOf("thu") !== -1 && slot.indexOf("4") !== -1)) {
          countY4Thu++;
        }
      }
    }
  }

  var slots = [
    ["Year 4 — Saturday 9:00 AM – 10:00 AM", MAX_CAPACITY_PER_SLOT, countY4Sat, '=B2-C2', '=IF(D2<=0, "FULL", D2&" SPACES LEFT")'],
    ["Year 5 — Saturday 10:00 AM – 11:00 AM", MAX_CAPACITY_PER_SLOT, countY5Sat, '=B3-C3', '=IF(D3<=0, "FULL", D3&" SPACES LEFT")'],
    ["Year 5 — Thursday 5:00 PM – 6:00 PM", MAX_CAPACITY_PER_SLOT, countY5Thu, '=B4-C4', '=IF(D4<=0, "FULL", D4&" SPACES LEFT")'],
    ["Year 4 — Thursday 6:00 PM – 7:00 PM", MAX_CAPACITY_PER_SLOT, countY4Thu, '=B5-C5', '=IF(D5<=0, "FULL", D5&" SPACES LEFT")']
  ];

  capacitySheet.getRange(2, 1, slots.length, 5).setValues(slots);

  showAlertSafely("✅ 'Capacity Tracker' updated successfully from sheet '" + studentSheet.getName() + "'!\n\nActive students detected:\n• Year 4 (Sat 9:00 AM): " + countY4Sat + "\n• Year 5 (Sat 10:00 AM): " + countY5Sat + "\n• Year 5 (Thu 5:00 PM): " + countY5Thu + "\n• Year 4 (Thu 6:00 PM): " + countY4Thu);
}

// --- 10. WEBHOOK: RECEIVE FORM SUBMISSIONS & STRIPE PAYMENT WEBHOOKS ---
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var sheet = getStudentSheet();
    var rawContents = e.postData.contents;
    var data = JSON.parse(rawContents);

    // =========================================================================
    // CASE A: STRIPE WEBHOOK EVENT (checkout.session.completed ONLY)
    // =========================================================================
    if (data && data.type && data.type === "checkout.session.completed") {
      var session = data.data.object;
      var stripeEventId = data.id || session.id || "";
      var stripeEmail = (session.customer_details && session.customer_details.email) ? session.customer_details.email.toLowerCase().trim() : "";
      
      // Deduplication Cache Check: Ignore if this exact Stripe event was processed in last 6 hours
      var cache = CacheService.getScriptCache();
      if (stripeEventId && cache.get(stripeEventId)) {
        Logger.log("Duplicate Stripe event ignored: " + stripeEventId);
        return ContentService
          .createTextOutput(JSON.stringify({ result: "duplicate_ignored", event_id: stripeEventId }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (stripeEventId) {
        cache.put(stripeEventId, "processed", 21600); // Cache for 6 hours
      }

      // Extract custom field (Child's Name) from Stripe checkout
      var stripeChildName = "";
      if (session.custom_fields && session.custom_fields.length > 0) {
        for (var f = 0; f < session.custom_fields.length; f++) {
          var field = session.custom_fields[f];
          if (field.text && field.text.value) {
            stripeChildName = field.text.value.toLowerCase().trim();
            break;
          }
        }
      }

      var lastRow = sheet.getLastRow();
      var matchedRow = -1;

      if (lastRow > 1) {
        var sheetData = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
        
        // 1st Priority Match: Child's registered name (Column C)
        if (stripeChildName) {
          for (var i = 0; i < sheetData.length; i++) {
            var rowChildName = sheetData[i][2].toString().toLowerCase().trim();
            if (rowChildName && (rowChildName === stripeChildName || stripeChildName.indexOf(rowChildName) !== -1 || rowChildName.indexOf(stripeChildName) !== -1)) {
              matchedRow = i + 2;
              break;
            }
          }
        }

        // 2nd Priority Match: Parent's email (Column E)
        if (matchedRow === -1 && stripeEmail) {
          for (var j = 0; j < sheetData.length; j++) {
            var rowEmail = sheetData[j][4].toString().toLowerCase().trim();
            if (rowEmail === stripeEmail) {
              matchedRow = j + 2;
              break;
            }
          }
        }
      }

      if (matchedRow !== -1) {
        // Send single instant Zoom link, turn GREEN, and mark Week X Paid
        processPostPaymentAction(sheet, matchedRow, "yes", true);

        return ContentService
          .createTextOutput(JSON.stringify({ result: "success", matched_row: matchedRow, status: "paid_and_zoom_sent" }))
          .setMimeType(ContentService.MimeType.JSON);
      } else {
        Logger.log("Stripe payment received but no matching row found for email: " + stripeEmail + ", child: " + stripeChildName);
        return ContentService
          .createTextOutput(JSON.stringify({ result: "unmatched", email: stripeEmail, child: stripeChildName }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // =========================================================================
    // CASE B: WEBSITE FORM SUBMISSION (Initial Registration)
    // =========================================================================
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var parentName = data.parentName || "";
    var childName = data.childName || "";
    var childSchool = data.childSchool || "";
    var parentEmail = data.parentEmail || "";
    var targetYear = data.targetYear || "";
    var customNote = data.notes || "";
    var sendFurtherEmail = "No";
    var emailStatus = "New Submission - Pending Confirmation";
    var weekNumber = 1;
    var week1Status = "Not Paid";

    // Ensure Column K (Week 1) header exists
    getWeekColumnIndex(sheet, 1);

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
      weekNumber,
      week1Status
    ]);

    var newRow = sheet.getLastRow();
    
    // Style Week 1 cell as Not Paid (Red text)
    sheet.getRange(newRow, 11).setFontColor("#b91c1c").setFontWeight("normal");
    
    // Auto-turn row 🔴 VIVID RED on new submission
    setRowHighlight(sheet, newRow, COLOR_RED_NEW);
    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", row: newRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Logic 11+ Google Sheet Webhook is active!");
}

// --- 11. ON-EDIT TRIGGER ---
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
    
    if (val === "yes" || val === "confirm") {
      processConfirmationAction(sheet, row, "yes");
    } else if (val === "zoom" || val === "send zoom") {
      processPostPaymentAction(sheet, row, "yes");
    } else if (val === "cancel" || val === "cancelled") {
      processCancelWeekAction(sheet, row);
    } else if (val === "reset" || val === "purple") {
      setRowHighlight(sheet, row, COLOR_PURPLE_RESET);
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

