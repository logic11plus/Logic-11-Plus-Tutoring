/**
 * ==============================================================================
 * LOGIC 11+ GOOGLE APPS SCRIPT: DYNAMIC WEEKLY TRACKING & AUTO-PAYMENT
 * ==============================================================================
 * 
 * VIVID COLOR LIFECYCLE:
 * 🔴 VIVID RED (#ff7b72)    : NEW SUBMISSION (No confirmation email sent yet)
 * 🟡 VIVID YELLOW (#ffe066) : CONFIRMATION SENT (Registered, awaiting payment)
 * 🔵 VIVID BLUE (#60a5fa)   : PAID (Payment confirmed in Stripe, Zoom link pending)
 * 🟢 VIVID GREEN (#4ade80)  : ZOOM SENT (Confirmed for this week, active session space)
 * 🟣 VIVID PURPLE (#c084fc) : SUNDAY RESET (Returning student, waiting for next week's payment)
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
 * Col K: Week 1 (Paid / Not Paid)
 * Col L: Week 2 (Paid / Not Paid)
 * Col M: Week 3 (Paid / Not Paid) ... (Added dynamically as students advance)
 */

var STRIPE_PAYMENT_LINK = "https://buy.stripe.com/YOUR_STRIPE_LINK_HERE";
var MAX_CAPACITY_PER_SLOT = 20;

// Vivid Color Hex Palette
var COLOR_RED_NEW       = "#ff7b72"; // Red: New lead, no email yet
var COLOR_YELLOW_CONF   = "#ffe066"; // Yellow: Confirmation sent, awaiting payment
var COLOR_BLUE_PAID     = "#60a5fa"; // Blue: Stripe paid, needs Zoom link
var COLOR_GREEN_ZOOM    = "#4ade80"; // Green: Zoom link sent & confirmed
var COLOR_PURPLE_RESET  = "#c084fc"; // Purple: Sunday reset, returning student awaiting next week's payment

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
    // Stage 2: Mark Paid (Yellow/Purple ➔ Blue)
    .addItem('💳 Mark Selected Row as PAID (Turns 🔵 Blue)', 'markSelectedRowAsPaid')
    .addSeparator()
    // Stage 3: Instant Zoom Link Dispatch (Blue ➔ Green)
    .addItem('🚀 Send Instant Zoom Link (Turns 🟢 Green)', 'sendPostPaymentEmailForActiveRow')
    .addItem('📝 Create Zoom Link Draft', 'createPostPaymentDraftForActiveRow')
    .addSeparator()
    // Stage 4: Returning Students & Weekly Reset
    .addItem('🔄 Run Sunday Weekly Reset (Only Paid Students Turn 🟣 Purple & Add Next Week Column)', 'runWeeklySessionReset')
    .addItem('🔁 Send Continuing Student Payment Email', 'sendContinuingEmailForActiveRow')
    .addItem('📝 Create Continuing Student Draft', 'createContinuingDraftForActiveRow')
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
  var colIndex = 10 + weekNum; // Week 1 -> Col 11 (K), Week 2 -> Col 12 (L), Week 3 -> Col 13 (M)...
  var headerCell = sheet.getRange(1, colIndex);
  var expectedHeader = "Week " + weekNum;
  
  if (!headerCell.getValue() || headerCell.getValue().toString().trim() === "") {
    headerCell.setValue(expectedHeader).setFontWeight("bold").setBackground("#e2e8f0");
  }
  return colIndex;
}

// Helper to mark a specific week as "Paid" or "Not Paid"
function setStudentWeekPaymentStatus(sheet, row, weekNum, isPaid) {
  var col = getWeekColumnIndex(sheet, weekNum);
  var cell = sheet.getRange(row, col);
  if (isPaid) {
    cell.setValue("Paid").setFontColor("#15803d").setFontWeight("bold");
  } else {
    cell.setValue("Not Paid").setFontColor("#b91c1c").setFontWeight("normal");
  }
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
      setStudentWeekPaymentStatus(sheet, rowNum, 1, false); // Ensure Week 1 says "Not Paid"
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

function markSelectedRowAsPaid() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { showAlertSafely("Please click on a row with a parent's details (Row 2 or below)."); return; }
  
  var childName = sheet.getRange(row, 3).getValue();
  var currentWeek = getStudentWeekNumber(sheet, row);
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  
  sheet.getRange(row, 8).setValue("Paid");
  sheet.getRange(row, 9).setValue("Payment Confirmed in Stripe on " + timestamp + " (Week " + currentWeek + ")");
  setStudentWeekPaymentStatus(sheet, row, currentWeek, true); // Mark Week X as "Paid"
  setRowHighlight(sheet, row, COLOR_BLUE_PAID); // Turns VIVID BLUE
  
  showAlertSafely("✅ " + (childName || "Student") + " marked as PAID for Week " + currentWeek + "!\nRow is now 🔵 BLUE. Click 'Send Instant Zoom Link' to email their class pass.");
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
  setStudentWeekPaymentStatus(sheet, row, 1, false);

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
    "Once your payment is received, you will receive your Post-Payment Confirmation email containing your child's live Zoom classroom link and joining passcode.\n\n" +
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
      showAlertSafely("✅ Confirmation email sent to " + parentEmail + "!\nRow is now 🟡 YELLOW (Awaiting Payment).");
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 5. ACTION 2: INSTANT POST-PAYMENT ZOOM DETAILS (BLUE ➔ GREEN) ---
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
      sheet.getRange(row, 8).setValue("Zoom Sent");
      setStudentWeekPaymentStatus(sheet, row, currentWeek, true); // Ensure Week X is marked "Paid"
      setRowHighlight(sheet, row, COLOR_GREEN_ZOOM); // Turns VIVID GREEN
      showAlertSafely("🚀 Instant Zoom link sent to " + parentEmail + " for Week " + currentWeek + "!\nRow is now 🟢 GREEN (Space Confirmed).");
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 6. ACTION 3: CONTINUING CUSTOMER RE-ENROLLMENT EMAIL ---
function processContinuingCustomerAction(sheet, row, actionType) {
  var parentName = sheet.getRange(row, 2).getValue();
  var childName = sheet.getRange(row, 3).getValue();
  var parentEmail = sheet.getRange(row, 5).getValue();
  var targetYear = sheet.getRange(row, 6).getValue().toString();
  var emailStatusCell = sheet.getRange(row, 9);
  
  // Note: Current student week is already set to the upcoming week by the Sunday reset
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
      emailStatusCell.setValue("Continuing Re-enrollment Sent on " + timestamp + " (Week " + targetWeek + ")");
      sheet.getRange(row, 8).setValue("Payment Due (Week " + targetWeek + ")");
      setRowHighlight(sheet, row, COLOR_PURPLE_RESET); // Keeps VIVID PURPLE until paid
      showAlertSafely("✅ Continuing Student re-enrollment notice for Week " + targetWeek + " sent to " + parentEmail + "!");
    }
  } catch (err) {
    emailStatusCell.setValue("Error: " + err.message);
    showAlertSafely("❌ Error: " + err.message);
  }
}

// --- 7. ACTION 4: SUNDAY WEEKLY RESET (ONLY ADVANCES STUDENTS WHO PAID FOR CURRENT WEEK) ---
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

    // Check if student was confirmed / paid / zoom sent for current week
    var isConfirmedForCurrentWeek = (currentSendVal.indexOf("zoom") !== -1 || currentSendVal.indexOf("paid") !== -1);

    if (isConfirmedForCurrentWeek) {
      var nextWeek = currentWeek + 1;
      
      // Update Column J to next week
      sheet.getRange(rowNum, 10).setValue(nextWeek);
      
      // Add Week (X+1) column & mark "Not Paid"
      setStudentWeekPaymentStatus(sheet, rowNum, nextWeek, false);
      
      // Turn Purple (Waiting for Next Week Payment)
      setRowHighlight(sheet, rowNum, COLOR_PURPLE_RESET);
      sheet.getRange(rowNum, 8).setValue("Waiting for Next Week Payment");
      sheet.getRange(rowNum, 9).setValue("Reset for Week " + nextWeek + " (Waiting Payment)");
      count++;
    }
  }

  showAlertSafely("🔄 Sunday Weekly Reset Complete:\n" + count + " active student(s) who completed their session advanced to their next week, received their new Week Column (Not Paid), and turned 🟣 PURPLE.");
}

// --- 8. ACTION 5: CAPACITY TRACKER TAB GENERATOR (20 SEATS PER SLOT) ---
function generateCapacityTrackerTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var capacitySheet = ss.getSheetByName("Capacity Tracker");
  
  if (!capacitySheet) {
    capacitySheet = ss.insertSheet("Capacity Tracker");
  }

  capacitySheet.clear();

  // Set Headers
  capacitySheet.getRange("A1:E1").setValues([[
    "Slot / Cohort", "Max Limit", "Active Confirmed Spaces (Green / Zoom Sent)", "Spaces Remaining", "Status"
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

// --- 9. WEBHOOK: RECEIVE FORM SUBMISSIONS & STRIPE PAYMENT WEBHOOKS ---
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var rawContents = e.postData.contents;
    var data = JSON.parse(rawContents);

    // =========================================================================
    // CASE A: STRIPE WEBHOOK EVENT (checkout.session.completed)
    // =========================================================================
    if (data && data.type && data.type.indexOf("checkout.session") !== -1) {
      var session = data.data.object;
      var stripeEmail = (session.customer_details && session.customer_details.email) ? session.customer_details.email.toLowerCase().trim() : "";
      
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
        var currentWeek = getStudentWeekNumber(sheet, matchedRow);
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

        // Step 1: Record Stripe payment in sheet, mark Week X Paid, & turn BLUE
        sheet.getRange(matchedRow, 8).setValue("Paid in Stripe");
        sheet.getRange(matchedRow, 9).setValue("Stripe Payment Verified on " + timestamp + " (Week " + currentWeek + ")");
        setStudentWeekPaymentStatus(sheet, matchedRow, currentWeek, true); // Mark Week X as "Paid"
        setRowHighlight(sheet, matchedRow, COLOR_BLUE_PAID); // 🔵 Auto-Turn BLUE

        // Step 2: Instantly send Zoom link & turn GREEN
        processPostPaymentAction(sheet, matchedRow, "yes");

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

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", row: newRow }))
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

// --- 10. ON-EDIT TRIGGER ---
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
    } else if (val === "paid") {
      markSelectedRowAsPaid();
    } else if (val === "zoom" || val === "send zoom") {
      processPostPaymentAction(sheet, row, "yes");
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

