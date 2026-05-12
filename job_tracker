// ============================================================
// JOB APPLICATION TRACKER — Google Apps Script
// - Gmail auto-detects new application confirmations → adds rows
// - To close an application: type anything in col E (Status) e.g. "rejected", "interview", "withdrew" — script skips it
// - Thread IDs are used to make the code stop re-adding the same email 
//
// COLUMNS:
//   A = Company | B = Position | C = Date Applied
//   D = Days Pending (auto) | E = Status (you fill) | F = Thread ID
// ============================================================

var props        = PropertiesService.getScriptProperties();
var RECIPIENT    = props.getProperty('EMAIL');
var SHEET_NAME   = props.getProperty('GSHEET');
var TRACKING_START = props.getProperty('BEGINDATE');

var CONFIRMATION_KEYWORDS = [
  "thank you for applying",
  "we have received your application",
  "application confirmation",
  "your application has been received",
  "currently under review",
  "our talent team will be reviewing",
  "successfully submitted",
  "we will review your application",
  "thrilled that you're interested"
];

var REJECTION_KEYWORDS = [
  "unfortunately",
  "unable to move forward",
  "not moving forward",
  "not selected",
  "move on with other candidates",
  "moving forward with other candidates",
  "decided to move forward with",
  "position has been filled",
  "selected another candidate",
  "we will not be",
  "decided not to proceed",
  "regret to inform",
  "not a match",
  "we won't be moving"
];

// ============================================================
// MAIN
// ============================================================
function runDailyJobTracker() {
  syncNewApplicationsFromGmail();
  autoDetectRejections();
  updateDaysPending();
  var applications = getActiveApplications();
  sendSummaryEmail(applications);
}

// ============================================================
// GMAIL SYNC — only adds rows with no existing thread ID match
// ============================================================
function syncNewApplicationsFromGmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var existingThreadIds = getAllThreadIds(sheet);

  var threads = GmailApp.search(
    "(\"thank you for applying\" OR \"application confirmation\" OR \"received your application\" OR \"reviewing your background\" OR \"thrilled that you're interested\") after:" + TRACKING_START,
    0, 50
  );

  threads.forEach(function(thread) {
    var threadId = thread.getId();
    if (existingThreadIds.indexOf(threadId) !== -1) return;

    var messages = thread.getMessages();
    var firstMessage = messages[0];
    var subject = firstMessage.getSubject();
    var body = firstMessage.getPlainBody().toLowerCase();
    var fullText = (subject + " " + body).toLowerCase();

    if (!containsAny(fullText, CONFIRMATION_KEYWORDS)) return;

    var parsed = parseCompanyAndPosition(subject, body, firstMessage.getFrom());
    var dateApplied = Utilities.formatDate(firstMessage.getDate(), "Asia/Bangkok", "yyyy-MM-dd");

    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 6).setValues([[
      parsed.company,
      parsed.position,
      dateApplied,
      0,
      "",       // Status — blank = active
      threadId
    ]]);
    sheet.getRange(newRow, 1, 1, 5).setBackground("#fff9c4"); // yellow = auto-added, verify
  });
}

// ============================================================
// AUTO-REJECT DETECTION
// ============================================================
function autoDetectRejections() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var today = new Date();

  data.forEach(function(row, i) {
    var company   = row[0];
    var dateApplied = row[2];
    var status    = row[4];
    var threadId  = row[5];

    if (!company || !dateApplied || status) return; // skip empty or already closed

    var rowNum = i + 2;
    var days = Math.floor((today - new Date(dateApplied)) / (1000 * 60 * 60 * 24));

    // --- RULE 1: Silent rejection after 30 days ---
    if (days > 30) {
      sheet.getRange(rowNum, 5).setValue("rejected (silent — 30d)");
      sheet.getRange(rowNum, 1, 1, 5).setBackground("#fce4ec");
      return;
    }

    // --- RULE 2: Rejection keyword found in thread replies ---
    if (!threadId) return;

    try {
      var thread = GmailApp.getThreadById(threadId);
      if (!thread) return;

      var messages = thread.getMessages();
      if (messages.length < 2) return; // only confirmation email so far, no reply

      for (var m = 1; m < messages.length; m++) {
        var fullText = messages[m].getSubject().toLowerCase() + " " +
                       messages[m].getPlainBody().toLowerCase();
        if (containsAny(fullText, REJECTION_KEYWORDS)) {
          sheet.getRange(rowNum, 5).setValue("rejected");
          sheet.getRange(rowNum, 1, 1, 5).setBackground("#fce4ec");
          break;
        }
      }
    } catch(e) {
      // Thread inaccessible — skip silently
    }
  });
}

// ============================================================
// UPDATE DAYS PENDING (col D) — active rows only
// ============================================================
function updateDaysPending() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var today = new Date();
  var data = sheet.getRange(2, 3, lastRow - 1, 3).getValues(); // cols C, D, E

  data.forEach(function(row, i) {
    var dateApplied = row[0];
    var status      = row[2];
    if (!dateApplied || status) return;
    var daysPending = Math.floor((today - new Date(dateApplied)) / (1000 * 60 * 60 * 24));
    sheet.getRange(i + 2, 4).setValue(daysPending);
  });
}

// ============================================================
// GET ACTIVE ROWS (Status is blank)
// ============================================================
function getActiveApplications() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var applications = [];

  data.forEach(function(row) {
    var company     = row[0];
    var position    = row[1];
    var dateApplied = row[2];
    var daysPending = row[3];
    var status      = row[4];
    if (!company || !dateApplied || status) return;
    applications.push({
      company: company,
      position: position,
      dateApplied: new Date(dateApplied),
      daysPending: daysPending
    });
  });

  applications.sort(function(a, b) { return a.dateApplied - b.dateApplied; });
  return applications;
}

// ============================================================
// SEND PLAIN EMAIL
// ============================================================
function sendSummaryEmail(apps) {
  var today = Utilities.formatDate(new Date(), "Asia/Bangkok", "d MMM yyyy");
  var lines = [];

  lines.push("Job Tracker — " + today);
  lines.push("Under review: " + apps.length + " application(s)");
  lines.push("---");

  if (apps.length === 0) {
    lines.push("No active applications. Go apply to something.");
  } else {
    apps.forEach(function(app) {
      var dateStr  = Utilities.formatDate(app.dateApplied, "Asia/Bangkok", "d MMM yyyy");
      var dayLabel = app.daysPending === 1 ? "1 day ago" : app.daysPending + " days ago";
      lines.push(app.company + " — " + app.position);
      lines.push("Applied: " + dateStr + " (" + dayLabel + ")");
      lines.push("");
    });
  }

  var subject = "Job Tracker: " + apps.length + " under review — " + today;
  GmailApp.sendEmail(RECIPIENT, subject, lines.join("\n"), { name: "Job Tracker" });
}

// ============================================================
// HELPERS
// ============================================================
function getAllThreadIds(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 6, lastRow - 1, 1).getValues()
    .map(function(r) { return r[0]; })
    .filter(Boolean);
}

function containsAny(text, keywords) {
  return keywords.some(function(k) { return text.indexOf(k.toLowerCase()) !== -1; });
}

function parseCompanyAndPosition(subject, body, sender) {
  var company  = "";
  var position = "";

  var companyMatch = subject.match(/applying to (.+?)(?:\s*[-–|!]|$)/i);
  if (companyMatch) company = companyMatch[1].trim();

  if (!company) {
    var atMatch = subject.match(/ at (.+?)(?:\s*[-–|!]|$)/i);
    if (atMatch) company = atMatch[1].trim();
  }

  if (!company) {
    var domainMatch = sender.match(/@([a-zA-Z0-9-]+)\./);
    if (domainMatch) {
      var domain  = domainMatch[1];
      var generic = ["successfactors","icims","greenhouse","lever","workday","taleo","smartrecruiters","gmail","mail"];
      if (generic.indexOf(domain) === -1) {
        company = domain.charAt(0).toUpperCase() + domain.slice(1);
      }
    }
  }

  if (!company) {
    var bodyCompany = body.match(/(?:applying to|interest in joining|working (?:for|at)) ([A-Z][a-zA-Z0-9\s]+?)(?:\.|,|\n)/);
    if (bodyCompany) company = bodyCompany[1].trim();
  }

  var posMatch = subject.match(/(?:confirmation for|applying for)(?: the)? (.+?)(?:\s*[-–|@]|$)/i);
  if (posMatch) position = posMatch[1].trim();

  if (!position) {
    var bodyPos = body.match(/(?:position|role|job title)[:\s]+([^\n,\.]+)/i);
    if (bodyPos) position = bodyPos[1].trim();
  }

  return {
    company:  company  || "— fill in —",
    position: position || "— fill in —"
  };
}

// ============================================================
// ONE-TIME SETUP
// ============================================================
function setup() {
  createSheet();
  createDailyTrigger();
}

function createSheet() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(SHEET_NAME);

  if (existing) {
    existing.getRange("E1").setValue("Status");
    existing.getRange("F1").setValue("Thread ID");
    existing.getRange("E1:F1").setFontWeight("bold");
    existing.setColumnWidth(5, 120);
    existing.setColumnWidth(6, 180);
    Logger.log("Existing sheet updated.");
    return;
  }

  var sheet = ss.insertSheet(SHEET_NAME);
  sheet.getRange("A1:F1").setValues([["Company Name", "Position", "Date Applied", "Days Pending", "Status", "Thread ID"]]);
  sheet.getRange("A1:F1").setFontWeight("bold");
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 180);
  Logger.log("Sheet created.");
}

function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "runDailyJobTracker") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("runDailyJobTracker")
    .timeBased()
    .everyDays(1)
    .atHour(11) // 11:00 UTC = 18:00 GMT+7
    .create();
  Logger.log("Trigger set: daily at 11:00 UTC (18:00 Bangkok)");
}