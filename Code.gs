const MEDIUMS = {
  english: { label: "English", parentFolder: "RevisionPapers", fileSuffix: "" },
  sinhala: { label: "Sinhala", parentFolder: "RevisionPapers - Sinhala", fileSuffix: " - Sinhala" }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle('Think Like ICT | Paper Automation System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Get the settings for a medium. A blank value counts as English;
 * an unrecognised value returns null so the student is never sent the wrong paper.
 */
function getMediumConfig(value) {
  const key = String(value || "").trim().toLowerCase() || "english";
  return MEDIUMS[key] || null;
}

/**
 * Get the 1-based position of the Medium column, adding the column if it is missing.
 */
function getMediumColumn(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.findIndex(h => String(h).trim().toLowerCase() === "medium");

  if (index !== -1) return index + 1;

  sheet.getRange(1, headers.length + 1).setValue("Medium");
  return headers.length + 1;
}

/**
 * Find a student's next paper in Drive.
 * Returns { file, fileName }, or { error } describing what is missing.
 */
function findPaper(grade, nextWeek, mediumValue) {

  if (!grade || isNaN(nextWeek)) {
    return { error: "Invalid grade/week" };
  }

  const medium = getMediumConfig(mediumValue);

  if (!medium) {
    return { error: "Invalid medium: " + mediumValue };
  }

  const parentFolders =
    DriveApp.getFoldersByName(medium.parentFolder);

  if (!parentFolders.hasNext()) {
    return { error: "Medium folder not found: " + medium.parentFolder };
  }

  const gradeFolders =
    parentFolders.next().getFoldersByName(grade);

  if (!gradeFolders.hasNext()) {
    return { error: `Grade folder not found: ${medium.parentFolder}/${grade}` };
  }

  const fileName =
    `week${String(nextWeek).padStart(2,'0')}${medium.fileSuffix}.pdf`;

  const files =
    gradeFolders.next().getFilesByName(fileName);

  if (!files.hasNext()) {
    return { error: "File not found: " + fileName };
  }

  return { file: files.next(), fileName: fileName };
}

function sendWeeklyPaper() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const studentSheet = ss.getSheetByName("Students");
  const logSheet = ss.getSheetByName("Logs");

  const mediumColumn = getMediumColumn(studentSheet);
  const data = studentSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const name = data[i][0];
    const email = data[i][1];
    const grade = String(data[i][2]).trim();
    const active = String(data[i][3]).trim().toLowerCase();
    let nextWeek = Number(data[i][4]);

    if (active !== "yes") continue;

    const paper = findPaper(grade, nextWeek, data[i][mediumColumn - 1]);

    if (paper.error) {

      logSheet.appendRow([
        new Date(),
        email,
        paper.error
      ]);

      continue;
    }

    const file = paper.file;
    const fileName = paper.fileName;

    GmailApp.sendEmail(
      email,
      `Weekly ${grade} Revision Paper`,
      `Hi ${name},

Here is your week ${nextWeek} revision paper.

- Tharin`,
      {
        attachments: [file.getBlob()]
      }
    );

    // Update NextWeek column
    studentSheet
      .getRange(i + 1, 5)
      .setValue(nextWeek + 1);

    // Update LastSentDate column
    studentSheet
      .getRange(i + 1, 6)
      .setValue(new Date());

    // Log success
    logSheet.appendRow([
      new Date(),
      email,
      "Sent " + fileName
    ]);

  }
  
  return "Successfully sent papers to all active students.";
}

/**
 * Check which paper each active student would get next, without sending anything.
 * Run this from the Apps Script editor and read the execution log.
 */
function previewNextSend() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const mediumColumn = getMediumColumn(sheet);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim().toLowerCase() !== "yes") continue;

    const grade = String(data[i][2]).trim();
    const paper = findPaper(grade, Number(data[i][4]), data[i][mediumColumn - 1]);

    Logger.log(`${data[i][1]} (${grade}): ${paper.error || paper.fileName}`);
  }
}

/**
 * Get dashboard data including summary stats, student list, and recent logs.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName("Students");
  const logSheet = ss.getSheetByName("Logs");

  const mediumColumn = getMediumColumn(studentSheet);
  const studentData = studentSheet.getDataRange().getValues();
  const headers = studentData.shift(); // Remove headers
  
  let totalStudents = studentData.length;
  let activeStudents = 0;
  let inactiveStudents = 0;
  
  const students = studentData.map(row => {
    const isActive = String(row[3]).trim().toLowerCase() === "yes";
    if (isActive) activeStudents++; else inactiveStudents++;

    const medium = getMediumConfig(row[mediumColumn - 1]);

    return {
      name: row[0],
      email: row[1],
      grade: row[2],
      medium: medium ? medium.label : String(row[mediumColumn - 1]).trim(),
      active: isActive,
      nextWeek: row[4],
      lastSent: row[5] ? Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), "MMM dd") : "Never"
    };
  });
  
  const logData = logSheet.getDataRange().getValues();
  logData.shift(); // Remove headers
  const logs = logData.slice(-10).reverse().map(row => ({
    timestamp: Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "MMM dd, hh:mm a"),
    email: row[1],
    message: row[2],
    status: row[2].includes("Sent") ? "Success" : (row[2].includes("Failed") || row[2].includes("not found") || row[2].includes("Invalid") ? "Error" : "Info")
  }));

  // Find last run date from logs
  let lastRun = "None";
  if (logData.length > 0) {
    const lastEntry = logData[logData.length - 1];
    lastRun = Utilities.formatDate(new Date(lastEntry[0]), Session.getScriptTimeZone(), "yyyy-MM-dd hh:mm a");
  }
  
  return {
    stats: {
      total: totalStudents,
      active: activeStudents,
      inactive: inactiveStudents,
      lastRun: lastRun
    },
    students: students,
    logs: logs
  };
}

/**
 * Add a new student to the sheet.
 */
function addStudent(student) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  sheet.appendRow([
    student.name,
    student.email,
    student.grade,
    "Yes", // default active
    student.nextWeek,
    "" // last sent
  ]);
  return "Student added successfully!";
}

/**
 * Update the active status of a student.
 */
function updateStudentStatus(email, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  const data = sheet.getRange("B:B").getValues(); // Check email column (B)
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sheet.getRange(i + 1, 4).setValue(status); // Column D
      return true;
    }
  }
  return false;
}

/**
 * Toggle active status of a student (Legacy).
 */

function getLogoDataUri() {
  const props = PropertiesService.getScriptProperties();
  const fileId = props.getProperty('LOGO_FILE_ID');

  if (!fileId) {
    throw new Error('LOGO_FILE_ID is not set in Script Properties.');
  }

  const file = DriveApp.getFileById(fileId);

  // Key includes the last update so a new logo shows without waiting for the cache
  const cache = CacheService.getScriptCache();
  const cacheKey = 'dashboard_logo_' + fileId + '_' + file.getLastUpdated().getTime();
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const blob = file.getBlob();
  const mimeType = blob.getContentType();
  const base64 = Utilities.base64Encode(blob.getBytes());
  const dataUri = 'data:' + mimeType + ';base64,' + base64;

  // Cache for 6 hours (max 21600 seconds)
  cache.put(cacheKey, dataUri, 21600);

  return dataUri;
}

