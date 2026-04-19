function doGet() {
  return HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle('Think Like ICT | Paper Automation System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sendWeeklyPaper() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const studentSheet = ss.getSheetByName("Students");
  const logSheet = ss.getSheetByName("Logs");

  const data = studentSheet.getDataRange().getValues();

  const parentFolder = DriveApp.getFoldersByName("RevisionPapers").next();

  for (let i = 1; i < data.length; i++) {

    const name = data[i][0];
    const email = data[i][1];
    const grade = String(data[i][2]).trim();
    const active = String(data[i][3]).trim().toLowerCase();
    let nextWeek = Number(data[i][4]);

    if (active !== "yes") continue;

    if (!grade || isNaN(nextWeek)) {

      logSheet.appendRow([
        new Date(),
        email,
        "Invalid grade/week"
      ]);

      continue;
    }

    const gradeFolders =
      parentFolder.getFoldersByName(grade);

    if (!gradeFolders.hasNext()) {

      logSheet.appendRow([
        new Date(),
        email,
        "Grade folder not found: " + grade
      ]);

      continue;
    }

    const gradeFolder = gradeFolders.next();

    const fileName =
      `week${String(nextWeek).padStart(2,'0')}.pdf`;

    const files =
      gradeFolder.getFilesByName(fileName);

    if (!files.hasNext()) {

      logSheet.appendRow([
        new Date(),
        email,
        "File not found: " + fileName
      ]);

      continue;
    }

    const file = files.next();

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
 * Get dashboard data including summary stats, student list, and recent logs.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName("Students");
  const logSheet = ss.getSheetByName("Logs");
  
  const studentData = studentSheet.getDataRange().getValues();
  const headers = studentData.shift(); // Remove headers
  
  let totalStudents = studentData.length;
  let activeStudents = 0;
  let inactiveStudents = 0;
  
  const students = studentData.map(row => {
    const isActive = String(row[3]).trim().toLowerCase() === "yes";
    if (isActive) activeStudents++; else inactiveStudents++;
    
    return {
      name: row[0],
      email: row[1],
      grade: row[2],
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
    status: row[2].includes("Sent") ? "Success" : (row[2].includes("Failed") || row[2].includes("not found") ? "Error" : "Info")
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

  const cache = CacheService.getScriptCache();
  const cacheKey = 'dashboard_logo_data_uri';
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const mimeType = blob.getContentType();
  const base64 = Utilities.base64Encode(blob.getBytes());
  const dataUri = 'data:' + mimeType + ';base64,' + base64;

  // Cache for 6 hours (max 21600 seconds)
  cache.put(cacheKey, dataUri, 21600);

  return dataUri;
}

