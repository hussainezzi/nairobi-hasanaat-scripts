/**
 * HASANAAT SYSTEM - BACKEND ENGINE v3.2
 * 1. Prevents duplicate entries (Same Action + Same Student + Same Day)
 * 2. Supports editable points (Max +/- 5)
 * 3. Hizb (House) Integration - Hizb at Col F (Index 5)
 */

// --- CONFIGURATION ---
const PICTURE_FOLDER_ID = "1zYRHe-DaxWSywZBWufLXNiqpE-slCGvJ";

/**
 * 1. WEB APP ENTRY POINT
 */
function doGet(e) {
  var page = e.parameter.page || 'Index'; 
  return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .setTitle('Hasanaat - MSB')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 2. PICTURE UPLOAD ENGINE
 */
function uploadEvidence(base64Data, logData) {
  try {
    const folder = DriveApp.getFolderById(PICTURE_FOLDER_ID);
    var type = Number(logData.points) >= 0 ? "Tashjee" : "Tanbeeh";
    var safeActionName = logData.action.replace(/[\\\/:*?"<>|]/g, "");
    var newFileName = "Pic_" + logData.its + "_" + logData.name + "_" + type + "_" + safeActionName + ".jpg";

    const parts = base64Data.split(',');
    const contentType = parts[0].split(':')[1].split(';')[0];
    const decodedData = Utilities.base64Decode(parts[1]);
    
    const blob = Utilities.newBlob(decodedData, contentType, newFileName);
    const file = folder.createFile(blob);
    return file.getUrl();
  } catch (e) {
    return "Upload Failed";
  }
}

/**
 * 3. MAIN DATA SUBMISSION
 */
function submitLog(logData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Logs");
  var studentSheet = ss.getSheetByName("Students");
  
  // --- DUPLICATE CHECK ---
  var logs = logSheet.getDataRange().getValues();
  var todayStr = new Date().toDateString();
  
  // Check against Action (Index 6)
  var alreadyExists = logs.some(r => {
    var rowDate = new Date(r[0]).toDateString();
    return rowDate === todayStr && 
           r[1].toString() === logData.its.toString() && 
           r[6].toString().trim() === logData.action.toString().trim();
  });

  if (alreadyExists) return "Error: This action has already been recorded for this student today.";

  // --- IMAGE HANDLING ---
  if (logData.photoData) {
    logData.photoUrl = uploadEvidence(logData.photoData, logData);
  }

  var points = Number(logData.points);
  var shouldNotifyParent = false;

  // --- UPDATED LOGIC GATE ---
  // If checkbox is ticked, always notify (for both + and - points)
  if (logData.notifyParent) {
    shouldNotifyParent = true;
  } 
  // Safety net: Auto-notify on 3rd strike even if checkbox is NOT ticked
  else if (points < 0) {
    var strikeCount = logSheet.getDataRange().getValues().filter(r => 
      r[1].toString() === logData.its.toString() && 
      r[6].toString().trim() === logData.action.toString().trim() && 
      Number(r[7]) < 0
    ).length;
    if (strikeCount >= 2) shouldNotifyParent = true;
  }

  // SAVE TO SPREADSHEET (Index 5 is Hizb)
  logSheet.appendRow([
    new Date(), 
    logData.its, 
    logData.name, 
    logData.classOnly, 
    logData.section, 
    logData.hizb, 
    logData.action, 
    points, 
    logData.teacher, 
    logData.comments, 
    logData.photoUrl || ""
  ]);
// --- EMAIL DISPATCH ---
  var studentData = studentSheet.getDataRange().getValues();
  var student = studentData.find(r => r[1].toString() === logData.its.toString());
  
  if (student) {
    // Indices updated: H (7) is Phone, so Emails start at I (8)
    var recipients = [
      { email: student[8].toString().trim(), role: 'parent' },  // Column I
      { email: student[9].toString().trim(), role: 'teacher' }, // Column J
      { email: student[10].toString().trim(), role: 'hos' },    // Column K
      { email: student[11].toString().trim(), role: 'masool' }, // Column L
      { email: student[12].toString().trim(), role: 'musaid' }  // Column M
    ];

    recipients.forEach(recp => {
      // Check if the email exists and contains an '@' symbol to be safe
      if (recp.email && recp.email.indexOf('@') > -1) {
        if (recp.role !== 'parent' || shouldNotifyParent) {
          sendCustomizedEmail(recp.email, recp.role, logData);
        }
      }
    });
  }
  return "Success!";
}

/**
 * 4. EMAIL GENERATOR (Mimicking your original style)
 */
function sendCustomizedEmail(toEmail, role, data) {
  var isPositive = Number(data.points) > 0;
  var subject = isPositive ? `Celebration: Positive Update for ${data.name}` : `Support Needed: Discipline Update for ${data.name}`;
  
  var greeting = (role === 'parent') ? `Respected Parent of ${data.name},` : `Respected ${role.charAt(0).toUpperCase() + role.slice(1)},`;

  var messageBody = isPositive ? 
    `<p>Wonderful news regarding <b>${data.name}</b>. Earned <b>${data.points} points</b> for <b>${data.action}</b>.</p>` :
    `<p>Partnering with you regarding <b>${data.name}</b>. A record was added for <b>${data.action}</b> (${data.points} points).</p>`;

  var htmlBody = `
    <div style="font-family: sans-serif; max-width: 600px; padding: 20px; border: 1px solid #eee; border-radius: 10px;">
      <div style="text-align: center;">
        <img src=https://res.cloudinary.com/de0cllasz/image/upload/v1771157375/nairobi-removebg-preview_kuxmjp.png" style="width: 80px;">
        <h2 style="color: ${isPositive ? '#198754' : '#d63384'};">Hasanaat Notification</h2>
      </div>
      <p>${greeting}</p>
      ${messageBody}
      <table style="width: 100%; background: #f9f9f9; padding: 10px; border-radius: 5px;">
        <tr><td><b>Activity:</b></td><td>${data.action}</td></tr>
        <tr><td><b>Points:</b></td><td>${data.points}</td></tr>
        <tr><td><b>Teacher:</b></td><td>${data.teacher}</td></tr>
        ${data.hizb ? `<tr><td><b>Hizb:</b></td><td>${data.hizb}</td></tr>` : ''}
      </table>
      <p>Best Regards,<br>Team Hasanaat</p>
    </div>`;

  MailApp.sendEmail({ to: toEmail, subject: subject, htmlBody: htmlBody });
}

/**
 * 4. SECURITY CHECK
 */
function verifyTeacher(passId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Teachers");
  var data = sheet.getDataRange().getValues();
  var teacher = data.find(r => r[1].toString().trim() === passId.toString().trim());
  return teacher ? { success: true, name: teacher[0] } : { success: false };
}



/**
 * 5. DROPDOWN DATA FETCHING
 */
function getActions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Actions");
  var data = sheet.getDataRange().getValues();
  data.shift(); // Remove header
  return data;
}


/**
 * FETCH STUDENT HISTORY & STATS
 * Used by LogView.html to display individual student records
 */
function getStudentLogs(its) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Logs");
    
    if (!logSheet) {
      return { error: "Logs sheet not found" };
    }

    var data = logSheet.getDataRange().getValues();
    data.shift(); // Remove headers (Date, ITS, Name, Class, Section, Hizb, Action, Points...)

    // Filter logs for the specific ITS
    var studentLogs = data.filter(r => r[1] && r[1].toString() === its.toString());
    
    if (studentLogs.length === 0) {
      return { logs: [], studentName: "Not Found", pos: 0, neg: 0, totalCredit: 53 };
    }

    // Initialize stats
    var pos = 0;
    var neg = 0;
    
    // Extract metadata from the first found log entry
    var studentName = studentLogs[0][2]; // Index 2: Name
    var hizbName = studentLogs[0][5] || "General"; // Index 5: Hizb
    
    // Process and format each log entry
    var formattedLogs = studentLogs.map(r => {
      var p = Number(r[7]); // Index 7: Points
      
      // Calculate running totals for the dashboard cards
      if (p > 0) {
        pos += p;
      } else {
        neg += Math.abs(p);
      }
      
      return {
        dateTime: Utilities.formatDate(new Date(r[0]), "GMT+5", "dd-MM-yyyy"), // Index 0: Date
        action: r[6], // Index 6: Action
        points: p
      };
    }).reverse(); // Shows newest entries at the top of the table

    return {
      studentName: studentName,
      hizb: hizbName,
      logs: formattedLogs,
      pos: pos,
      neg: neg,
      totalCredit: (53 + pos - neg) // Base 53 points + Net change
    };

  } catch (e) {
    Logger.log("Error in getStudentLogs: " + e.toString());
    return { error: e.toString() };
  }
}

function getStudentList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  var data = sheet.getDataRange().getValues();
  data.shift(); 
  return data.map(r => ({ 
    its: r[1], 
    name: r[0], 
    class: r[2], 
    section: r[3], 
    hizb: r[4] || "Unassigned" // Column E (Index 4)
  }));
}

/**
 * 6. CLASS-WISE DASHBOARD
 */
function getDashboardData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  var data = sheet.getDataRange().getValues();
  data.shift();

  var classData = {};
  data.forEach(r => {
    var className = r[3] + " " + r[4]; // Class + Section
    var points = Number(r[7]) || 0;   // Index 7 (Column H)
    if (!classData[className]) classData[className] = { tashjee: 0, tanbeeh: 0 };
    if (points > 0) classData[className].tashjee += points;
    else if (points < 0) classData[className].tanbeeh += Math.abs(points);
  });

  var labels = Object.keys(classData).sort();
  return {
    labels: labels,
    tashjee: labels.map(l => classData[l].tashjee),
    tanbeeh: labels.map(l => classData[l].tanbeeh)
  };
}

/**
 * 7. LEADERBOARDS (TOP 10)
 */
function getTopStudents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  var data = sheet.getDataRange().getValues();
  data.shift(); 
  
  var studentScores = {};
  data.forEach(r => {
    var its = r[1], name = r[2], points = Number(r[7]) || 0;
    if (points > 0) {
      if (!studentScores[its]) studentScores[its] = { name: name, points: 0 };
      studentScores[its].points += points;
    }
  });

  return Object.values(studentScores)
    .sort((a, b) => b.points - a.points)
    .slice(0, 10);
}
/**
 * REVAMPED MASTER DASHBOARD DATA
 * Optimized for speed, reliability, and empty-row safety.
 */
function getQuranicDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Logs");
  
  // 1. GATHER DATA & CONSTANTS
  const allData = logSheet.getDataRange().getValues();
  if (allData.length <= 1) return { error: "No data found in Logs" };
  
  const headers = allData.shift(); 
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const thirtyDaysAgo = now.getTime() - (30 * 24 * 60 * 60 * 1000);

  // Containers
  const hizbStats = {};
  const studentStats = {};
  const classStats = {};

  // 2. SINGLE-PASS PROCESSING
  allData.forEach(row => {
    // Index Mapping (Adjust if columns change): 
    // 0:Date, 1:ITS, 2:Name, 3:Class, 4:Sec, 5:Hizb, 7:Points
    const [rawDate, its, name, className, section, hizbRaw, , pointsRaw] = row;
    
    // Safety check: Skip if no ITS or Date
    if (!its || !rawDate) return;

    const dateObj = new Date(rawDate);
    const dateTime = dateObj.getTime();
    const pts = Number(pointsRaw) || 0;
    const hizb = hizbRaw || "Unassigned";
    const fullClassName = `${className} ${section}`.trim();

    // A. Hizb Logic (House Totals & Trends)
    if (!hizbStats[hizb]) hizbStats[hizb] = { total: 0, yesterdayTotal: 0 };
    hizbStats[hizb].total += pts;
    if (dateTime < todayStart) hizbStats[hizb].yesterdayTotal += pts;

    // B. Student Logic (Leaderboard & Shields)
    if (!studentStats[its]) {
      studentStats[its] = { name: name, pos: 0, neg: 0, className: fullClassName, lastTanbeeh: 0 };
    }
    if (pts > 0) studentStats[its].pos += pts;
    else {
      studentStats[its].neg += Math.abs(pts);
      if (dateTime > studentStats[its].lastTanbeeh) studentStats[its].lastTanbeeh = dateTime;
    }

    // C. Class Logic (Class Toppers)
    if (!classStats[fullClassName]) classStats[fullClassName] = { totalPos: 0, students: {} };
    if (pts > 0) {
      classStats[fullClassName].totalPos += pts;
      classStats[fullClassName].students[its] = (classStats[fullClassName].students[its] || 0) + pts;
    }
  });

  // 3. POST-PROCESSING (Sorting & Formatting)

  // Hizb Stats + Mover of the Day
  const finalHizbStats = Object.keys(hizbStats).map(name => {
    const change = hizbStats[name].total - hizbStats[name].yesterdayTotal;
    return {
      name: name,
      points: hizbStats[name].total,
      trend: change > 0 ? 'up' : (change < 0 ? 'down' : 'stable'),
      changeAmount: change
    };
  }).sort((a, b) => b.points - a.points);

  // Top Overall Students
  const topOverall = Object.values(studentStats)
    .sort((a, b) => b.pos - a.pos)
    .slice(0, 10)
    .map(s => ({ name: s.name, totalPos: s.pos }));

  // Clean Sheets (Shields)
  const cleanSheets = Object.values(studentStats)
    .filter(s => s.lastTanbeeh < thirtyDaysAgo && s.pos > 0)
    .slice(0, 12);

  // Class Toppers
  const classToppers = Object.keys(classStats).map(cName => {
    const students = classStats[cName].students;
    const bestIts = Object.keys(students).sort((a, b) => students[b] - students[a])[0];
    return {
      className: cName,
      name: bestIts ? studentStats[bestIts].name : "No Data",
      pts: bestIts ? students[bestIts] : 0
    };
  }).sort((a, b) => a.className.localeCompare(b.className));

  // Weekly Winner (Class with highest total positive points)
  const weeklyWinner = Object.keys(classStats).length > 0 
    ? Object.keys(classStats).sort((a, b) => classStats[b].totalPos - classStats[a].totalPos)[0]
    : "--";

  // Mover of the Day (Hizb with highest change today)
  const mover = [...finalHizbStats].sort((a, b) => b.changeAmount - a.changeAmount)[0];

  return {
    hizbStats: finalHizbStats,
    topOverall: topOverall,
    cleanSheets: cleanSheets,
    classToppers: classToppers,
    weeklyClassWinner: weeklyWinner,
    moverOfDay: mover ? mover.name : "--"
  };
}

/**
 * 9. MONTHLY PARENT REPORTS
 */
function sendMonthlyParentReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Logs");
  var studentSheet = ss.getSheetByName("Students");
  
  var logs = logSheet.getDataRange().getValues();
  var students = studentSheet.getDataRange().getValues();
  students.shift(); // Remove header
  
  var now = new Date();
  var currentMonth = now.getMonth();
  var currentYear = now.getFullYear();

  students.forEach(student => {
    var studentName = student[0];
    var itsId = student[1].toString();
    var parentEmail = student[7]; // Verify index based on your sheet
    
    if (!parentEmail) return;

    var monthlyLogs = logs.filter(row => {
      var logDate = new Date(row[0]);
      return row[1].toString() === itsId && 
             logDate.getMonth() === currentMonth && 
             logDate.getFullYear() === currentYear;
    });

    if (monthlyLogs.length > 0) {
      sendSummaryEmail(parentEmail, studentName, monthlyLogs);
    }
  });
}

function sendSummaryEmail(toEmail, name, logs) {
  var totalPoints = 0;
  var tableRows = "";

  logs.forEach(log => {
    var action = log[6]; // Index 6
    var pts = Number(log[7]); // Index 7
    totalPoints += pts;
    var color = pts >= 0 ? "#198754" : "#d63384";
    
    tableRows += `
      <tr>
        <td style="padding: 8px; border-bottom: 1px solid #eee;">${Utilities.formatDate(new Date(log[0]), "GMT+5", "dd-MM-yyyy")}</td>
        <td style="padding: 8px; border-bottom: 1px solid #eee;">${action}</td>
        <td style="padding: 8px; border-bottom: 1px solid #eee; color: ${color}; font-weight: bold;">${pts}</td>
      </tr>`;
  });

  var htmlBody = `
    <div style="font-family: sans-serif; max-width: 600px; border: 1px solid #ddd; padding: 20px; border-radius: 10px;">
      <div style="text-align: center;">
        <img src="https://res.cloudinary.com/de0cllasz/image/upload/v1771157375/nairobi-removebg-preview_kuxmjp.png" style="width: 70px;">
        <h2 style="color: #0d6efd;">Monthly Progress Report</h2>
        <p><b>Student:</b> ${name}</p>
      </div>
      <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">
        <thead><tr style="background: #f8f9fa; text-align: left;"><th style="padding: 8px;">Date</th><th style="padding: 8px;">Activity</th><th style="padding: 8px;">Points</th></tr></thead>
        <tbody>${tableRows}</tbody>
      </table>
      <div style="margin-top: 20px; padding: 15px; background: #e7f1ff; border-radius: 5px; text-align: center;">
        <h3 style="margin: 0;">Net Balance: ${totalPoints} Points</h3>
      </div>
    </div>`;

  MailApp.sendEmail({ to: toEmail, subject: `Monthly Report: ${name}`, htmlBody: htmlBody });
}

/**
 * 10. EXECUTIVE SUMMARY (HEAD OF SCHOOL)
 */
function sendMonthlyHeadOfSchoolSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Logs");
  var studentSheet = ss.getSheetByName("Students");
  
  var logs = logSheet.getDataRange().getValues();
  var studentData = studentSheet.getDataRange().getValues();
  
  var hosEmail = studentData[1][9]; // Column J
  if (!hosEmail) return;

  var now = new Date();
  var currentMonth = now.getMonth();
  var monthName = Utilities.formatDate(now, "GMT", "MMMM yyyy");

  var monthlyLogs = logs.filter(row => {
    var d = new Date(row[0]);
    return d.getMonth() === currentMonth && d.getFullYear() === now.getFullYear();
  });

  if (monthlyLogs.length === 0) return;

  var studentStats = {}; 
  var classStats = {}; 
  var hizbStats = {};

  monthlyLogs.forEach(log => {
    var its = log[1], name = log[2], className = log[3] + " " + log[4];
    var hizb = log[5] || "Unassigned"; // Column F
    var pts = Number(log[7]);           // Column H
    
    studentStats[its] = studentStats[its] || { name: name, points: 0, class: className };
    studentStats[its].points += pts;

    classStats[className] = classStats[className] || { points: 0 };
    classStats[className].points += pts;

    hizbStats[hizb] = (hizbStats[hizb] || 0) + pts;
  });

  var top10 = Object.values(studentStats).sort((a,b) => b.points - a.points).slice(0, 10);
  var hizbTable = `
    <h3 style="color: #6f42c1; border-bottom: 2px solid #6f42c1; margin-top: 30px;">Hizb Standings</h3>
    <table style="width: 100%; border-collapse: collapse;">
      <tr style="background: #f1f1f1;"><th style="text-align: left; padding: 5px;">Hizb</th><th style="text-align: right; padding: 5px;">Total Points</th></tr>
      ${Object.keys(hizbStats).sort().map(h => `<tr><td style="padding: 5px;">${h}</td><td style="text-align: right; font-weight: bold;">${hizbStats[h]}</td></tr>`).join('')}
    </table>`;

  var htmlBody = `
    <div style="font-family: sans-serif; max-width: 700px; padding: 20px;">
      <h2 style="color: #0d6efd;">Executive Summary - ${monthName}</h2>
      <h3 style="color: #198754; border-bottom: 2px solid #198754;">Top 10 Performers</h3>
      <table style="width: 100%;">${top10.map(s => `<tr><td>${s.name} (${s.class})</td><td style="text-align: right; color: green;">+${s.points}</td></tr>`).join('')}</table>
      ${hizbTable}
      <h3 style="color: #0d6efd; border-bottom: 2px solid #0d6efd; margin-top: 30px;">Class Engagement</h3>
      <table style="width: 100%;">${Object.keys(classStats).sort().map(c => `<tr><td>${c}</td><td style="text-align: right;">${classStats[c].points}</td></tr>`).join('')}</table>
    </div>`;

  MailApp.sendEmail({ to: hosEmail, subject: `HOS Executive Summary - ${monthName}`, htmlBody: htmlBody });
}

/**
 * 11. UTILITIES
 */
function forceAuth() {
  DriveApp.getFolderById(PICTURE_FOLDER_ID);
}




