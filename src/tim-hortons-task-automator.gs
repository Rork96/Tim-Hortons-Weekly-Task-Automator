function setupSettingsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = spreadsheet.getSheetByName("Settings");
  
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet("Settings");
    settingsSheet.getRange("A1").setValue("Tim Hortons Task Manager - Settings and Instructions").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
    settingsSheet.getRange("A2").setValue("Instructions:").setFontWeight("bold").setFontSize(12);
    settingsSheet.getRange("A3:A11").setValues([
      ["1. This sheet is for managing the list of managers and their email addresses."],
      ["2. Update the 'Manager Name' and 'Email Address' columns below with the correct information."],
      ["3. To skip sending an email to a manager, enter '-' in the 'Email Address' column."],
      ["4. To send tasks, go to the 'Tasks' menu in the top bar and select 'Preview & Send Tasks'."],
      ["5. To test emails, select 'Tasks' > 'Send Test Emails' and enter an email address."],
      ["6. Emails will be sent directly to the managers' email addresses listed below."],
      ["7. To add announcements for all managers, edit the 'Announcements' sheet."],
      ["8. To view logs of all actions, select 'Tasks' > 'View Logs'."],
      ["9. For help, contact Pavlo: Email: pavlo.tsyhanash@gmail.com | Telegram: https://t.me/R_O_R_K"]
    ]);
    settingsSheet.getRange("A13:B13").setValues([["Manager Name", "Email Address"]]).setFontWeight("bold").setBackground("#4A2C2A").setFontColor("white");
    settingsSheet.getRange("A14:B23").setValues([
      ["Alex", "alex@example.com"],
      ["Carys", "carys@example.com"],
      ["Clara", "clara@example.com"],
      ["Dario", "dario@example.com"],
      ["Debbie", "debbie@example.com"],
      ["Diksha", "diksha@example.com"],
      ["Harneet", "harneet@example.com"],
      ["Moksha", "moksha@example.com"],
      ["Pratima", "pratima@example.com"],
      ["Lailene", "lailene@example.com"]
    ]);
    settingsSheet.getRange("A3:A11").setWrap(true);
    settingsSheet.setColumnWidth(1, 300);
    settingsSheet.setColumnWidth(2, 300);
  }
}

function setupAnnouncementsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let announcementsSheet = spreadsheet.getSheetByName("Announcements");
  
  if (!announcementsSheet) {
    announcementsSheet = spreadsheet.insertSheet("Announcements");
    announcementsSheet.getRange("A1").setValue("Tim Hortons Task Manager - Announcements").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center").setBackground("#4A2C2A").setFontColor("white");
    announcementsSheet.getRange("A2").setValue("Enter the latest announcements below. This text will be included in all emails sent to managers. For help, contact Pavlo: Email: pavlo.tsyhanash@gmail.com | Telegram: https://t.me/R_O_R_K").setFontStyle("italic").setWrap(true);
    announcementsSheet.getRange("A3").setValue("No announcements at this time.").setWrap(true);
    announcementsSheet.setColumnWidth(1, 600);
  }
}

function setupLogsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let logsSheet = spreadsheet.getSheetByName("Logs");
  
  if (!logsSheet) {
    logsSheet = spreadsheet.insertSheet("Logs");
    logsSheet.getRange("A1:C1").setValues([["Timestamp", "Action", "Details"]]).setFontWeight("bold").setBackground("#4A2C2A").setFontColor("white");
    logsSheet.setColumnWidth(1, 200);
    logsSheet.setColumnWidth(2, 200);
    logsSheet.setColumnWidth(3, 400);
  }
  return logsSheet;
}

function logAction(action, details) {
  const logsSheet = setupLogsSheet();
  const timestamp = new Date().toISOString();
  logsSheet.appendRow([timestamp, action, details]);
}

function setupMissingSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    "Leaders Board",
    "Self-Evaluations",
    "Calibration & Receiving Logs",
    "Weekly Reports - Section Assignments",
    "CLUSTER - Training Videos",
    "Supervisor Task Assignments",
    "Deadlines Submissions"
  ];

  requiredSheets.forEach(sheetName => {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.getRange("A1").setValue(`Placeholder for ${sheetName}`).setFontWeight("bold");
      logAction("Setup Missing Sheet", `Created missing sheet: ${sheetName}`);
    }
  });
}

function getAnnouncements() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const announcementsSheet = spreadsheet.getSheetByName("Announcements");
  if (!announcementsSheet) {
    setupAnnouncementsSheet();
    return "No announcements at this time.";
  }
  const announcementText = announcementsSheet.getRange("A3").getValue().trim();
  return announcementText || "No announcements at this time.";
}

function showAnnouncements() {
  const ui = SpreadsheetApp.getUi();
  const announcementText = getAnnouncements();
  ui.alert("Announcements", announcementText, ui.ButtonSet.OK);
}

function showLogs() {
  const ui = SpreadsheetApp.getUi();
  const logsSheet = setupLogsSheet();
  const logsData = logsSheet.getDataRange().getValues();
  let logsText = "Logs:\n\n";
  
  for (let i = 1; i < logsData.length; i++) {
    const [timestamp, action, details] = logsData[i];
    logsText += `${timestamp} | ${action} | ${details}\n`;
  }
  
  if (logsData.length <= 1) {
    logsText += "No logs available.";
  }
  
  ui.alert("Logs", logsText, ui.ButtonSet.OK);
}

function getManagersAndEmails() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = spreadsheet.getSheetByName("Settings");
  if (!settingsSheet) {
    throw new Error("Settings sheet not found. Please run setupSettingsSheet() first.");
  }

  const data = settingsSheet.getDataRange().getValues();
  let managers = [];
  let emailList = {};
  let errors = [];

  for (let row = 13; row < data.length; row++) {
    let manager = data[row][0] ? data[row][0].trim() : "";
    let email = data[row][1] ? data[row][1].trim() : "";
    
    if (manager === "" && email === "") continue;
    if (manager === "Manager Name" || email === "Email Address") continue;

    if (!manager) {
      errors.push(`Row ${row + 1} in Settings: Manager Name is empty.`);
      continue;
    }
    if (!email) {
      errors.push(`Row ${row + 1} in Settings: Email Address is empty for manager ${manager}.`);
      continue;
    }
    if (email !== "-" && !email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
      errors.push(`Row ${row + 1} in Settings: Invalid email address for manager ${manager}: ${email}`);
      continue;
    }
    if (managers.includes(manager)) {
      errors.push(`Row ${row + 1} in Settings: Duplicate manager name found: ${manager}`);
      continue;
    }

    managers.push(manager);
    emailList[manager] = email;
  }

  if (errors.length > 0) {
    throw new Error(errors.join("\n"));
  }

  return { managementList: managers, emailList };
}

function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  const instructions = [
    "Welcome to the Tim Hortons Task Manager!",
    "",
    "Instructions:",
    "1. Go to the 'Settings' sheet to update the list of managers and their email addresses.",
    "2. To skip sending an email to a manager, enter '-' in the 'Email Address' column.",
    "3. To send tasks, go to the 'Tasks' menu and select 'Preview & Send Tasks'.",
    "4. To test emails, select 'Tasks' > 'Send Test Emails' and enter an email address.",
    "5. Emails will be sent directly to the managers' email addresses.",
    "6. To add announcements, edit the 'Announcements' sheet.",
    "7. For help, contact Pavlo: Email: pavlo.tsyhanash@gmail.com | Telegram: https://t.me/R_O_R_K"
  ].join("\n");
  ui.alert("Instructions", instructions, ui.ButtonSet.OK);
}

function collectTasks() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  setupMissingSheets();
  const sheets = spreadsheet.getSheets();
  
  let tasksByPerson = {};
  const { managementList, emailList } = getManagersAndEmails();

  const daysOfWeek = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  
  const today = new Date();
  const dayOfWeek = today.getDay();
  const daysSinceMonday = (dayOfWeek === 0 ? 6 : dayOfWeek - 1);
  const mondayThisWeek = new Date(today);
  mondayThisWeek.setDate(today.getDate() - daysSinceMonday);

  const week1Start = new Date(mondayThisWeek);
  const week1End = new Date(mondayThisWeek);
  week1End.setDate(mondayThisWeek.getDate() + 6);
  
  const week2Start = new Date(mondayThisWeek);
  week2Start.setDate(mondayThisWeek.getDate() + 7);
  const week2End = new Date(mondayThisWeek);
  week2End.setDate(mondayThisWeek.getDate() + 13);

  const formatDate = (date) => `${months[date.getMonth()]} ${date.getDate()}`;
  const week1Label = `${formatDate(week1Start)} - ${formatDate(week1End)}`;
  const week2Label = `${formatDate(week2Start)} - ${formatDate(week2End)}`;

  const dayToDateWeek1 = {};
  const dayToDateWeek2 = {};
  daysOfWeek.forEach((day, index) => {
    const week1Date = new Date(week1Start);
    week1Date.setDate(week1Start.getDate() + index);
    const week2Date = new Date(week2Start);
    week2Date.setDate(week2Start.getDate() + index);
    dayToDateWeek1[day] = `${day}, ${formatDate(week1Date)}`;
    dayToDateWeek2[day] = `${day}, ${formatDate(week2Date)}`;
  });

  let unknownNames = new Set();
  let warnings = [];

  const emailQuota = MailApp.getRemainingDailyQuota();
  const emailsToSend = managementList.filter(person => emailList[person] !== "-").length + (warnings.length > 0 ? 1 : 0);
  logAction("Check Email Quota", `Remaining email quota: ${emailQuota}, needed: ${emailsToSend}`);
  if (emailQuota < emailsToSend) {
    const errorMsg = `Error: Email quota exceeded. Remaining quota: ${emailQuota}, needed: ${emailsToSend}. Please wait until the quota resets (00:00 UTC).`;
    logAction("Quota Error", errorMsg);
    GmailApp.sendEmail(
      "pavlo.tsyhanash@gmail.com",
      "Script Error: Email Quota Exceeded",
      errorMsg
    );
    return { tasksByPerson, emailList, dayToDateWeek1, dayToDateWeek2, managementList };
  }

  const leadersSheet = sheets.find(sheet => sheet.getName() === "Leaders Board");
  const leadersData = leadersSheet.getDataRange().getValues();
  logAction("Process Sheet", `Leaders Board sheet has ${leadersData.length} rows`);

  const weeks = [];
  let currentWeek = null;
  for (let row = 0; row < leadersData.length; row++) {
    const cellA = leadersData[row][0] ? leadersData[row][0].toString().trim() : "";
    if (cellA.match(/^[A-Za-z]{3}\s\d{1,2}\s-\s[A-Za-z]{3}\s\d{1,2}$/)) {
      currentWeek = { label: cellA, startRow: row + 1 };
      weeks.push(currentWeek);
    } else if (cellA === "" && currentWeek) {
      currentWeek.endRow = row - 1;
      currentWeek = null;
    }
  }
  if (currentWeek && !currentWeek.endRow) {
    currentWeek.endRow = leadersData.length - 1;
  }

  if (weeks.length < 2) {
    let errorMsg = `The Leaders Board sheet does not have enough weeks. Expected at least 2 weeks, but found ${weeks.length}.`;
    logAction("Validation Error", errorMsg);
    warnings.push(errorMsg);
  }

  for (let weekIndex = 0; weekIndex < Math.min(weeks.length, 2); weekIndex++) {
    const week = weeks[weekIndex];
    const weekLabel = weekIndex === 0 ? week1Label : week2Label;
    const dayToDate = weekIndex === 0 ? dayToDateWeek1 : dayToDateWeek2;

    for (let row = week.startRow, dayIndex = 0; row <= week.endRow; row++, dayIndex++) {
      if (row >= leadersData.length || !leadersData[row] || !leadersData[row][0]) {
        warnings.push(`Row ${row + 1} in Leaders Board (Week ${weekIndex + 1}) is missing. Expected day: ${daysOfWeek[dayIndex]}`);
        continue;
      }
      let day = leadersData[row][0] ? leadersData[row][0].trim().toLowerCase() : "";
      if (!day) {
        warnings.push(`Row ${row + 1} in Leaders Board (Week ${weekIndex + 1}) has empty day. Expected: ${daysOfWeek[dayIndex]}`);
        continue;
      }
      day = day.charAt(0).toUpperCase() + day.slice(1);
      if (day !== daysOfWeek[dayIndex]) {
        warnings.push(`Row ${row + 1} in Leaders Board (Week ${weekIndex + 1}) has incorrect day. Expected: ${daysOfWeek[dayIndex]}, Found: ${day}`);
        continue;
      }
      logAction("Process Leaders Board", `Week ${weekIndex + 1}, Row ${row + 1}: Day = ${day}`);
      for (let col = 1; col < 5; col++) {
        let person = leadersData[row][col] ? leadersData[row][col].trim() : "-";
        let role = leadersData[0][col] || "Unknown Role";
        if (person && person !== "-") {
          let matchedPerson = Object.keys(emailList).find(key => key.toLowerCase() === person.toLowerCase());
          if (matchedPerson) {
            if (!tasksByPerson[matchedPerson]) tasksByPerson[matchedPerson] = [];
            tasksByPerson[matchedPerson].push(`Leaders Board (${weekLabel}): ${role} on ${day}`);
            logAction("Assign Task", `Added task for ${matchedPerson}: Leaders Board (${weekLabel}): ${role} on ${day}`);
          } else {
            if (managementList.some(manager => manager.toLowerCase() === person.toLowerCase())) {
              unknownNames.add(person);
              logAction("Warning", `Person ${person} not found in emailList`);
            }
          }
        }
      }
    }
  }

  if (unknownNames.size > 0) {
    let warningMsg = `The following names in the Leaders Board sheet were not found in the email list: ${Array.from(unknownNames).join(", ")}. Please check the spelling or add them to the email list in the Settings sheet.`;
    warnings.push(warningMsg);
  }

  const selfEvalSheet = sheets.find(sheet => sheet.getName() === "Self-Evaluations");
  const selfEvalData = selfEvalSheet.getDataRange().getValues();
  for (let row = 1; row < 4; row++) {
    if (row >= selfEvalData.length) {
      warnings.push(`Row ${row + 1} in Self-Evaluations (Week 1) is missing.`);
      continue;
    }
    let person = selfEvalData[row][1] ? selfEvalData[row][1].trim() : "-";
    if (person === "-") {
      logAction("Skip Row", `Row ${row + 1} in Self-Evaluations (Week 1) was empty; set to "-".`);
      continue;
    }
    let type = selfEvalData[row][0] || "Unknown Type";
    let matchedPerson = Object.keys(emailList).find(key => key.toLowerCase() === person.toLowerCase());
    if (matchedPerson) {
      if (!tasksByPerson[matchedPerson]) tasksByPerson[matchedPerson] = [];
      tasksByPerson[matchedPerson].push(`Self-Evaluations (${week1Label}): ${type}`);
    } else if (person && person !== "-") {
      if (managementList.some(manager => manager.toLowerCase() === person.toLowerCase())) {
        unknownNames.add(person);
      }
    }
  }
  for (let row = 5; row < 8; row++) {
    if (row >= selfEvalData.length) {
      warnings.push(`Row ${row + 1} in Self-Evaluations (Week 2) is missing.`);
      continue;
    }
    let person = selfEvalData[row][1] ? selfEvalData[row][1].trim() : "-";
    if (person === "-") {
      logAction("Skip Row", `Row ${row + 1} in Self-Evaluations (Week 2) was empty; set to "-".`);
      continue;
    }
    let type = selfEvalData[row][0] || "Unknown Type";
    let matchedPerson = Object.keys(emailList).find(key => key.toLowerCase() === person.toLowerCase());
    if (matchedPerson) {
      if (!tasksByPerson[matchedPerson]) tasksByPerson[matchedPerson] = [];
      tasksByPerson[matchedPerson].push(`Self-Evaluations (${week2Label}): ${type}`);
    } else if (person && person !== "-") {
      if (managementList.some(manager => manager.toLowerCase() === person.toLowerCase())) {
        unknownNames.add(person);
      }
    }
  }

  const calibSheet = sheets.find(sheet => sheet.getName() === "Calibration & Receiving Logs");
  const calibData = calibSheet.getDataRange().getValues();
  for (let row = 1; row < 8; row++) {
    if (row >= calibData.length) {
      warnings.push(`Row ${row + 1} in Calibration & Receiving Logs (Week 1) is missing.`);
      continue;
    }
    let person = calibData[row][1] ? calibData[row][1].trim() : "-";
    if (person === "-") {
      logAction("Skip Row", `Row ${row + 1} in Calibration & Receiving Logs (Week 1) was empty; set to "-".`);
      continue;
    }
    let type = calibData[row][0] || "Unknown Type";
    let matchedPerson = Object.keys(emailList).find(key => key.toLowerCase() === person.toLowerCase());
    if (matchedPerson) {
      if (!tasksByPerson[matchedPerson]) tasksByPerson[matchedPerson] = [];
      tasksByPerson[matchedPerson].push(`Calibration & Receiving Logs (${week1Label}): ${type}`);
    } else if (person && person !== "-") {
      if (managementList.some(manager => manager.toLowerCase() === person.toLowerCase())) {
        unknownNames.add(person);
      }
    }
  }
  for (let row = 9; row < 16; row++) {
    if (row >= calibData.length) {
      warnings.push(`Row ${row + 1} in Calibration & Receiving Logs (Week 2) is missing.`);
      continue;
    }
    let person = calibData[row][1] ? calibData[row][1].trim() : "-";
    if (person === "-") {
      logAction("Skip Row", `Row ${row + 1} in Calibration & Receiving Logs (Week 2) was empty; set to "-".`);
      continue;
    }
    let type = calibData[row][0] || "Unknown Type";
    let matchedPerson = Object.keys(emailList).find(key => key.toLowerCase() === person.toLowerCase());
    if (matchedPerson) {
      if (!tasksByPerson[matchedPerson]) tasksByPerson[matchedPerson] = [];
      tasksByPerson[matchedPerson].push(`Calibration & Receiving Logs (${week2Label}): ${type}`);
    } else if (person && person !== "-") {
      if (managementList.some(manager => manager.toLowerCase() === person.toLowerCase())) {
        unknownNames.add(person);
      }
    }
  }

  const week1StartMarch = new Date(week1Start);
  week1StartMarch.setDate(week1Start.getDate() - 21);
  const week1EndMarch = new Date(week1StartMarch);
  week1EndMarch.setDate(week1StartMarch.getDate() + 6);
  const week2StartApril = new Date(week2Start);
  const week2EndApril = new Date(week2End);
  const marchLabel = `${formatDate(week1StartMarch)} - ${formatDate(week1EndMarch)}`;
  const aprilLabel = `${formatDate(week2StartApril)} - ${formatDate(week2EndApril)}`;

  const weeklyReportsSheet = sheets.find(sheet => sheet.getName() === "Weekly Reports - Section Assignments");
  const weeklyReportsData = weeklyReportsSheet.getDataRange().getValues();
  for (let row = 1; row < 8; row++) {
    if (row >= weeklyReportsData.length) {
      warnings.push(`Row ${row + 1} in Weekly Reports - Section Assignments (Week 1) is missing.`);
      continue;
    }
    let person = weeklyReportsData[row][1] ? weeklyReportsData[row][1].trim() : "-";
    if (person === "-") {
      logAction("Skip Row", `Row ${row + 1} in Weekly Reports - Section Assignments (Week 1) was empty; set to "-".`);
      continue;
    }
    let section = weeklyReportsData[row][0] || "Unknown Section";
    let matchedPerson = Object.keys(emailList).find(key => key.toLowerCase() === person.toLowerCase());
    if (matchedPerson) {
      if (!tasksByPerson[matchedPerson]) tasksByPerson[matchedPerson] = [];
      tasksByPerson[matchedPerson].push(`Weekly Reports - Section Assignments (April, starting ${marchLabel}): ${section}`);
    } else if (person && person !== "-") {
      if (managementList.some(manager => manager.toLowerCase() === person.toLowerCase())) {
        unknownNames.add(person);
      }
    }
  }
  for (let row = 9; row < 16; row++) {
    if (row >= weeklyReportsData.length) {
      warnings.push(`Row ${row + 1} in Weekly Reports - Section Assignments (Week 2) is missing.`);
      continue;
    }
    let person = weeklyReportsData[row][1] ? weeklyReportsData[row][1].trim() : "-";
    if (person === "-") {
      logAction("Skip Row", `Row ${row + 1} in Weekly Reports - Section Assignments (Week 2) was empty; set to "-".`);
      continue;
    }
    let section = weeklyReportsData[row][0] || "Unknown Section";
    let matchedPerson = Object.keys(emailList).find(key => key.toLowerCase() === person.toLowerCase());
    if (matchedPerson) {
      if (!tasksByPerson[matchedPerson]) tasksByPerson[matchedPerson] = [];
      tasksByPerson[matchedPerson].push(`Weekly Reports - Section Assignments (May, starting ${aprilLabel}): ${section}`);
    } else if (person && person !== "-") {
      if (managementList.some(manager => manager.toLowerCase() === person.toLowerCase())) {
        unknownNames.add(person);
      }
    }
  }

  const clusterSheet = sheets.find(sheet => sheet.getName() === "CLUSTER - Training Videos");
  const clusterData = clusterSheet.getDataRange().getValues();
  for (let col = 0; col < clusterData[0].length; col++) {
    let person = clusterData[0][col] ? clusterData[0][col].trim() : "-";
    let matchedPerson = Object.keys(emailList).find(key => key.toLowerCase() === person.toLowerCase());
    let trainees = [];
    for (let row = 1; row < clusterData.length; row++) {
      let trainee = clusterData[row][col] ? clusterData[row][col].trim() : "";
      if (trainee) trainees.push(trainee);
    }
    if (matchedPerson) {
      if (!tasksByPerson[matchedPerson]) tasksByPerson[matchedPerson] = [];
      tasksByPerson[matchedPerson].push(`CLUSTER - Training Videos: Responsible for ${trainees.join(", ")}`);
    } else if (person && person !== "-") {
      if (managementList.some(manager => manager.toLowerCase() === person.toLowerCase())) {
        unknownNames.add(person);
      }
    }
  }

  const supervisorSheet = sheets.find(sheet => sheet.getName() === "Supervisor Task Assignments");
  const supervisorData = supervisorSheet.getDataRange().getValues();
  logAction("Process Sheet", `Supervisor Task Assignments sheet has ${supervisorData.length} rows`);
  
  const headerKeywords = ["person", "assign to", "leaders board - planning"];
  
  for (let row = 2; row <= 4; row++) {
    if (row >= supervisorData.length || !supervisorData[row] || !supervisorData[row][0] || supervisorData[row][0].trim() === "") {
      logAction("Skip Row", `Row ${row + 1} in Supervisor Task Assignments (Daily Entry) is empty; skipping.`);
      continue;
    }
    let person = supervisorData[row][0].trim();
    logAction("Process Supervisor Tasks", `Processing Supervisor Task Assignments (Daily Entry), Row ${row + 1}: Person = ${person}`);
    if (headerKeywords.includes(person.toLowerCase()) || person === "") {
      warnings.push(`Row ${row + 1} in Supervisor Task Assignments (Daily Entry) contains invalid Person value: "${person}".`);
      continue;
    }
    let matchedPerson = Object.keys(emailList).find(key => key.toLowerCase() === person.toLowerCase());
    let timeCards = supervisorData[row][1] || "Not specified";
    let paperWaste = supervisorData[row][2] || "Not specified";
    if (matchedPerson) {
      if (!tasksByPerson[matchedPerson]) tasksByPerson[matchedPerson] = [];
      tasksByPerson[matchedPerson].push(`Supervisor Task Assignments - Daily Entry: Time-Cards Update: ${timeCards}`);
      tasksByPerson[matchedPerson].push(`Supervisor Task Assignments - Daily Entry: Paper Waste Submission: ${paperWaste}`);
      logAction("Assign Task", `Added tasks for ${matchedPerson}: Time-Cards Update: ${timeCards}, Paper Waste Submission: ${paperWaste}`);
    } else if (person && person !== "-") {
      if (managementList.some(manager => manager.toLowerCase() === person.toLowerCase())) {
        unknownNames.add(person);
        logAction("Warning", `Person ${person} not found in emailList`);
      }
    }
  }

  for (let row = 7; row <= 9; row++) {
    if (row >= supervisorData.length || !supervisorData[row] || !supervisorData[row][1] || supervisorData[row][1].trim() === "") {
      logAction("Skip Row", `Row ${row + 1} in Supervisor Task Assignments (Leaders Board - Planning) is empty; skipping.`);
      continue;
    }
    let type = supervisorData[row][0] || "Not specified";
    let person = supervisorData[row][1].trim();
    logAction("Process Supervisor Tasks", `Processing Supervisor Task Assignments (Leaders Board - Planning), Row ${row + 1}: Person = ${person}`);
    if (headerKeywords.includes(person.toLowerCase()) || person === "") {
      warnings.push(`Row ${row + 1} in Supervisor Task Assignments (Leaders Board - Planning) contains invalid Person value: "${person}".`);
      continue;
    }
    let matchedPerson = Object.keys(emailList).find(key => key.toLowerCase() === person.toLowerCase());
    if (matchedPerson) {
      if (!tasksByPerson[matchedPerson]) tasksByPerson[matchedPerson] = [];
      tasksByPerson[matchedPerson].push(`Supervisor Task Assignments - Leaders Board - Planning: ${type}`);
      logAction("Assign Task", `Added task for ${matchedPerson}: Leaders Board - Planning: ${type}`);
    } else if (person && person !== "-") {
      if (managementList.some(manager => manager.toLowerCase() === person.toLowerCase())) {
        unknownNames.add(person);
        logAction("Warning", `Person ${person} not found in emailList`);
      }
    }
  }

  const deadlinesSheet = sheets.find(sheet => sheet.getName() === "Deadlines Submissions");
  const deadlinesData = deadlinesSheet.getDataRange().getValues();
  const deadlines = [];
  for (let row = 1; row < deadlinesData.length; row++) {
    let type = deadlinesData[row][0] || "Unknown Deadline";
    let details = deadlinesData[row][1] || "No details";
    deadlines.push(`Deadlines Submissions: ${type}: ${details}`);
  }
  for (let person in tasksByPerson) {
    tasksByPerson[person].push(...deadlines);
  }

  if (unknownNames.size > 0) {
    let warningMsg = `The following names were not found in the email list: ${Array.from(unknownNames).join(", ")}. Please check the spelling or add them to the email list in the Settings sheet.`;
    warnings.push(warningMsg);
  }

  if (warnings.length > 0) {
    GmailApp.sendEmail(
      "pavlo.tsyhanash@gmail.com",
      "Script Warnings",
      warnings.join("\n\n")
    );
    logAction("Send Warning Email", `Sent warning email with messages: ${warnings.join("; ")}`);
  }

  return { tasksByPerson, emailList, dayToDateWeek1, dayToDateWeek2, managementList };
}

function previewTasks() {
  const ui = SpreadsheetApp.getUi();
  try {
    const { tasksByPerson, emailList, dayToDateWeek1, dayToDateWeek2, managementList } = collectTasks();
    let previewText = "Tasks Preview:\n\n";

    for (let person in tasksByPerson) {
      if (!managementList.includes(person)) continue;
      let email = emailList[person];
      previewText += `To: ${person} (${email})\n`;
      previewText += "Tasks:\n";
      for (let task of tasksByPerson[person]) {
        previewText += `- ${task}\n`;
      }
      previewText += "\n";
    }

    const response = ui.alert(
      'Preview Tasks Before Sending',
      previewText,
      ui.ButtonSet.YES_NO
    );

    if (response == ui.Button.YES) {
      sendEmails(tasksByPerson, emailList, dayToDateWeek1, dayToDateWeek2, managementList, false);
      ui.alert('Success', 'Emails have been sent to the managers!', ui.ButtonSet.OK);
    } else {
      ui.alert('Cancelled', 'You can now edit the sheets and try again.', ui.ButtonSet.OK);
    }
  } catch (e) {
    logAction("Error in Preview Tasks", `Error: ${e.message}`);
    ui.alert('Error', `An error occurred: ${e.message}\nPlease check the 'Logs' sheet or contact Pavlo for assistance.`, ui.ButtonSet.OK);
  }
}

function sendTestEmails() {
  const ui = SpreadsheetApp.getUi();
  try {
    const testEmailPrompt = ui.prompt(
      'Send Test Emails',
      'Enter the email address where all test emails should be sent:',
      ui.ButtonSet.OK_CANCEL
    );

    if (testEmailPrompt.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Cancelled', 'Test email sending was cancelled.', ui.ButtonSet.OK);
      return;
    }

    const testEmail = testEmailPrompt.getResponseText().trim();
    if (!testEmail.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
      ui.alert('Error', 'Invalid email address. Please enter a valid email.', ui.ButtonSet.OK);
      return;
    }

    const { tasksByPerson, emailList, dayToDateWeek1, dayToDateWeek2, managementList } = collectTasks();
    sendEmails(tasksByPerson, emailList, dayToDateWeek1, dayToDateWeek2, managementList, true, testEmail);
    ui.alert('Success', `Test emails have been sent to ${testEmail}!`, ui.ButtonSet.OK);
  } catch (e) {
    logAction("Error in Send Test Emails", `Error: ${e.message}`);
    ui.alert('Error', `An error occurred: ${e.message}\nPlease check the 'Logs' sheet or contact Pavlo for assistance.`, ui.ButtonSet.OK);
  }
}

function autoSendTasks() {
  try {
    const { tasksByPerson, emailList, dayToDateWeek1, dayToDateWeek2, managementList } = collectTasks();
    sendEmails(tasksByPerson, emailList, dayToDateWeek1, dayToDateWeek2, managementList, false);
    logAction("Auto Send Complete", "Emails have been sent to the managers.");
  } catch (e) {
    logAction("Error in Auto Send", `Error: ${e.message}`);
    GmailApp.sendEmail(
      "pavlo.tsyhanash@gmail.com",
      "Script Error in Auto Send",
      `An error occurred while running autoSendTasks: ${e.message}`
    );
  }
}

function sendEmails(tasksByPerson, emailList, dayToDateWeek1, dayToDateWeek2, managementList, isTestMode = false, testEmail = null) {
  const periodLabel = `${Object.values(dayToDateWeek1)[0].split(",")[1].trim()} - ${Object.values(dayToDateWeek2)[6].split(",")[1].trim()}`;
  
  for (let person in tasksByPerson) {
    if (!managementList.includes(person)) continue;
    let originalEmail = emailList[person];
    
    if (originalEmail === "-") {
      logAction("Skip Email", `Skipping email for ${person} because email address is set to "-".`);
      continue;
    }

    let subject = isTestMode ? `[TEST] Tasks for ${person} (${periodLabel})` : `Tasks for ${person} (${periodLabel})`;
    
    let sections = {};
    for (let task of tasksByPerson[person]) {
      let section = task.split(":")[0];
      if (!sections[section]) sections[section] = [];
      sections[section].push(task.replace(`${section}: `, ""));
    }

    let leadersBoardHtml = "";
    if (sections["Leaders Board"]) {
      const daysOfWeek = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
      let week1Tasks = {};
      let week2Tasks = {};
      daysOfWeek.forEach(day => {
        week1Tasks[day] = { "Morning Chart": false, "Morning 2pm Post": false, "Evening Chart": false, "Evening 10pm Post": false };
        week2Tasks[day] = { "Morning Chart": false, "Morning 2pm Post": false, "Evening Chart": false, "Evening 10pm Post": false };
      });

      for (let task of sections["Leaders Board"]) {
        let [week, rest] = task.split("): ");
        let [role, day] = rest.split(" on ");
        day = day.trim().charAt(0).toUpperCase() + day.slice(1).toLowerCase();
        let targetWeek = week.includes(dayToDateWeek1["Monday"].split(",")[1].trim().split(" ")[0]) ? week1Tasks : week2Tasks;
        if (targetWeek[day]) {
          targetWeek[day][role] = true;
          logAction("Assign Email Task", `Assigned task in ${week} for ${person}: ${role} on ${day}`);
        } else {
          logAction("Warning", `Day "${day}" not found in target week for task: ${task}`);
        }
      }

      leadersBoardHtml += `
        <h2 style="color: #4A2C2A; background-color: #F5F5F5; padding: 10px; border-radius: 5px;">Leaders Board (${Object.values(dayToDateWeek1)[0].split(",")[1].trim().split(" ")[0]} ${Object.values(dayToDateWeek1)[0].split(",")[1].trim().split(" ")[1]} - ${Object.values(dayToDateWeek1)[6].split(",")[1].trim().split(" ")[1]})</h2>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
          <tr style="background-color: #4A2C2A; color: white;">
            <th style="padding: 10px; border: 1px solid #ddd;">Day</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Morning Chart</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Morning 2pm Post</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Evening Chart</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Evening 10pm Post</th>
          </tr>
      `;
      for (let day of daysOfWeek) {
        leadersBoardHtml += `
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd;">${dayToDateWeek1[day]}</td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: ${week1Tasks[day]["Morning Chart"] ? "#E6F4EA" : ""};">${week1Tasks[day]["Morning Chart"] ? "✔" : ""}</td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: ${week1Tasks[day]["Morning 2pm Post"] ? "#E6F4EA" : ""};">${week1Tasks[day]["Morning 2pm Post"] ? "✔" : ""}</td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: ${week1Tasks[day]["Evening Chart"] ? "#E6F0FA" : ""};">${week1Tasks[day]["Evening Chart"] ? "✔" : ""}</td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: ${week1Tasks[day]["Evening 10pm Post"] ? "#E6F0FA" : ""};">${week1Tasks[day]["Evening 10pm Post"] ? "✔" : ""}</td>
          </tr>
        `;
      }
      leadersBoardHtml += `</table>`;

      leadersBoardHtml += `
        <h2 style="color: #4A2C2A; background-color: #F5F5F5; padding: 10px; border-radius: 5px;">Leaders Board (${Object.values(dayToDateWeek2)[0].split(",")[1].trim().split(" ")[0]} ${Object.values(dayToDateWeek2)[0].split(",")[1].trim().split(" ")[1]} - ${Object.values(dayToDateWeek2)[6].split(",")[1].trim().split(" ")[1]})</h2>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
          <tr style="background-color: #4A2C2A; color: white;">
            <th style="padding: 10px; border: 1px solid #ddd;">Day</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Morning Chart</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Morning 2pm Post</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Evening Chart</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Evening 10pm Post</th>
          </tr>
      `;
      for (let day of daysOfWeek) {
        leadersBoardHtml += `
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd;">${dayToDateWeek2[day]}</td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: ${week2Tasks[day]["Morning Chart"] ? "#E6F4EA" : ""};">${week2Tasks[day]["Morning Chart"] ? "✔" : ""}</td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: ${week2Tasks[day]["Morning 2pm Post"] ? "#E6F4EA" : ""};">${week2Tasks[day]["Morning 2pm Post"] ? "✔" : ""}</td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: ${week2Tasks[day]["Evening Chart"] ? "#E6F0FA" : ""};">${week2Tasks[day]["Evening Chart"] ? "✔" : ""}</td>
            <td style="padding: 10px; border: 1px solid #ddd; background-color: ${week2Tasks[day]["Evening 10pm Post"] ? "#E6F0FA" : ""};">${week2Tasks[day]["Evening 10pm Post"] ? "✔" : ""}</td>
          </tr>
        `;
      }
      leadersBoardHtml += `</table>`;
      delete sections["Leaders Board"];
    }

    let otherSectionsHtml = "";
    for (let section in sections) {
      otherSectionsHtml += `
        <h2 style="color: #4A2C2A; background-color: #F5F5F5; padding: 10px; border-radius: 5px;">${section}</h2>
        <ul style="margin-bottom: 20px; padding-left: 20px;">
      `;
      for (let task of sections[section]) {
        let [dateOrType, details] = task.includes(":") ? task.split(": ") : [task, ""];
        otherSectionsHtml += `
          <li style="margin-bottom: 5px;">
            <strong>${dateOrType}</strong>${details ? `: ${details}` : ""}
          </li>
        `;
      }
      otherSectionsHtml += `</ul>`;
    }

    const announcementText = getAnnouncements();
    const announcementsHtml = `
      <h2 style="color: #4A2C2A; background-color: #F5F5F5; padding: 10px; border-radius: 5px;">Announcements</h2>
      <p style="color: #333; margin-bottom: 20px;">${announcementText}</p>
    `;

    const contactIconsHtml = `
      <div style="text-align: center; margin-top: 20px;">
        <a href="mailto:pavlo.tsyhanash@gmail.com" style="margin: 0 10px;">
          <img src="https://ssl.gstatic.com/ui/v1/icons/mail/images/favicon5.ico" alt="Email" style="width: 16px; height: 16px; opacity: 0.6;" onmouseover="this.style.opacity='1'; this.style.filter='drop-shadow(0 0 2px #4A2C2A);'" onmouseout="this.style.opacity='0.6'; this.style.filter='none';">
        </a>
        <a href="https://t.me/R_O_R_K" style="margin: 0 10px;">
          <img src="https://telegram.org/favicon.ico" alt="Telegram" style="width: 16px; height: 16px; opacity: 0.6;" onmouseover="this.style.opacity='1'; this.style.filter='drop-shadow(0 0 2px #4A2C2A);'" onmouseout="this.style.opacity='0.6'; this.style.filter='none';">
        </a>
      </div>
    `;

    let htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
        <div style="text-align: center; margin-bottom: 20px;">
          <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/8/85/Tim_Hortons_logo_%28original%29.svg/500px-Tim_Hortons_logo_%28original%29.svg.png" alt="Tim Hortons Logo" style="max-width: 150px;" />
        </div>
        <h1 style="color: #4A2C2A; text-align: center;">Hello ${person}!</h1>
        <p style="color: #333; text-align: center; margin-bottom: 30px;">
          ${isTestMode ? `This email would have been sent to ${originalEmail}.<br>` : ""}
          Here are your tasks for the period ${periodLabel}:
        </p>
        ${leadersBoardHtml}
        ${otherSectionsHtml}
        ${announcementsHtml}
        <div style="text-align: center; margin-top: 30px;">
          <p style="color: #4A2C2A;">Best regards,<br>Your Tim Hortons Team</p>
          ${contactIconsHtml}
        </div>
      </div>
    `;

    const emailToSend = isTestMode ? testEmail : originalEmail;
    try {
      GmailApp.sendEmail(emailToSend, subject, "", { htmlBody: htmlBody });
      logAction("Send Email", `Sent email to ${emailToSend} with subject: ${subject}`);
      Utilities.sleep(1000);
    } catch (e) {
      logAction("Email Send Error", `Failed to send email to ${emailToSend}: ${e.message}`);
      GmailApp.sendEmail(
        "pavlo.tsyhanash@gmail.com",
        "Script Error: Email Send Failure",
        `Failed to send email to ${emailToSend}: ${e.message}`
      );
    }
  }
}

function onOpen() {
  setupSettingsSheet();
  setupAnnouncementsSheet();
  setupLogsSheet();
  setupMissingSheets();
  SpreadsheetApp.getUi()
    .createMenu('Tasks')
    .addItem('Preview & Send Tasks', 'previewTasks')
    .addItem('Send Test Emails', 'sendTestEmails')
    .addItem('Auto Send Tasks (No Preview)', 'autoSendTasks')
    .addSeparator()
    .addItem('Show Instructions', 'showInstructions')
    .addItem('Show Announcements', 'showAnnouncements')
    .addItem('View Logs', 'showLogs')
    .addToUi();
}
