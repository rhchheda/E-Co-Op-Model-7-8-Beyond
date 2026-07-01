/**
 * E-Co-Op Program Setup Script
 * KLE Technological University | C-SHINE
 *
 * One-time setup: builds a professionally formatted, validated, and
 * partially locked Google Sheet workbook seeded with the real program data
 * (teams, tasks, faculty mentors, industry mentors), including:
 *  - Teams: 6 structured student slots (Name/SRN/Phone), a live Team Size
 *    formula, and a live Duplicate SRN? formula that flags a student SRN
 *    repeated within the same team OR appearing on more than one team.
 *  - Tasks: due dates stored as real dates (DD-MM-YYYY format) where the
 *    original value was unambiguous, plus live Days Until Due / Overdue?
 *    formulas.
 *  - Overview: intelligence rollups (overdue tasks, teams missing a
 *    mentor, duplicate SRNs found, average team size).
 *
 * HOW TO USE:
 * 1. Create a new blank Google Sheet (sheets.new)
 * 2. Extensions > Apps Script
 * 3. Delete any starter code, paste this entire file in
 * 4. Click Run > setupWorkbook (first run will ask you to authorize)
 * 5. Switch back to the Sheet — it will be fully built
 * 6. File > Share > "Anyone with the link" > Viewer (needed so the portal
 *    website can read it). Editors should still be added by email for editing.
 *
 * Re-running setupWorkbook() is safe — it clears and rebuilds each tab from
 * scratch, so don't re-run it after you've started editing real data by hand,
 * or your edits will be overwritten.
 */

const BRAND_COLOR = "#0d1b3e";
const BRAND_TEXT = "#FFFFFF";
const MAX_ROWS = 200;
const MAX_STUDENTS = 6;

// Teams column layout (1-indexed). Student fields are grouped by type
// (all Names, then all SRNs, then all Phones) rather than per-student, so
// duplicate-SRN and team-size formulas can reference simple contiguous
// ranges instead of stitching together interleaved columns.
const TEAMS_COL = {
  ID: 1, DEPARTMENT: 2, PROJECT: 3,
  NAME_START: 4,                      // 4..9   Student 1-6 Name
  SRN_START: 4 + MAX_STUDENTS,        // 10..15 Student 1-6 SRN
  PHONE_START: 4 + MAX_STUDENTS * 2,  // 16..21 Student 1-6 Phone
  FACULTY_MENTOR: 4 + MAX_STUDENTS * 3,     // 22
  FACULTY_PHONE: 5 + MAX_STUDENTS * 3,      // 23
  END_USERS: 6 + MAX_STUDENTS * 3,          // 24
  DESCRIPTION: 7 + MAX_STUDENTS * 3,        // 25
  EXCITEMENT: 8 + MAX_STUDENTS * 3,         // 26
  STATUS: 9 + MAX_STUDENTS * 3,             // 27
  TEAM_SIZE: 10 + MAX_STUDENTS * 3,         // 28 (formula)
  DUP_SRN: 11 + MAX_STUDENTS * 3,           // 29 (formula)
};
const TEAMS_NUM_COLS = TEAMS_COL.DUP_SRN;

const TASKS_COL = {
  SL: 1, ACTIVITY: 2, DETAILS: 3, DUE: 4, OWNER: 5, STATUS: 6, OUTCOMES: 7, REMARKS: 8,
  DAYS_UNTIL_DUE: 9,  // formula
  OVERDUE: 10,        // formula
};
const TASKS_NUM_COLS = TASKS_COL.OVERDUE;

function colToLetter(col) {
  let letter = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  buildTeamsSheet(ss);
  buildTasksSheet(ss);
  buildFacultyMentorsSheet(ss);
  buildIndustryMentorsSheet(ss);
  buildOverviewSheet(ss);

  const def = ss.getSheetByName("Sheet1");
  if (def && def.getLastRow() === 0) {
    ss.deleteSheet(def);
  }

  ss.setActiveSheet(ss.getSheetByName("Overview"));
  SpreadsheetApp.getUi().alert("Setup complete. Review the Overview tab, then share the sheet as 'Anyone with the link: Viewer' so the portal website can read it.");
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
    sheet.clearFormats();
    sheet.setConditionalFormatRules([]);
  } else {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function styleHeader(sheet, numCols) {
  const header = sheet.getRange(1, 1, 1, numCols);
  header.setBackground(BRAND_COLOR);
  header.setFontColor(BRAND_TEXT);
  header.setFontWeight("bold");
  header.setVerticalAlignment("middle");
  header.setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 32);
}

function bandRows(sheet, startRow, numRows, numCols) {
  if (numRows <= 0) return;
  sheet.getRange(startRow, 1, numRows, numCols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

function protectHeader(sheet, numCols) {
  const range = sheet.getRange(1, 1, 1, numCols);
  const protection = range.protect().setDescription("Header - locked");
  protection.setWarningOnly(false);
  const me = Session.getEffectiveUser();
  protection.removeEditors(protection.getEditors());
  if (me && me.getEmail()) protection.addEditor(me);
}

function protectFormulaColumns(sheet, startCol, numCols, numRows) {
  const range = sheet.getRange(2, startCol, numRows, numCols);
  const protection = range.protect().setDescription("Formula columns - locked");
  protection.setWarningOnly(true); // warn, don't hard-block, since Apps Script itself must still write here on rebuild
}

function addDropdown(sheet, row, col, numRows, choices) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(choices, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, col, numRows, 1).setDataValidation(rule);
}

// ---------------- TEAMS ----------------
function buildTeamsSheet(ss) {
  const headers = ["ID", "Department", "Project Name"];
  for (let i = 1; i <= MAX_STUDENTS; i++) headers.push("Student " + i + " Name");
  for (let i = 1; i <= MAX_STUDENTS; i++) headers.push("Student " + i + " SRN");
  for (let i = 1; i <= MAX_STUDENTS; i++) headers.push("Student " + i + " Phone");
  headers.push("Faculty Mentor", "Faculty Phone", "End Users", "Description", "Excitement / Interest", "Team Status", "Team Size", "Duplicate SRN?");

  const data = [
  [1, "Computer Science and Engineering", "Sintex Cleanliness Detection Using IoT and AI Model", "Subramanya Tiluvalli", "Praveen Angadi", "Basavant Mahalingpur", "Rohit Reddy", "", "", "01FE23BCS279", "01FE23BCS298", "01FE23BCS272", "01FE23BCS293", "", "", "9110607686", "9972436163", "9353944089", "9731590011", "", "", "Geeta Sannakki", "9620161698", "Apartment Complexes", "A box with a camera that is fitted in the overhead tank and connects with internet to send notification either through whatsapp or telegram to apartment or big building complexes.", "High and are interested to continue till productization", ""],
  [2, "Computer Science and Engineering", "AI-Enabled eCommerce for Retailers", "Vishal Naik", "Madhura", "Pratham Kathare", "Aditya Naik", "", "", "01FE23BCS294", "01FE23BCS103", "01FE23BCS085", "01FE23BCS211", "", "", "6366349040", "7019325490", "8277777747", "8762539424", "", "", "Lalita Madanabhavi", "8050458109", "Local Vendors and Customers", "An e-commerce software platform for enabling business for local vendors and customers", "High and are interested to continue till productization", ""],
  [3, "Computer Science and Engineering", "Integrated academic performance monitoring system", "Anagha Nadgouda", "Sara Patil", "Zaid Momin", "Sushant Maheshwari", "", "", "01FE23BCS076", "01FE23BCS050", "01FE23BCS121", "01FE23BCS211", "", "", "99804 54365", "9663081911", "7975332473", "8050003427", "", "", "Lalita Madanabhavi", "8050458109", "Students, Universities and colleges", "A recommendation system based on students' portfolios and trend analysis.", "High and are interested to continue till productization. There is a disconnect in the project title and what they are doing.", ""],
  [4, "Computer Science and Engineering", "Decentralized Academic Credential Verification System using Self-Sovereign Identity", "Swateja Patil", "Priyanka N D", "", "", "", "", "01FE23BCS161", "01FE23BCS184", "", "", "", "", "9900657272", "8073333610", "", "", "", "", "Pooja Shettar", "9964141448", "Academic Institutions, employee verification required institutions", "A credential verification system built using blockchain", "High and are interested to continue till productization", ""],
  [5, "Mechanical Engineering", "Camera based navigation system for visually impaired", "Basuraj Shibargatti", "Akash Halesh Walad", "Vrushab Kadam", "Sachin Channaveer", "Manoj Rode", "Surajkumar B", "01FE23BME120", "01FE23BME133", "01FE23BME139", "01FE23BME149", "01FE23BME152", "01FE23BME106", "8660128071", "7411445754", "9535188782", "7019541358", "7406342310", "9036008183", "Gururaj Fattepur", "9739461325", "Blind people", "A handheld device that can assist blind people during walking", "High", ""],
  [6, "Mechanical Engineering", "Adaptive sujok reflexology-based acupressure therapy device", "Swaroop Biradar", "Vivek Padi", "Vishal Chavannavar", "Shivaraj Hiremath", "Siddamma Meti", "Gayatri Angadi", "01FE23BME054", "01FE23BME058", "01FE23BME080", "01FE23BME081", "01FE23BME141", "01FE23BME035", "6361159255", "8431978516", "9449570463", "9960981497", "7975690922", "9449101661", "Gururaj Fattepur", "9739461325", "Therapists and customers who prefer reflexology", "A machine that can detect the palm and appropriately apply pressure based on SUJOK reflexology techniques", "HIgh", ""],
  [7, "Mechanical Engineering", "Smart shopping cart", "Teerth S. Kulkarni", "Anjali V. Purohit", "Sejal V. Komalapur", "Karthik G. Nadurmath", "Manjunath P. Badiger", "", "01FE24BME405", "01FE23BME068", "01FE24BME430", "01FE24BME424", "01FE24BME435", "", "9019846034", "8554005888", "7348921572", "7019709717", "8971709642", "", "Veerabhadrayya Hiremath, Nagaraj Ekbote", "9900506604, 9591017854", "Shopping centers or shopping malls where shopping carts are extensively used", "A shopping cart with a display, item identification, and a payment portal", "High", ""],
  [8, "Mechanical Engineering", "Smart Servo-Based Electromechanical Disc Lock for two wheelers", "VRUSHABRAJ KALYANKAR", "RAHUL S HALADANDIMATH", "TEJAS HINDASAGERI", "ADIT V PATIL", "ROHIT MADAR", "VRASHABHA S VASTRAD", "01FE21BME014", "01FE22BME050", "01FE22BME416", "01FE22BME427", "01FE23BME076", "01FE23BME108", "7899424187", "6362568936", "9380244302", "7406216111", "9686523466", "9606742193", "Praveen Petkar, Mantesh Choukimath", "9964476030, 7760072120", "All Bike users", "A disc brake lock unit", "High", ""],
  [9, "Automation and Robotics", "Automated PCB Defect Detection System Using Computer Vision", "Shreya Chaligeri", "Rama Kulkarni", "Abhishek Hiremath", "Manjunath Sali", "", "", "01FE23BAR017", "01FE23BAR052", "01FE24BAR402", "01FE24BAR407", "", "", "7204724565", "", "", "", "", "", "Vinod Meti", "9986356557", "SME Industries, AEQUS", "A machine which can pass the PCBs automatically and detect the defects and finally sort the PCBs as good and bad ones.", "Excited and willing to continue till productization", ""],
  [10, "Automation and Robotics", "Vision-Based Detection and Counting of Cookware in the Consumer Goods Industry", "Harsha Kampli", "Sofiya Kittur", "Usaid Mansabdar", "Virupaksha Amarshetti", "Gaurav Sooji", "", "01FE23BAR054", "01FE23BAR013", "01FE23BAR041", "01FE23BAR039", "01FE20BAR002", "", "9152493660", "", "", "", "", "", "Vijay Mahantesh", "9980205502", "Aequs, SME Industries - Large Industries", "A software that can visually count and classify different types of vessels or ccokware", "Excited and willing to continue till productization", ""],
  [11, "Electrical and Electronics Engineering", "Dual-Output Quadratic Boost Converter with Enhanced Voltage Conversion Ratio.", "Dipti.S", "Sanjana.I.K", "Rakshita.S.B", "Amitvikram yeri", "", "", "01fe23bee005", "01fe23bee010", "01fe23bee023", "01fe23bee024", "", "", "6360873027", "9945840123", "7619125343", "9110260512", "", "", "Vinod Patil", "9880998964", "EV Charging stations, high voltage battery charging", "They are still thinking on this aspect", "High", ""],
  [12, "Electrical and Electronics Engineering", "Position and Trajectory Control of a 2-DOF Robotic Arm.", "Shivkumar G", "Abhishek wadeyar", "Akash Kempanavvar", "", "", "", "01fe22bee082", "01fe23bee103", "01fe23bee086", "", "", "", "6361406074", "7619457369", "7619303794", "", "", "", "", "", "SME industries such as PCB handling, which require smooth and soft handling of the items", "No clarity on this part", "They are excited. But they are not sure what their product idea would be by the end of 7th sem.", ""]
];

  const sheet = getOrCreateSheet(ss, "Teams");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (data.length) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  styleHeader(sheet, headers.length);
  bandRows(sheet, 2, MAX_ROWS, headers.length);
  protectHeader(sheet, headers.length);

  sheet.setColumnWidth(TEAMS_COL.PROJECT, 240);
  sheet.getRange(2, TEAMS_COL.DESCRIPTION, MAX_ROWS, 1).setWrap(true);
  sheet.getRange(2, TEAMS_COL.EXCITEMENT, MAX_ROWS, 1).setWrap(true);

  addDropdown(sheet, 2, TEAMS_COL.STATUS, MAX_ROWS, ["Not Started", "On Track", "At Risk", "Behind", "Completed"]);

  const statusRange = sheet.getRange(2, TEAMS_COL.STATUS, MAX_ROWS, 1);
  const rules = [
    {v: "On Track", c: "#D9EAD3"}, {v: "Completed", c: "#B6D7A8"},
    {v: "At Risk", c: "#FFF2CC"}, {v: "Behind", c: "#F4CCCC"}, {v: "Not Started", c: "#EFEFEF"},
  ].map(r => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r.v).setBackground(r.c).setRanges([statusRange]).build());

  // ---- Team Size formula: count of filled Name slots ----
  const nameColLetterStart = colToLetter(TEAMS_COL.NAME_START);
  const nameColLetterEnd = colToLetter(TEAMS_COL.NAME_START + MAX_STUDENTS - 1);
  const teamSizeFormulas = [];
  for (let r = 2; r <= MAX_ROWS + 1; r++) {
    teamSizeFormulas.push([`=COUNTA(${nameColLetterStart}${r}:${nameColLetterEnd}${r})`]);
  }
  sheet.getRange(2, TEAMS_COL.TEAM_SIZE, MAX_ROWS, 1).setFormulas(teamSizeFormulas);

  // ---- Duplicate SRN? formula: flags a repeated SRN within the row OR
  // anywhere else in the whole SRN range (cross-team duplicate) ----
  const srnColLetterStart = colToLetter(TEAMS_COL.SRN_START);
  const srnColLetterEnd = colToLetter(TEAMS_COL.SRN_START + MAX_STUDENTS - 1);
  const srnFullRange = `$${srnColLetterStart}$2:$${srnColLetterEnd}$${MAX_ROWS + 1}`;
  const dupFormulas = [];
  for (let r = 2; r <= MAX_ROWS + 1; r++) {
    const rowRange = `${srnColLetterStart}${r}:${srnColLetterEnd}${r}`;
    dupFormulas.push([
      `=IF(COUNTA(${rowRange})=0,"",IF(SUMPRODUCT((COUNTIF(${srnFullRange},${rowRange})>1)*(${rowRange}<>""))>0,"⚠ Duplicate SRN",""))`
    ]);
  }
  sheet.getRange(2, TEAMS_COL.DUP_SRN, MAX_ROWS, 1).setFormulas(dupFormulas);

  const dupRange = sheet.getRange(2, TEAMS_COL.DUP_SRN, MAX_ROWS, 1);
  sheet.setConditionalFormatRules(rules.concat([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("⚠ Duplicate SRN").setBackground("#F4CCCC").setFontColor("#C0392B").setRanges([dupRange]).build(),
  ]));

  protectFormulaColumns(sheet, TEAMS_COL.TEAM_SIZE, 2, MAX_ROWS);
}

// ---------------- TASKS ----------------
function buildTasksSheet(ss) {
  const headers = ["Sl. No", "Task / Activity", "Details / Link", "Due Date", "Owner", "Status", "Outcomes", "Remarks", "Days Until Due", "Overdue?"];
  const data = [
  [1, "Receiving shortlisted projects on minor projects after PIC scrutiny from all departments", "", "17-06-2026", "", "Complete", "List of project teams with details", "4 departments have responded with their lists (CSE, EEE, Mech and A&R)"],
  [2, "CSHINE review of shortlisted minor projects and finalizing for E-Co-Op after talking to student teams", "", "25-06-2026", "SK", "Complete", "Compiled list of project teams with details", "10 to 11 teams have been identified from all 4 departments that shared the initial project teams list."],
  [3, "Creating a WhatsApp group containing shortlisted student teams, faculty mentors, and CSHINE", "", "27-06-2026", "SK", "Complete", "A WhatsApp group", "SK has added one student per team to the WhatsApp group and requested that the student add other teammates to the group."],
  [4, "Creating a WhatsApp group containing faculty mentors and CSHINE", "", "27-06-2026", "SK", "Complete", "A WhatsApp group", "VijayaMahantesh from A&R is pending"],
  [5, "Pre Market Survey presentation and questionnaire template design", "", "01-07-2026", "SK and NK", "Pending", "A PPT and a Questionnaire template", "[VERIFY DATE: auto-corrected from an ambiguous spreadsheet date]"],
  [6, "Faculty mentors orientation on E-Co-Op and Customer insight survey", "https://docs.google.com/presentation/d/1Fxhi1FARlmZpbjGLy0wmuqngIC8VpPiCw2Kxa9I6UUc/edit?slide=id.g3f2671f6f93_0_0#slide=id.g3f2671f6f93_0_0", "3rd July @2.00 PM to 3.00 PM", "SK and NK", "Pending", "", ""],
  [7, "Faculty mentors orientation on TRL levels", "", "3rd July at 3.00 pm to 4.00 pm", "Nandish Harti and Nandish Humbi", "Pending", "", ""],
  [8, "First meeting with students.", "https://docs.google.com/presentation/d/1P-jpenXq0YETBT5rVBG9UMOpKcBNMpO53CtrfwGOrQs/edit?slide=id.p#slide=id.p", "7th to 9th July", "SK and NK", "Pending", "", ""],
  [9, "Customer insight Survey Reporting by students", "", "31-07-2026", "Students and Faculty", "Pending", "", ""],
  [10, "Compiling Customer Insight survey reports", "", "04-08-2026", "SK and Faculty", "Pending", "", "[VERIFY DATE: auto-corrected from an ambiguous spreadsheet date]"],
  [11, "Identifying industry experts", "https://docs.google.com/spreadsheets/d/1grgp4eipX0JBLVPNOSQ0X7VIAwFBGpdikpAXGv1CvfI/edit?usp=drive_link", "15-07-2026", "SK and NK", "Pending", "A list of proactive industry mentors", ""],
  [12, "Have industry experts check on team MVP statements based on their pre market survey", "", "15-08-2026", "Industry Mentors", "Pending", "", ""],
  [13, "List of Backup Projects", "", "30-07-2026", "SK and NK", "Pending", "", ""],
  [14, "Second meeting with Students during the first week of the semester start", "", "10th to 25th August", "SK and NK", "Pending", "", ""]
];
  const parseableDueRows = [1, 2, 3, 4, 5, 9, 10, 11, 12, 13];

  const sheet = getOrCreateSheet(ss, "Tasks");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (data.length) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  styleHeader(sheet, headers.length);
  bandRows(sheet, 2, MAX_ROWS, headers.length);
  protectHeader(sheet, headers.length);

  sheet.setColumnWidth(TASKS_COL.ACTIVITY, 320);
  sheet.setColumnWidth(TASKS_COL.REMARKS, 260);
  sheet.getRange(2, TASKS_COL.ACTIVITY, MAX_ROWS, 1).setWrap(true);
  sheet.getRange(2, TASKS_COL.REMARKS, MAX_ROWS, 1).setWrap(true);

  addDropdown(sheet, 2, TASKS_COL.STATUS, MAX_ROWS, ["Pending", "In Progress", "Complete", "Blocked"]);

  // Convert unambiguous "dd-mm-yyyy" due dates to real Date objects so
  // they sort correctly and the Overdue/Days-Until-Due formulas work.
  parseableDueRows.forEach((rowNum) => {
    const cell = sheet.getRange(rowNum + 1, TASKS_COL.DUE); // rowNum is 1-indexed data row -> sheet row = rowNum+1
    const raw = String(cell.getValue());
    const m = raw.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (m) {
      const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
      cell.setValue(d);
    }
  });
  sheet.getRange(2, TASKS_COL.DUE, MAX_ROWS, 1).setNumberFormat("dd-mm-yyyy");

  const dueColLetter = colToLetter(TASKS_COL.DUE);
  const statusColLetter = colToLetter(TASKS_COL.STATUS);
  const daysFormulas = [];
  const overdueFormulas = [];
  for (let r = 2; r <= MAX_ROWS + 1; r++) {
    daysFormulas.push([`=IF(ISNUMBER(${dueColLetter}${r}),${dueColLetter}${r}-TODAY(),"")`]);
    overdueFormulas.push([`=IF(AND(ISNUMBER(${dueColLetter}${r}),${dueColLetter}${r}<TODAY(),${statusColLetter}${r}<>"Complete"),"⚠ OVERDUE","")`]);
  }
  sheet.getRange(2, TASKS_COL.DAYS_UNTIL_DUE, MAX_ROWS, 1).setFormulas(daysFormulas);
  sheet.getRange(2, TASKS_COL.OVERDUE, MAX_ROWS, 1).setFormulas(overdueFormulas);

  const statusRange = sheet.getRange(2, TASKS_COL.STATUS, MAX_ROWS, 1);
  const overdueRange = sheet.getRange(2, TASKS_COL.OVERDUE, MAX_ROWS, 1);
  const rules = [
    {v: "Complete", c: "#B6D7A8"}, {v: "In Progress", c: "#FFF2CC"},
    {v: "Pending", c: "#EFEFEF"}, {v: "Blocked", c: "#F4CCCC"},
  ].map(r => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r.v).setBackground(r.c).setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("⚠ OVERDUE").setBackground("#F4CCCC").setFontColor("#C0392B").setRanges([overdueRange]).build());
  sheet.setConditionalFormatRules(rules);

  protectFormulaColumns(sheet, TASKS_COL.DAYS_UNTIL_DUE, 2, MAX_ROWS);
}

// ---------------- FACULTY MENTORS (derived, read-only) ----------------
function buildFacultyMentorsSheet(ss) {
  const headers = ["Name", "Phone", "Teams Mentored", "Number of Teams"];
  const sheet = getOrCreateSheet(ss, "Faculty Mentors");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length);
  sheet.setColumnWidth(3, 320);
  sheet.getRange(2, 3, MAX_ROWS, 1).setWrap(true);

  const protection = sheet.protect().setDescription("Derived from Teams tab - edit Teams instead");
  protection.setWarningOnly(true);

  rebuildFacultyMentorsFromTeams(ss);
}

// Recomputes the Faculty Mentors tab from the Teams tab. Safe to call any
// time Teams changes (also called by the admin app's Code.gs after CRUD).
function rebuildFacultyMentorsFromTeams(ss) {
  const teamsSheet = ss.getSheetByName("Teams");
  const facSheet = ss.getSheetByName("Faculty Mentors");
  if (!teamsSheet || !facSheet) return;

  const lastRow = teamsSheet.getLastRow();
  const map = {};
  if (lastRow >= 2) {
    const rows = teamsSheet.getRange(2, 1, lastRow - 1, TEAMS_NUM_COLS).getValues();
    rows.forEach((row) => {
      const project = row[TEAMS_COL.PROJECT - 1];
      const fm = String(row[TEAMS_COL.FACULTY_MENTOR - 1] || "");
      const fp = String(row[TEAMS_COL.FACULTY_PHONE - 1] || "");
      if (!fm) return;
      const names = fm.split(",").map(s => s.trim()).filter(Boolean);
      const phones = fp.split(",").map(s => s.trim());
      names.forEach((name, i) => {
        if (!map[name]) map[name] = { phone: phones[i] || "", teams: [] };
        if (project) map[name].teams.push(project);
      });
    });
  }

  const existingLastRow = facSheet.getLastRow();
  if (existingLastRow > 1) facSheet.getRange(2, 1, existingLastRow - 1, 4).clearContent();
  const outRows = Object.keys(map).map(name => [name, map[name].phone, map[name].teams.join(", "), map[name].teams.length]);
  if (outRows.length) facSheet.getRange(2, 1, outRows.length, 4).setValues(outRows);
}

// ---------------- INDUSTRY MENTORS ----------------
function buildIndustryMentorsSheet(ss) {
  const headers = ["Name", "Industry", "Background", "Phone", "Email", "Willingness to Participate"];
  const data = [
  ["Vaibhav", "EARTKey", "Electronics", "", "", "Not yet contacted"],
  ["Chetan Shettar", "", "Electronics", "", "", "Not yet contacted"],
  ["Anand Kadkol", "", "Business", "", "", "Not yet contacted"]
];

  const sheet = getOrCreateSheet(ss, "Industry Mentors");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (data.length) sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  styleHeader(sheet, headers.length);
  bandRows(sheet, 2, Math.max(data.length, 1), headers.length);
  protectHeader(sheet, headers.length);

  addDropdown(sheet, 2, 6, MAX_ROWS, ["Yes", "No", "Maybe", "Not yet contacted"]);
}

// ---------------- OVERVIEW ----------------
function buildOverviewSheet(ss) {
  const sheet = getOrCreateSheet(ss, "Overview");
  sheet.getRange(1, 1).setValue("E-Co-Op Program Tracker").setFontSize(18).setFontWeight("bold");
  sheet.getRange(2, 1).setValue("CSHINE, KLE Technological University").setFontSize(11).setFontColor("#666666");

  const teamsProjectCol = colToLetter(TEAMS_COL.PROJECT);
  const teamsMentorCol = colToLetter(TEAMS_COL.FACULTY_MENTOR);
  const teamsSizeCol = colToLetter(TEAMS_COL.TEAM_SIZE);
  const teamsDupCol = colToLetter(TEAMS_COL.DUP_SRN);
  const tasksStatusCol = colToLetter(TASKS_COL.STATUS);
  const tasksActivityCol = colToLetter(TASKS_COL.ACTIVITY);
  const tasksOverdueCol = colToLetter(TASKS_COL.OVERDUE);

  const rows = [
    ["Total Teams", `=COUNTA(Teams!${teamsProjectCol}2:${teamsProjectCol}${MAX_ROWS + 1})`],
    ["Faculty Mentors", `=COUNTA('Faculty Mentors'!A2:A)`],
    ["Industry Mentors", `=COUNTA('Industry Mentors'!A2:A)`],
    ["Tasks Complete", `=COUNTIF(Tasks!${tasksStatusCol}2:${tasksStatusCol}${MAX_ROWS + 1},"Complete")&" / "&COUNTA(Tasks!${tasksActivityCol}2:${tasksActivityCol}${MAX_ROWS + 1})`],
    ["Overdue Tasks", `=COUNTIF(Tasks!${tasksOverdueCol}2:${tasksOverdueCol}${MAX_ROWS + 1},"⚠ OVERDUE")`],
    ["Teams Missing Faculty Mentor", `=COUNTIFS(Teams!${teamsProjectCol}2:${teamsProjectCol}${MAX_ROWS + 1},"<>",Teams!${teamsMentorCol}2:${teamsMentorCol}${MAX_ROWS + 1},"")`],
    ["Duplicate SRNs Found", `=COUNTIF(Teams!${teamsDupCol}2:${teamsDupCol}${MAX_ROWS + 1},"⚠ Duplicate SRN")`],
    ["Average Team Size", `=IFERROR(ROUND(AVERAGEIF(Teams!${teamsProjectCol}2:${teamsProjectCol}${MAX_ROWS + 1},"<>",Teams!${teamsSizeCol}2:${teamsSizeCol}${MAX_ROWS + 1}),1),0)`],
  ];
  sheet.getRange(4, 1, rows.length, 2).setValues(rows);
  sheet.getRange(4, 1, rows.length, 1).setFontWeight("bold");
  sheet.autoResizeColumns(1, 2);
  sheet.getRange(4 + rows.length + 1, 1).setValue("This sheet is the live source of truth for the E-Co-Op portal website. Edit Teams / Tasks / Industry Mentors directly here - changes appear on the site automatically.").setFontStyle("italic").setFontColor("#666666").setWrap(true);
  sheet.setColumnWidth(1, 260);
}
