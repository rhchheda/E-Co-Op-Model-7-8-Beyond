/**
 * E-Co-Op Program Setup Script
 * KLE Technological University | C-SHINE
 *
 * One-time setup: builds a professionally formatted, validated, and
 * partially locked Google Sheet workbook seeded with the real program data.
 *
 * Tabs built: Overview, Teams, Tasks, Evaluation, Surveys, Faculty Mentors,
 * Industry Mentors.
 *
 * IMPORTANT ROW-FORMULA DESIGN NOTE (read before changing anything below):
 * Formula columns (Team Size, Duplicate SRN?, Days Until Due, Overdue?,
 * Stage 1 Total, etc.) are applied ONE ROW AT A TIME via the applyXRowFormulas_
 * functions below - both here at setup time (looped over the real seeded
 * rows only) AND by Code.gs every time a row is added through the admin app.
 * They are deliberately NOT bulk-pre-filled across all MAX_ROWS rows the way
 * dropdown validation and conditional formatting are. Pre-filling a formula
 * into an otherwise-empty row would give that row real cell content, which
 * makes Sheets' getLastRow() count it as "used" - breaking new-row detection
 * for every future Add. Dropdowns/formatting are safe to pre-fill because
 * they aren't cell content and don't affect getLastRow().
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
// (all Names, then all SRNs, then all Phones) so duplicate-SRN and
// team-size formulas can reference simple contiguous ranges.
const TEAMS_COL = {
  ID: 1, DEPARTMENT: 2, PROJECT: 3,
  NAME_START: 4,                             // 4..9   Student 1-6 Name
  SRN_START: 4 + MAX_STUDENTS,                // 10..15 Student 1-6 SRN
  PHONE_START: 4 + MAX_STUDENTS * 2,          // 16..21 Student 1-6 Phone
  FACULTY_MENTOR: 4 + MAX_STUDENTS * 3,       // 22
  FACULTY_PHONE: 5 + MAX_STUDENTS * 3,        // 23
  FACULTY_EMAIL: 6 + MAX_STUDENTS * 3,        // 24
  END_USERS: 7 + MAX_STUDENTS * 3,            // 25
  DESCRIPTION: 8 + MAX_STUDENTS * 3,          // 26
  EXCITEMENT: 9 + MAX_STUDENTS * 3,           // 27
  STATUS: 10 + MAX_STUDENTS * 3,              // 28
  MENTOR_NOTES: 11 + MAX_STUDENTS * 3,        // 29
  TEAM_SIZE: 12 + MAX_STUDENTS * 3,           // 30 (formula)
  DUP_SRN: 13 + MAX_STUDENTS * 3,             // 31 (formula)
};
const TEAMS_NUM_COLS = TEAMS_COL.DUP_SRN;

const TASKS_COL = {
  SL: 1, ACTIVITY: 2, DETAILS: 3, DUE: 4, OWNER: 5, STATUS: 6, OUTCOMES: 7, REMARKS: 8,
  DAYS_UNTIL_DUE: 9,  // formula
  OVERDUE: 10,        // formula
};

const EVAL_COL = {
  TEAM_ID: 1, PROJECT: 2 /* formula */,
  PROBLEM: 3, NOVELTY: 4, PROTOTYPE: 5, TEAM_CAP: 6, COMMERCIAL: 7,
  STAGE1_TOTAL: 8 /* formula */, STAGE1_RESULT: 9 /* formula */,
  MARKET: 10, SCALABILITY: 11, COACHABILITY: 12,
  PANEL_NOTES: 13, PANEL_VERDICT: 14,
};

const SURVEY_COL = {
  TEAM_ID: 1, PROJECT: 2 /* formula */,
  INTERVIEWS_DONE: 3, INTERVIEWS_TARGET: 4,
  PROGRESS: 5 /* formula */,
  REPORT_LINK: 6, REVIEW_STATUS: 7, FACULTY_COMMENTS: 8, INSIGHTS_SUMMARY: 9,
};

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
  buildEvaluationSheet(ss);
  buildSurveysSheet(ss);
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

function addDropdown(sheet, row, col, numRows, choices) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(choices, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, col, numRows, 1).setDataValidation(rule);
}

function addNumberRangeValidation(sheet, row, col, numRows, min, max) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(min, max)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, col, numRows, 1).setDataValidation(rule);
}

// ==================== TEAMS ====================

function buildTeamsSheet(ss) {
  const headers = ["ID", "Department", "Project Name"];
  for (let i = 1; i <= MAX_STUDENTS; i++) headers.push("Student " + i + " Name");
  for (let i = 1; i <= MAX_STUDENTS; i++) headers.push("Student " + i + " SRN");
  for (let i = 1; i <= MAX_STUDENTS; i++) headers.push("Student " + i + " Phone");
  headers.push(
    "Faculty Mentor", "Faculty Phone", "Faculty Mentor Email",
    "End Users", "Description", "Excitement / Interest", "Team Status", "Mentor Notes",
    "Team Size", "Duplicate SRN?"
  );

  const data = [
  [1, "Computer Science and Engineering", "Sintex Cleanliness Detection Using IoT and AI Model", "Subramanya Tiluvalli", "Praveen Angadi", "Basavant Mahalingpur", "Rohit Reddy", "", "", "01FE23BCS279", "01FE23BCS298", "01FE23BCS272", "01FE23BCS293", "", "", "9110607686", "9972436163", "9353944089", "9731590011", "", "", "Geeta Sannakki", "9620161698", "", "Apartment Complexes", "A box with a camera that is fitted in the overhead tank and connects with internet to send notification either through whatsapp or telegram to apartment or big building complexes.", "High and are interested to continue till productization", "", ""],
  [2, "Computer Science and Engineering", "AI-Enabled eCommerce for Retailers", "Vishal Naik", "Madhura", "Pratham Kathare", "Aditya Naik", "", "", "01FE23BCS294", "01FE23BCS103", "01FE23BCS085", "01FE23BCS211", "", "", "6366349040", "7019325490", "8277777747", "8762539424", "", "", "Lalita Madanabhavi", "8050458109", "", "Local Vendors and Customers", "An e-commerce software platform for enabling business for local vendors and customers", "High and are interested to continue till productization", "", ""],
  [3, "Computer Science and Engineering", "Integrated academic performance monitoring system", "Anagha Nadgouda", "Sara Patil", "Zaid Momin", "Sushant Maheshwari", "", "", "01FE23BCS076", "01FE23BCS050", "01FE23BCS121", "01FE23BCS211", "", "", "99804 54365", "9663081911", "7975332473", "8050003427", "", "", "Lalita Madanabhavi", "8050458109", "", "Students, Universities and colleges", "A recommendation system based on students' portfolios and trend analysis.", "High and are interested to continue till productization. There is a disconnect in the project title and what they are doing.", "", ""],
  [4, "Computer Science and Engineering", "Decentralized Academic Credential Verification System using Self-Sovereign Identity", "Swateja Patil", "Priyanka N D", "", "", "", "", "01FE23BCS161", "01FE23BCS184", "", "", "", "", "9900657272", "8073333610", "", "", "", "", "Pooja Shettar", "9964141448", "", "Academic Institutions, employee verification required institutions", "A credential verification system built using blockchain", "High and are interested to continue till productization", "", ""],
  [5, "Mechanical Engineering", "Camera based navigation system for visually impaired", "Basuraj Shibargatti", "Akash Halesh Walad", "Vrushab Kadam", "Sachin Channaveer", "Manoj Rode", "Surajkumar B", "01FE23BME120", "01FE23BME133", "01FE23BME139", "01FE23BME149", "01FE23BME152", "01FE23BME106", "8660128071", "7411445754", "9535188782", "7019541358", "7406342310", "9036008183", "Gururaj Fattepur", "9739461325", "", "Blind people", "A handheld device that can assist blind people during walking", "High", "", ""],
  [6, "Mechanical Engineering", "Adaptive sujok reflexology-based acupressure therapy device", "Swaroop Biradar", "Vivek Padi", "Vishal Chavannavar", "Shivaraj Hiremath", "Siddamma Meti", "Gayatri Angadi", "01FE23BME054", "01FE23BME058", "01FE23BME080", "01FE23BME081", "01FE23BME141", "01FE23BME035", "6361159255", "8431978516", "9449570463", "9960981497", "7975690922", "9449101661", "Gururaj Fattepur", "9739461325", "", "Therapists and customers who prefer reflexology", "A machine that can detect the palm and appropriately apply pressure based on SUJOK reflexology techniques", "HIgh", "", ""],
  [7, "Mechanical Engineering", "Smart shopping cart", "Teerth S. Kulkarni", "Anjali V. Purohit", "Sejal V. Komalapur", "Karthik G. Nadurmath", "Manjunath P. Badiger", "", "01FE24BME405", "01FE23BME068", "01FE24BME430", "01FE24BME424", "01FE24BME435", "", "9019846034", "8554005888", "7348921572", "7019709717", "8971709642", "", "Veerabhadrayya Hiremath, Nagaraj Ekbote", "9900506604, 9591017854", "", "Shopping centers or shopping malls where shopping carts are extensively used", "A shopping cart with a display, item identification, and a payment portal", "High", "", ""],
  [8, "Mechanical Engineering", "Smart Servo-Based Electromechanical Disc Lock for two wheelers", "VRUSHABRAJ KALYANKAR", "RAHUL S HALADANDIMATH", "TEJAS HINDASAGERI", "ADIT V PATIL", "ROHIT MADAR", "VRASHABHA S VASTRAD", "01FE21BME014", "01FE22BME050", "01FE22BME416", "01FE22BME427", "01FE23BME076", "01FE23BME108", "7899424187", "6362568936", "9380244302", "7406216111", "9686523466", "9606742193", "Praveen Petkar, Mantesh Choukimath", "9964476030, 7760072120", "", "All Bike users", "A disc brake lock unit", "High", "", ""],
  [9, "Automation and Robotics", "Automated PCB Defect Detection System Using Computer Vision", "Shreya Chaligeri", "Rama Kulkarni", "Abhishek Hiremath", "Manjunath Sali", "", "", "01FE23BAR017", "01FE23BAR052", "01FE24BAR402", "01FE24BAR407", "", "", "7204724565", "", "", "", "", "", "Vinod Meti", "9986356557", "", "SME Industries, AEQUS", "A machine which can pass the PCBs automatically and detect the defects and finally sort the PCBs as good and bad ones.", "Excited and willing to continue till productization", "", ""],
  [10, "Automation and Robotics", "Vision-Based Detection and Counting of Cookware in the Consumer Goods Industry", "Harsha Kampli", "Sofiya Kittur", "Usaid Mansabdar", "Virupaksha Amarshetti", "Gaurav Sooji", "", "01FE23BAR054", "01FE23BAR013", "01FE23BAR041", "01FE23BAR039", "01FE20BAR002", "", "9152493660", "", "", "", "", "", "Vijay Mahantesh", "9980205502", "", "Aequs, SME Industries - Large Industries", "A software that can visually count and classify different types of vessels or ccokware", "Excited and willing to continue till productization", "", ""],
  [11, "Electrical and Electronics Engineering", "Dual-Output Quadratic Boost Converter with Enhanced Voltage Conversion Ratio.", "Dipti.S", "Sanjana.I.K", "Rakshita.S.B", "Amitvikram yeri", "", "", "01fe23bee005", "01fe23bee010", "01fe23bee023", "01fe23bee024", "", "", "6360873027", "9945840123", "7619125343", "9110260512", "", "", "Vinod Patil", "9880998964", "", "EV Charging stations, high voltage battery charging", "They are still thinking on this aspect", "High", "", ""],
  [12, "Electrical and Electronics Engineering", "Position and Trajectory Control of a 2-DOF Robotic Arm.", "Shivkumar G", "Abhishek wadeyar", "Akash Kempanavvar", "", "", "", "01fe22bee082", "01fe23bee103", "01fe23bee086", "", "", "", "6361406074", "7619457369", "7619303794", "", "", "", "", "", "", "SME industries such as PCB handling, which require smooth and soft handling of the items", "No clarity on this part", "They are excited. But they are not sure what their product idea would be by the end of 7th sem.", "", ""]
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
  sheet.getRange(2, TEAMS_COL.MENTOR_NOTES, MAX_ROWS, 1).setWrap(true);

  addDropdown(sheet, 2, TEAMS_COL.STATUS, MAX_ROWS, ["Not Started", "On Track", "At Risk", "Behind", "Completed"]);

  const statusRange = sheet.getRange(2, TEAMS_COL.STATUS, MAX_ROWS, 1);
  const rules = [
    {v: "On Track", c: "#D9EAD3"}, {v: "Completed", c: "#B6D7A8"},
    {v: "At Risk", c: "#FFF2CC"}, {v: "Behind", c: "#F4CCCC"}, {v: "Not Started", c: "#EFEFEF"},
  ].map(r => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r.v).setBackground(r.c).setRanges([statusRange]).build());

  const dupRange = sheet.getRange(2, TEAMS_COL.DUP_SRN, MAX_ROWS, 1);
  sheet.setConditionalFormatRules(rules.concat([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("⚠ Duplicate SRN").setBackground("#F4CCCC").setFontColor("#C0392B").setRanges([dupRange]).build(),
  ]));

  // Formulas: only for the rows that actually have data - see the design
  // note at the top of this file for why.
  for (let r = 2; r <= data.length + 1; r++) applyTeamRowFormulas_(sheet, r);

  const protection = sheet.getRange(2, TEAMS_COL.TEAM_SIZE, MAX_ROWS, 2).protect().setDescription("Formula columns - locked");
  protection.setWarningOnly(true);
}

// Sets the Team Size / Duplicate SRN? formulas for ONE row. Called at setup
// time (looped over real rows) and by Code.gs right after a new row is added.
function applyTeamRowFormulas_(sheet, r) {
  const nameStart = colToLetter(TEAMS_COL.NAME_START);
  const nameEnd = colToLetter(TEAMS_COL.NAME_START + MAX_STUDENTS - 1);
  sheet.getRange(r, TEAMS_COL.TEAM_SIZE).setFormula(`=COUNTA(${nameStart}${r}:${nameEnd}${r})`);

  const srnStart = colToLetter(TEAMS_COL.SRN_START);
  const srnEnd = colToLetter(TEAMS_COL.SRN_START + MAX_STUDENTS - 1);
  const rowRange = `${srnStart}${r}:${srnEnd}${r}`;
  const fullRange = `$${srnStart}$2:$${srnEnd}$${MAX_ROWS + 1}`;
  sheet.getRange(r, TEAMS_COL.DUP_SRN).setFormula(
    `=IF(COUNTA(${rowRange})=0,"",IF(SUMPRODUCT((COUNTIF(${fullRange},${rowRange})>1)*(${rowRange}<>""))>0,"⚠ Duplicate SRN",""))`
  );
}

// ==================== TASKS ====================

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

  parseableDueRows.forEach((rowNum) => {
    const cell = sheet.getRange(rowNum + 1, TASKS_COL.DUE);
    const raw = String(cell.getValue());
    const m = raw.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (m) {
      cell.setValue(new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])));
    }
  });
  sheet.getRange(2, TASKS_COL.DUE, MAX_ROWS, 1).setNumberFormat("dd-mm-yyyy");

  const statusRange = sheet.getRange(2, TASKS_COL.STATUS, MAX_ROWS, 1);
  const overdueRange = sheet.getRange(2, TASKS_COL.OVERDUE, MAX_ROWS, 1);
  const rules = [
    {v: "Complete", c: "#B6D7A8"}, {v: "In Progress", c: "#FFF2CC"},
    {v: "Pending", c: "#EFEFEF"}, {v: "Blocked", c: "#F4CCCC"},
  ].map(r => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r.v).setBackground(r.c).setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("⚠ OVERDUE").setBackground("#F4CCCC").setFontColor("#C0392B").setRanges([overdueRange]).build());
  sheet.setConditionalFormatRules(rules);

  for (let r = 2; r <= data.length + 1; r++) applyTaskRowFormulas_(sheet, r);

  sheet.getRange(2, TASKS_COL.DAYS_UNTIL_DUE, MAX_ROWS, 2).protect().setDescription("Formula columns - locked").setWarningOnly(true);
}

function applyTaskRowFormulas_(sheet, r) {
  const dueCol = colToLetter(TASKS_COL.DUE);
  const statusCol = colToLetter(TASKS_COL.STATUS);
  sheet.getRange(r, TASKS_COL.DAYS_UNTIL_DUE).setFormula(`=IF(ISNUMBER(${dueCol}${r}),${dueCol}${r}-TODAY(),"")`);
  sheet.getRange(r, TASKS_COL.OVERDUE).setFormula(`=IF(AND(ISNUMBER(${dueCol}${r}),${dueCol}${r}<TODAY(),${statusCol}${r}<>"Complete"),"⚠ OVERDUE","")`);
}

// ==================== EVALUATION ====================
// Stage 1 (departmental screening): 5 criteria, 1-5 each, 20% weight each
// (so raw sum x4 = /100). 70/100 threshold to advance.
// Stage 2 (C-SHINE vetting panel): qualitative 1-5 on 3 dimensions + verdict.

function buildEvaluationSheet(ss) {
  const headers = [
    "Team ID", "Project Name",
    "Problem Significance (1-5)", "Technical Novelty (1-5)", "Prototype Maturity (1-5)",
    "Team Capability (1-5)", "Commercial / IP Potential (1-5)",
    "Stage 1 Total (/100)", "Stage 1 Result",
    "Market Viability (1-5)", "Scalability (1-5)", "Team Coachability (1-5)",
    "Panel Notes", "Panel Verdict",
  ];
  const sheet = getOrCreateSheet(ss, "Evaluation");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length);
  bandRows(sheet, 2, MAX_ROWS, headers.length);
  protectHeader(sheet, headers.length);

  sheet.setColumnWidth(EVAL_COL.PROJECT, 220);
  sheet.setColumnWidth(EVAL_COL.PANEL_NOTES, 260);
  sheet.getRange(2, EVAL_COL.PANEL_NOTES, MAX_ROWS, 1).setWrap(true);

  [EVAL_COL.PROBLEM, EVAL_COL.NOVELTY, EVAL_COL.PROTOTYPE, EVAL_COL.TEAM_CAP, EVAL_COL.COMMERCIAL,
   EVAL_COL.MARKET, EVAL_COL.SCALABILITY, EVAL_COL.COACHABILITY].forEach((col) => {
    addNumberRangeValidation(sheet, 2, col, MAX_ROWS, 1, 5);
  });
  addDropdown(sheet, 2, EVAL_COL.PANEL_VERDICT, MAX_ROWS, ["Pending", "Admitted", "Not Admitted", "Deferred"]);

  const resultRange = sheet.getRange(2, EVAL_COL.STAGE1_RESULT, MAX_ROWS, 1);
  const verdictRange = sheet.getRange(2, EVAL_COL.PANEL_VERDICT, MAX_ROWS, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Pass").setBackground("#D9EAD3").setRanges([resultRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Below Threshold").setBackground("#F4CCCC").setRanges([resultRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Admitted").setBackground("#D9EAD3").setRanges([verdictRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Not Admitted").setBackground("#F4CCCC").setRanges([verdictRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Deferred").setBackground("#FFF2CC").setRanges([verdictRange]).build(),
  ]);

  sheet.getRange(2, EVAL_COL.PROJECT, MAX_ROWS, 1).protect().setDescription("Formula - locked").setWarningOnly(true);
  sheet.getRange(2, EVAL_COL.STAGE1_TOTAL, MAX_ROWS, 2).protect().setDescription("Formula - locked").setWarningOnly(true);

  // No rows are seeded (no evaluation data exists yet in the source files) -
  // rows get added through the admin app, which applies formulas per row.
}

function applyEvaluationRowFormulas_(sheet, r) {
  const teamIdCol = colToLetter(EVAL_COL.TEAM_ID);
  sheet.getRange(r, EVAL_COL.PROJECT).setFormula(
    `=IFERROR(VLOOKUP(${teamIdCol}${r},Teams!$${colToLetter(TEAMS_COL.ID)}$2:$${colToLetter(TEAMS_COL.PROJECT)}$${MAX_ROWS + 1},${TEAMS_COL.PROJECT - TEAMS_COL.ID + 1},FALSE),"")`
  );
  const c1 = colToLetter(EVAL_COL.PROBLEM), c5 = colToLetter(EVAL_COL.COMMERCIAL);
  sheet.getRange(r, EVAL_COL.STAGE1_TOTAL).setFormula(`=IF(COUNT(${c1}${r}:${c5}${r})<5,"",SUM(${c1}${r}:${c5}${r})*4)`);
  const totalCol = colToLetter(EVAL_COL.STAGE1_TOTAL);
  sheet.getRange(r, EVAL_COL.STAGE1_RESULT).setFormula(`=IF(${totalCol}${r}="","",IF(${totalCol}${r}>=70,"Pass","Below Threshold"))`);
}

// ==================== SURVEYS ====================
// Tracks the Customer Insight Survey deliverable per team (see
// portal/assets/resources/Customer_Insight_Survey.pptx for the training
// material this operationalises).

function buildSurveysSheet(ss) {
  const headers = [
    "Team ID", "Project Name", "Interviews Completed", "Target Interviews",
    "Progress", "Report Link", "Faculty Review Status", "Faculty Comments", "Key Insights Summary",
  ];
  const sheet = getOrCreateSheet(ss, "Surveys");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length);
  bandRows(sheet, 2, MAX_ROWS, headers.length);
  protectHeader(sheet, headers.length);

  sheet.setColumnWidth(SURVEY_COL.PROJECT, 220);
  sheet.setColumnWidth(SURVEY_COL.INSIGHTS_SUMMARY, 260);
  sheet.getRange(2, SURVEY_COL.INSIGHTS_SUMMARY, MAX_ROWS, 1).setWrap(true);
  sheet.getRange(2, SURVEY_COL.FACULTY_COMMENTS, MAX_ROWS, 1).setWrap(true);

  addDropdown(sheet, 2, SURVEY_COL.REVIEW_STATUS, MAX_ROWS, ["Not Started", "Submitted", "Under Review", "Reviewed", "Needs Revision"]);

  const statusRange = sheet.getRange(2, SURVEY_COL.REVIEW_STATUS, MAX_ROWS, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Reviewed").setBackground("#D9EAD3").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Needs Revision").setBackground("#F4CCCC").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Under Review").setBackground("#FFF2CC").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Not Started").setBackground("#EFEFEF").setRanges([statusRange]).build(),
  ]);

  sheet.getRange(2, SURVEY_COL.PROJECT, MAX_ROWS, 1).protect().setDescription("Formula - locked").setWarningOnly(true);
  sheet.getRange(2, SURVEY_COL.PROGRESS, MAX_ROWS, 1).protect().setDescription("Formula - locked").setWarningOnly(true);
}

function applySurveyRowFormulas_(sheet, r) {
  const teamIdCol = colToLetter(SURVEY_COL.TEAM_ID);
  sheet.getRange(r, SURVEY_COL.PROJECT).setFormula(
    `=IFERROR(VLOOKUP(${teamIdCol}${r},Teams!$${colToLetter(TEAMS_COL.ID)}$2:$${colToLetter(TEAMS_COL.PROJECT)}$${MAX_ROWS + 1},${TEAMS_COL.PROJECT - TEAMS_COL.ID + 1},FALSE),"")`
  );
  const doneCol = colToLetter(SURVEY_COL.INTERVIEWS_DONE), targetCol = colToLetter(SURVEY_COL.INTERVIEWS_TARGET);
  sheet.getRange(r, SURVEY_COL.PROGRESS).setFormula(
    `=IF(OR(${doneCol}${r}="",${targetCol}${r}="",${targetCol}${r}=0),"",ROUND(${doneCol}${r}/${targetCol}${r}*100,0)&"%")`
  );
}

// ==================== FACULTY MENTORS (derived, read-only) ====================

function buildFacultyMentorsSheet(ss) {
  const headers = ["Name", "Phone", "Email", "Teams Mentored", "Number of Teams"];
  const sheet = getOrCreateSheet(ss, "Faculty Mentors");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length);
  sheet.setColumnWidth(4, 320);
  sheet.getRange(2, 4, MAX_ROWS, 1).setWrap(true);

  const protection = sheet.protect().setDescription("Derived from Teams tab - edit Teams instead");
  protection.setWarningOnly(true);

  rebuildFacultyMentorsFromTeams(ss);
}

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
      const fe = String(row[TEAMS_COL.FACULTY_EMAIL - 1] || "");
      if (!fm) return;
      const names = fm.split(",").map(s => s.trim()).filter(Boolean);
      const phones = fp.split(",").map(s => s.trim());
      const emails = fe.split(",").map(s => s.trim());
      names.forEach((name, i) => {
        if (!map[name]) map[name] = { phone: phones[i] || "", email: emails[i] || "", teams: [] };
        if (project) map[name].teams.push(project);
      });
    });
  }

  const existingLastRow = facSheet.getLastRow();
  if (existingLastRow > 1) facSheet.getRange(2, 1, existingLastRow - 1, 5).clearContent();
  const outRows = Object.keys(map).map(name => [name, map[name].phone, map[name].email, map[name].teams.join(", "), map[name].teams.length]);
  if (outRows.length) facSheet.getRange(2, 1, outRows.length, 5).setValues(outRows);
}

// ==================== INDUSTRY MENTORS ====================

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

// ==================== OVERVIEW ====================

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
  const evalResultCol = colToLetter(EVAL_COL.STAGE1_RESULT);
  const evalVerdictCol = colToLetter(EVAL_COL.PANEL_VERDICT);
  const surveyStatusCol = colToLetter(SURVEY_COL.REVIEW_STATUS);
  const surveyTeamIdCol = colToLetter(SURVEY_COL.TEAM_ID);

  const rows = [
    ["Total Teams", `=COUNTA(Teams!${teamsProjectCol}2:${teamsProjectCol}${MAX_ROWS + 1})`],
    ["Faculty Mentors", `=COUNTA('Faculty Mentors'!A2:A)`],
    ["Industry Mentors", `=COUNTA('Industry Mentors'!A2:A)`],
    ["Tasks Complete", `=COUNTIF(Tasks!${tasksStatusCol}2:${tasksStatusCol}${MAX_ROWS + 1},"Complete")&" / "&COUNTA(Tasks!${tasksActivityCol}2:${tasksActivityCol}${MAX_ROWS + 1})`],
    ["Overdue Tasks", `=COUNTIF(Tasks!${tasksOverdueCol}2:${tasksOverdueCol}${MAX_ROWS + 1},"⚠ OVERDUE")`],
    ["Teams Missing Faculty Mentor", `=COUNTIFS(Teams!${teamsProjectCol}2:${teamsProjectCol}${MAX_ROWS + 1},"<>",Teams!${teamsMentorCol}2:${teamsMentorCol}${MAX_ROWS + 1},"")`],
    ["Duplicate SRNs Found", `=COUNTIF(Teams!${teamsDupCol}2:${teamsDupCol}${MAX_ROWS + 1},"⚠ Duplicate SRN")`],
    ["Average Team Size", `=IFERROR(ROUND(AVERAGEIF(Teams!${teamsProjectCol}2:${teamsProjectCol}${MAX_ROWS + 1},"<>",Teams!${teamsSizeCol}2:${teamsSizeCol}${MAX_ROWS + 1}),1),0)`],
    ["Teams Passed Stage 1 Screening", `=COUNTIF(Evaluation!${evalResultCol}2:${evalResultCol}${MAX_ROWS + 1},"Pass")`],
    ["Teams Admitted (Stage 2 Panel)", `=COUNTIF(Evaluation!${evalVerdictCol}2:${evalVerdictCol}${MAX_ROWS + 1},"Admitted")`],
    ["Surveys Reviewed", `=COUNTIF(Surveys!${surveyStatusCol}2:${surveyStatusCol}${MAX_ROWS + 1},"Reviewed")&" / "&COUNTA(Surveys!${surveyTeamIdCol}2:${surveyTeamIdCol}${MAX_ROWS + 1})`],
  ];
  sheet.getRange(4, 1, rows.length, 2).setValues(rows);
  sheet.getRange(4, 1, rows.length, 1).setFontWeight("bold");
  sheet.autoResizeColumns(1, 2);
  sheet.getRange(4 + rows.length + 1, 1).setValue("This sheet is the live source of truth for the E-Co-Op portal website. Edit Teams / Tasks / Evaluation / Surveys / Industry Mentors directly here - changes appear on the site automatically.").setFontStyle("italic").setFontColor("#666666").setWrap(true);
  sheet.setColumnWidth(1, 260);
}
