/**
 * E-Co-Op Program Setup Script
 * KLE Technological University | C-SHINE
 *
 * One-time setup: builds a professionally formatted, validated, and
 * partially locked Google Sheet workbook seeded with the real program data
 * (teams, tasks, faculty mentors, industry mentors).
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

const BRAND_COLOR = "#0B3D91";      // header background
const BRAND_TEXT = "#FFFFFF";
const BAND_COLOR = "#F1F5FB";

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  buildTeamsSheet(ss);
  buildTasksSheet(ss);
  buildFacultyMentorsSheet(ss);
  buildIndustryMentorsSheet(ss);
  buildOverviewSheet(ss);

  // Remove the default empty "Sheet1" if it's still there and empty
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
    const rules = sheet.getConditionalFormatRules();
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
  const range = sheet.getRange(startRow, 1, numRows, numCols);
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
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

// ---------------- TEAMS ----------------
function buildTeamsSheet(ss) {
  const headers = ["ID", "Department", "Project Name", "Students (Name / SRN / Phone)",
    "Faculty Mentor", "Faculty Phone", "End Users", "Description", "Excitement / Interest", "Team Status"];
  const data = [
  [1, "Computer Science and Engineering", "Sintex Cleanliness Detection Using IoT and AI Model", "Subramanya Tiluvalli (01FE23BCS279, 9110607686)\nPraveen Angadi (01FE23BCS298, 9972436163)\nBasavant Mahalingpur (01FE23BCS272, 9353944089)\nRohit Reddy (01FE23BCS293, 9731590011)", "Geeta Sannakki", "9620161698", "Apartment Complexes", "A box with a camera that is fitted in the overhead tank and connects with internet to send notification either through whatsapp or telegram to apartment or big building complexes.", "High and are interested to continue till productization", ""],
  [2, "Computer Science and Engineering", "AI-Enabled eCommerce for Retailers", "Vishal Naik (01FE23BCS294, 6366349040)\nMadhura (01FE23BCS103, 7019325490)\nPratham Kathare (01FE23BCS085, 8277777747)\nAditya Naik (01FE23BCS211, 8762539424)", "Lalita Madanabhavi", "8050458109", "Local Vendors and Customers", "An e-commerce software platform for enabling business for local vendors and customers", "High and are interested to continue till productization", ""],
  [3, "Computer Science and Engineering", "Integrated academic performance monitoring system", "Anagha Nadgouda (01FE23BCS076, 99804 54365)\nSara Patil (01FE23BCS050, 9663081911)\nZaid Momin (01FE23BCS121, 7975332473)\nSushant Maheshwari (01FE23BCS211, 8050003427)", "Lalita Madanabhavi", "8050458109", "Students, Universities and colleges", "A recommendation system based on students' portfolios and trend analysis.", "High and are interested to continue till productization. There is a disconnect in the project title and what they are doing.", ""],
  [4, "Computer Science and Engineering", "Decentralized Academic Credential Verification System using Self-Sovereign Identity", "Swateja Patil (01FE23BCS161, 9900657272)\nPriyanka N D (01FE23BCS184, 8073333610)", "Pooja Shettar", "9964141448", "Academic Institutions, employee verification required institutions", "A credential verification system built using blockchain", "High and are interested to continue till productization", ""],
  [5, "Mechanical Engineering", "Camera based navigation system for visually impaired", "Basuraj Shibargatti (01FE23BME120, 8660128071)\nAkash Halesh Walad (01FE23BME133, 7411445754)\nVrushab Kadam (01FE23BME139, 9535188782)\nSachin Channaveer (01FE23BME149, 7019541358)\nManoj Rode (01FE23BME152, 7406342310)\nSurajkumar B (01FE23BME106, 9036008183)", "Gururaj Fattepur", "9739461325", "Blind people", "A handheld device that can assist blind people during walking", "High", ""],
  [6, "Mechanical Engineering", "Adaptive sujok reflexology-based acupressure therapy device", "Swaroop Biradar (01FE23BME054, 6361159255)\nVivek Padi (01FE23BME058, 8431978516)\nVishal Chavannavar (01FE23BME080, 9449570463)\nShivaraj Hiremath (01FE23BME081, 9960981497)\nSiddamma Meti (01FE23BME141, 7975690922)\nGayatri Angadi (01FE23BME035, 9449101661)", "Gururaj Fattepur", "9739461325", "Therapists and customers who prefer reflexology", "A machine that can detect the palm and appropriately apply pressure based on SUJOK reflexology techniques", "HIgh", ""],
  [7, "Mechanical Engineering", "Smart shopping cart", "Teerth S. Kulkarni (01FE24BME405, 9019846034)\nAnjali V. Purohit (01FE23BME068, 8554005888)\nSejal V. Komalapur (01FE24BME430, 7348921572)\nKarthik G. Nadurmath (01FE24BME424, 7019709717)\nManjunath P. Badiger (01FE24BME435, 8971709642)", "Veerabhadrayya Hiremath, Nagaraj Ekbote", "9900506604, 9591017854", "Shopping centers or shopping malls where shopping carts are extensively used", "A shopping cart with a display, item identification, and a payment portal", "High", ""],
  [8, "Mechanical Engineering", "Smart Servo-Based Electromechanical Disc Lock for two wheelers", "VRUSHABRAJ KALYANKAR (01FE21BME014, 7899424187)\nRAHUL S HALADANDIMATH (01FE22BME050, 6362568936)\nTEJAS HINDASAGERI (01FE22BME416, 9380244302)\nADIT V PATIL (01FE22BME427, 7406216111)\nROHIT MADAR (01FE23BME076, 9686523466)\nVRASHABHA S VASTRAD (01FE23BME108, 9606742193)", "Praveen Petkar, Mantesh Choukimath", "9964476030, 7760072120", "All Bike users", "A disc brake lock unit", "High", ""],
  [9, "Automation and Robotics", "Automated PCB Defect Detection System Using Computer Vision", "Shreya Chaligeri (01FE23BAR017, 7204724565)\nRama Kulkarni (01FE23BAR052)\nAbhishek Hiremath (01FE24BAR402)\nManjunath Sali (01FE24BAR407)", "Vinod Meti", "9986356557", "SME Industries, AEQUS", "A machine which can pass the PCBs automatically and detect the defects and finally sort the PCBs as good and bad ones.", "Excited and willing to continue till productization", ""],
  [10, "Automation and Robotics", "Vision-Based Detection and Counting of Cookware in the Consumer Goods Industry", "Harsha Kampli (01FE23BAR054, 9152493660)\nSofiya Kittur (01FE23BAR013)\nUsaid Mansabdar (01FE23BAR041)\nVirupaksha Amarshetti (01FE23BAR039)\nGaurav Sooji (01FE20BAR002)", "Vijay Mahantesh", "9980205502", "Aequs, SME Industries - Large Industries", "A software that can visually count and classify different types of vessels or ccokware", "Excited and willing to continue till productization", ""],
  [11, "Electrical and Electronics Engineering", "Dual-Output Quadratic Boost Converter with Enhanced Voltage Conversion Ratio.", "Dipti.S (01fe23bee005, 6360873027)\nSanjana.I.K (01fe23bee010, 9945840123)\nRakshita.S.B (01fe23bee023, 7619125343)\nAmitvikram yeri (01fe23bee024, 9110260512)", "Vinod Patil", "9880998964", "EV Charging stations, high voltage battery charging", "They are still thinking on this aspect", "High", ""],
  [12, "Electrical and Electronics Engineering", "Position and Trajectory Control of a 2-DOF Robotic Arm.", "Shivkumar G (01fe22bee082, 6361406074)\nAbhishek wadeyar (01fe23bee103, 7619457369)\nAkash Kempanavvar (01fe23bee086, 7619303794)", "", "", "SME industries such as PCB handling, which require smooth and soft handling of the items", "No clarity on this part", "They are excited. But they are not sure what their product idea would be by the end of 7th sem.", ""]
];

  const sheet = getOrCreateSheet(ss, "Teams");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (data.length) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  styleHeader(sheet, headers.length);
  bandRows(sheet, 2, Math.max(data.length, 1), headers.length);
  protectHeader(sheet, headers.length);

  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 260);
  sheet.getRange(2, 4, Math.max(data.length, 200), 1).setWrap(true);
  sheet.getRange(2, 8, Math.max(data.length, 200), 1).setWrap(true);

  const maxRows = 200;
  addDropdown(sheet, 2, 10, maxRows, ["Not Started", "On Track", "At Risk", "Behind", "Completed"]);

  // Conditional formatting on Team Status
  const statusRange = sheet.getRange(2, 10, maxRows, 1);
  const rules = [
    {v: "On Track", c: "#D9EAD3"},
    {v: "Completed", c: "#B6D7A8"},
    {v: "At Risk", c: "#FFF2CC"},
    {v: "Behind", c: "#F4CCCC"},
    {v: "Not Started", c: "#EFEFEF"},
  ].map(r => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(r.v)
    .setBackground(r.c)
    .setRanges([statusRange])
    .build());
  sheet.setConditionalFormatRules(rules);
}

// ---------------- TASKS ----------------
function buildTasksSheet(ss) {
  const headers = ["Sl. No", "Task / Activity", "Details / Link", "Due Date", "Owner", "Status", "Outcomes", "Remarks"];
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

  const sheet = getOrCreateSheet(ss, "Tasks");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (data.length) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  styleHeader(sheet, headers.length);
  bandRows(sheet, 2, Math.max(data.length, 1), headers.length);
  protectHeader(sheet, headers.length);

  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(8, 260);
  sheet.getRange(2, 2, Math.max(data.length, 200), 1).setWrap(true);
  sheet.getRange(2, 8, Math.max(data.length, 200), 1).setWrap(true);

  const maxRows = 200;
  addDropdown(sheet, 2, 6, maxRows, ["Pending", "In Progress", "Complete", "Blocked"]);

  const statusRange = sheet.getRange(2, 6, maxRows, 1);
  const rules = [
    {v: "Complete", c: "#B6D7A8"},
    {v: "In Progress", c: "#FFF2CC"},
    {v: "Pending", c: "#EFEFEF"},
    {v: "Blocked", c: "#F4CCCC"},
  ].map(r => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(r.v)
    .setBackground(r.c)
    .setRanges([statusRange])
    .build());
  sheet.setConditionalFormatRules(rules);

  // Highlight the two rows whose original date was ambiguous
  const dataVals = sheet.getRange(2, 4, data.length, 1).getValues();
  for (let i = 0; i < dataVals.length; i++) {
    const remarksCell = sheet.getRange(2 + i, 8);
    if (String(remarksCell.getValue()).indexOf("VERIFY DATE") !== -1) {
      sheet.getRange(2 + i, 4).setBackground("#FCE8B2").setNote("Original file had an ambiguous date format here - please confirm the actual due date.");
    }
  }
}

// ---------------- FACULTY MENTORS (derived, read-only) ----------------
function buildFacultyMentorsSheet(ss) {
  const headers = ["Name", "Phone", "Teams Mentored", "Number of Teams"];
  const data = [
  ["Geeta Sannakki", "9620161698", "Sintex Cleanliness Detection Using IoT and AI Model", 1],
  ["Lalita Madanabhavi", "8050458109", "AI-Enabled eCommerce for Retailers, Integrated academic performance monitoring system", 2],
  ["Pooja Shettar", "9964141448", "Decentralized Academic Credential Verification System using Self-Sovereign Identity", 1],
  ["Gururaj Fattepur", "9739461325", "Camera based navigation system for visually impaired, Adaptive sujok reflexology-based acupressure therapy device", 2],
  ["Veerabhadrayya Hiremath", "9900506604", "Smart shopping cart", 1],
  ["Nagaraj Ekbote", "9591017854", "Smart shopping cart", 1],
  ["Praveen Petkar", "9964476030", "Smart Servo-Based Electromechanical Disc Lock for two wheelers", 1],
  ["Mantesh Choukimath", "7760072120", "Smart Servo-Based Electromechanical Disc Lock for two wheelers", 1],
  ["Vinod Meti", "9986356557", "Automated PCB Defect Detection System Using Computer Vision", 1],
  ["Vijay Mahantesh", "9980205502", "Vision-Based Detection and Counting of Cookware in the Consumer Goods Industry", 1],
  ["Vinod Patil", "9880998964", "Dual-Output Quadratic Boost Converter with Enhanced Voltage Conversion Ratio.", 1]
];

  const sheet = getOrCreateSheet(ss, "Faculty Mentors");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (data.length) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  styleHeader(sheet, headers.length);
  bandRows(sheet, 2, Math.max(data.length, 1), headers.length);
  sheet.setColumnWidth(3, 320);
  sheet.getRange(2, 3, Math.max(data.length, 1), 1).setWrap(true);

  // Whole sheet is derived from Teams - protect with a warning rather than a hard lock
  const protection = sheet.protect().setDescription("Derived from Teams tab - edit Teams instead");
  protection.setWarningOnly(true);
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
  if (data.length) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  styleHeader(sheet, headers.length);
  bandRows(sheet, 2, Math.max(data.length, 1), headers.length);
  protectHeader(sheet, headers.length);

  const maxRows = 100;
  addDropdown(sheet, 2, 6, maxRows, ["Yes", "No", "Maybe", "Not yet contacted"]);
}

// ---------------- OVERVIEW ----------------
function buildOverviewSheet(ss) {
  const sheet = getOrCreateSheet(ss, "Overview");
  sheet.getRange(1, 1).setValue("E-Co-Op Program Tracker").setFontSize(18).setFontWeight("bold");
  sheet.getRange(2, 1).setValue("CSHINE, KLE Technological University").setFontSize(11).setFontColor("#666666");
  sheet.getRange(4, 1, 1, 2).setValues([["Total Teams", "=COUNTA(Teams!A2:A)"]]);
  sheet.getRange(5, 1, 1, 2).setValues([["Faculty Mentors", "=COUNTA('Faculty Mentors'!A2:A)"]]);
  sheet.getRange(6, 1, 1, 2).setValues([["Industry Mentors", "=COUNTA('Industry Mentors'!A2:A)"]]);
  sheet.getRange(7, 1, 1, 2).setValues([["Tasks Complete", "=COUNTIF(Tasks!F2:F,\"Complete\")&\" / \"&COUNTA(Tasks!B2:B)"]]);
  sheet.getRange(4, 1, 4, 1).setFontWeight("bold");
  sheet.autoResizeColumns(1, 2);
  sheet.getRange(9, 1).setValue("This sheet is the live source of truth for the E-Co-Op portal website. Edit Teams / Tasks / Industry Mentors directly here - changes appear on the site automatically.").setFontStyle("italic").setFontColor("#666666").setWrap(true);
  sheet.setColumnWidth(1, 260);
}
