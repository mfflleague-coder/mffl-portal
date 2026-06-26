function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playersSheet = ss.getSheetByName("Players");

  const row = e.values;

  const timestamp = row[0];
  const firstName = row[1];
  const surname = row[2];
  const team = row[6];
  const photo = row[7];

  const route1Answer = row[8];
  const route1Evidence = row[10];

  const route2Answer = row[11];
  const route2Evidence = row[13];

  const route3Answer = row[14];
  const weightKg = row[15];
  const heightCm = row[16];

  const route4Answer = row[17];

  const step6Answer = row[20];

  const step6ClubName = row[24];
  const step6LeagueDivision = row[25];
  const step6Season = row[26];
  const step6Involvement = row[27];
  const step6ExtraDetails = row[28];

  const nextRow = playersSheet.getLastRow() + 1;
  const playerNumber = nextRow - 1;
  const playerId = "MFFL26" + String(playerNumber).padStart(4, "0");

  let eligibilityRoute = "";
  let eligibilityEvidence = "";

  if (route1Answer === "Yes") {
    eligibilityRoute = "Route 1 - Current Weight Loss Programme";
    eligibilityEvidence = route1Evidence || "";
  } else if (route2Answer === "Yes") {
    eligibilityRoute = "Route 2 - Recent Former Member";
    eligibilityEvidence = route2Evidence || "";
  } else if (route3Answer === "Yes") {
    eligibilityRoute = "Route 3 - BMI Eligibility";
  } else if (route4Answer === "Yes") {
    eligibilityRoute = "Route 4 - Existing MFFL Player";
  } else {
    eligibilityRoute = "No Eligible Route";
  }

  let bmi = "";

  if (eligibilityRoute === "Route 3 - BMI Eligibility" && weightKg && heightCm) {
    const heightM = Number(heightCm) / 100;
    bmi = (Number(weightKg) / (heightM * heightM)).toFixed(1);
  }

  let step6Status = "Clear";
  let step6Review = "";
  let step6Note = "";

  if (step6Answer === "Yes" || step6Answer === "Not sure") {
    step6Status = "Review Required";
    step6Review = "View Details";

    step6Note =
      "Club: " + (step6ClubName || "Not provided") + "\n" +
      "League/Division: " + (step6LeagueDivision || "Not provided") + "\n" +
      "Season: " + (step6Season || "Not provided") + "\n" +
      "Involvement: " + (step6Involvement || "Not provided") + "\n" +
      "Extra Details: " + (step6ExtraDetails || "None provided");
  }

  let status = "Pending";

  if (step6Status === "Review Required") {
    status = "Committee Review";
  }

  const fullName = firstName + " " + surname;

  playersSheet.appendRow([
    playerId,
    status,
    timestamp,
    firstName,
    surname,
    fullName,
    team,
    photo,
    eligibilityRoute,
    bmi,
    eligibilityEvidence,
    step6Status,
    step6Review,
    "",
    ""
  ]);

  const addedRow = playersSheet.getLastRow();

  if (step6Review === "View Details") {
    playersSheet.getRange(addedRow, 13).setNote(step6Note);
  }
}
function onEdit(e) {

  const sheet = e.range.getSheet();

  if (sheet.getName() !== "Players") return;

  // Only react when Status (Column B) changes
  if (e.range.getColumn() !== 2) return;

  const row = e.range.getRow();
  const status = e.range.getValue();

  const approvedDateCell = sheet.getRange(row, 14); // Column N

  if (status === "Approved" && approvedDateCell.isBlank()) {
    approvedDateCell.setValue(new Date());
  }

}
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const approved = ss.getSheetByName("Approved Players");
  const teams = ss.getSheetByName("Teams");
  const sessions = ss.getSheetByName("Portal Sessions");
  const audit = ss.getSheetByName("Audit Log");

  const type = e.parameter.type || "";

  if (type === "login") {
    return loginTeam(e, teams, sessions, audit);
  }

  if (type === "session") {
    return getSessionTeam(e, approved, teams, sessions, audit);
  }

  return output({ success: false, error: "Invalid request" });
}


// =========================
// LOGIN
// =========================

function loginTeam(e, teamsSheet, sessionsSheet, auditSheet) {
  const teamId = String(e.parameter.teamId || "").trim();
  const adminCode = String(e.parameter.adminCode || "").trim();

  if (!teamId || !adminCode) {
    logAudit(auditSheet, teamId, "LOGIN_FAILED", "Missing Team ID or Admin Code");
    return output({ success: false, error: "Missing Team ID or Admin Code" });
  }

  const rows = teamsSheet.getDataRange().getValues();
  const data = rows.slice(1);

  const teamRowIndex = data.findIndex(r =>
    String(r[5]).trim().toLowerCase() === teamId.toLowerCase() &&
    String(r[6]).trim() === adminCode &&
    r[9] === true
  );

  if (teamRowIndex === -1) {
    logAudit(auditSheet, teamId, "LOGIN_FAILED", "Invalid login attempt");
    return output({ success: false, error: "Invalid Team ID or Admin Code" });
  }

  const team = data[teamRowIndex];

  const sessionId = Utilities.getUuid();
  const created = new Date();
  const expires = new Date(created.getTime() + 12 * 60 * 60 * 1000);

  sessionsSheet.appendRow([
    sessionId,
    team[5],
    team[2],
    created,
    expires
  ]);

  teamsSheet.getRange(teamRowIndex + 2, 9).setValue(created);

  logAudit(auditSheet, team[5], "LOGIN_SUCCESS", team[0]);

  return output({
    success: true,
    sessionId: sessionId,
    team: {
      name: team[0],
      slug: team[1],
      league: team[2],
      logo: "https://portal.mffl.co.uk/assets/logos/" + team[1] + ".png"
    }
  });
}


// =========================
// SESSION CHECK + TEAM DATA
// =========================

function getSessionTeam(e, playersSheet, teamsSheet, sessionsSheet, auditSheet) {
  const sessionId = String(e.parameter.sessionId || "").trim();

  if (!sessionId) {
    return output({ success: false, error: "Missing session" });
  }

  const sessionRows = sessionsSheet.getDataRange().getValues().slice(1);
  const session = sessionRows.find(r => String(r[0]) === sessionId);

  if (!session) {
    return output({ success: false, error: "Invalid session" });
  }

  const now = new Date();
  const expires = new Date(session[4]);

  if (now > expires) {
    logAudit(auditSheet, session[1], "SESSION_EXPIRED", sessionId);
    return output({ success: false, error: "Session expired" });
  }

  const teamId = session[1];

  const teams = teamsSheet.getDataRange().getValues().slice(1);

  const team = teams.find(r =>
    String(r[5]).trim().toLowerCase() === String(teamId).trim().toLowerCase() &&
    r[9] === true
  );

  if (!team) {
    return output({ success: false, error: "Team not found or disabled" });
  }

  const players = playersSheet
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(r => r[6] === team[0])
    .map(r => ({
      playerId: r[0],
      name: r[5],
      photo: driveImage(r[7]),
      route: r[8],
      showBMI: String(r[8]).startsWith("Route 3"),
      bmi: r[9],
      note: r[14] || ""
    }));

  logAudit(auditSheet, team[5], "VIEW_TEAM", team[0]);

  return output({
    success: true,
    team: {
      name: team[0],
      slug: team[1],
      league: team[2],
      logo: "https://portal.mffl.co.uk/assets/logos/" + team[1] + ".png"
    },
    players: players
  });
}


// =========================
// DRIVE IMAGE
// =========================

function driveImage(url) {
  if (!url) return "";

  const match = String(url).match(/[-\w]{25,}/);

  if (!match) return "";

  return "https://drive.google.com/thumbnail?id=" + match[0] + "&sz=w800";
}


// =========================
// AUDIT LOG
// =========================

function logAudit(sheet, teamId, action, details) {
  sheet.appendRow([
    new Date(),
    teamId || "",
    action,
    details || ""
  ]);
}


// =========================
// OUTPUT JSON
// =========================

function output(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
