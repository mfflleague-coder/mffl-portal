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
  const playerId = "MFFL26" + String(nextRow - 1).padStart(4, "0");

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

  let status = step6Status === "Review Required" ? "Committee Review" : "Pending";
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
  if (e.range.getColumn() !== 2) return;

  const row = e.range.getRow();
  const status = e.range.getValue();

  if (status !== "Approved") return;

  const approvedDateCell = sheet.getRange(row, 14);

  if (approvedDateCell.isBlank()) {
    approvedDateCell.setValue(new Date());
  }

  const playerId = sheet.getRange(row, 1).getValue();
  const photoUrl = sheet.getRange(row, 8).getValue();

  if (playerId && photoUrl) {
    uploadPlayerPhotoToHostinger(playerId, photoUrl);
  }
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const approved = ss.getSheetByName("Approved Players");
  const teams = ss.getSheetByName("Teams");
  const fixtures = ss.getSheetByName("Fixtures");
  const leagueAdmins = ss.getSheetByName("League Admins");
  const sessions = ss.getSheetByName("Portal Sessions");
  const audit = ss.getSheetByName("Audit Log");

  const type = String(e.parameter.type || "");

  if (type === "login") {
    return loginPortal(e, teams, leagueAdmins, sessions, audit);
  }

  if (type === "session") {
    return getSessionDashboard(e, approved, teams, fixtures, sessions, audit);
  }

  if (type === "team") {
    return getTeamBySession(e, approved, teams, sessions, audit);
  }

  return output({ success: false, error: "Invalid request" });
}

function loginPortal(e, teamsSheet, leagueAdminsSheet, sessionsSheet, auditSheet) {
  const userId = String(e.parameter.teamId || "").trim();
  const adminCode = String(e.parameter.adminCode || "").trim();

  if (!userId || !adminCode) {
    logAudit(auditSheet, userId, "LOGIN_FAILED", "Missing ID or code");
    return output({ success: false, error: "Please enter your ID and code." });
  }

  const leagueAdmin = findLeagueAdmin(userId, adminCode, leagueAdminsSheet);

  if (leagueAdmin) {
    const sessionId = createSession(sessionsSheet, {
      subjectId: leagueAdmin.adminId,
      league: leagueAdmin.league,
      role: leagueAdmin.role,
      scope: leagueAdmin.league,
      name: leagueAdmin.name
    });

    updateLeagueAdminLastLogin(leagueAdminsSheet, leagueAdmin.rowNumber);
    logAudit(auditSheet, leagueAdmin.adminId, "LOGIN_SUCCESS", leagueAdmin.role + " - " + leagueAdmin.league);

    return output({
      success: true,
      sessionId,
      role: leagueAdmin.role,
      mode: "league",
      league: leagueAdmin.league,
      name: leagueAdmin.name
    });
  }

  const teamAdmin = findTeamAdmin(userId, adminCode, teamsSheet);

  if (teamAdmin) {
    const sessionId = createSession(sessionsSheet, {
      subjectId: teamAdmin.teamId,
      league: teamAdmin.league,
      role: "Team Admin",
      scope: teamAdmin.slug,
      name: teamAdmin.team
    });

    updateTeamLastLogin(teamsSheet, teamAdmin.rowNumber);
    logAudit(auditSheet, teamAdmin.teamId, "LOGIN_SUCCESS", teamAdmin.team);

    return output({
      success: true,
      sessionId,
      role: "Team Admin",
      mode: "team",
      team: {
        name: teamAdmin.team,
        slug: teamAdmin.slug,
        league: teamAdmin.league,
        logo: logoUrl(teamAdmin.slug)
      }
    });
  }

  logAudit(auditSheet, userId, "LOGIN_FAILED", "Invalid ID or code");
  return output({ success: false, error: "Invalid ID or code." });
}

function findLeagueAdmin(userId, adminCode, sheet) {
  if (!sheet) return null;

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];

    const active = r[5] === true;

    if (
      active &&
      String(r[1]).trim().toLowerCase() === userId.toLowerCase() &&
      String(r[2]).trim() === adminCode
    ) {
      return {
        rowNumber: i + 1,
        name: r[0],
        adminId: String(r[1]).trim(),
        league: String(r[3]).trim(),
        role: String(r[4]).trim() || "League Admin"
      };
    }
  }

  return null;
}

function findTeamAdmin(userId, adminCode, sheet) {
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];

    if (
      r[9] === true &&
      String(r[5]).trim().toLowerCase() === userId.toLowerCase() &&
      String(r[6]).trim() === adminCode
    ) {
      return {
        rowNumber: i + 1,
        team: r[0],
        slug: r[1],
        league: r[2],
        teamId: String(r[5]).trim()
      };
    }
  }

  return null;
}

function getSessionDashboard(e, playersSheet, teamsSheet, fixturesSheet, sessionsSheet, auditSheet) {
  const sessionCheck = validateSession(e, sessionsSheet);

  if (!sessionCheck.success) {
    return output(sessionCheck);
  }

  const session = sessionCheck.session;

  if (session.role === "Team Admin") {
    const team = findTeamByTeamId(teamsSheet, session.subjectId);

    if (!team) {
      return output({ success: false, error: "Team not found or disabled" });
    }

    const players = getPlayersForTeam(playersSheet, team.name);
    const opposition = getTodaysOpposition(fixturesSheet, teamsSheet, playersSheet, team.name);

    logAudit(auditSheet, session.subjectId, "VIEW_TEAM", team.name);

    return output({
      success: true,
      mode: "team",
      role: session.role,
      team,
      players,
      opposition
    });
  }

  const teams = getTeamsForAdmin(teamsSheet, session.league, session.role);

  logAudit(auditSheet, session.subjectId, "VIEW_DASHBOARD", session.role + " - " + session.league);

  return output({
    success: true,
    mode: "league",
    role: session.role,
    league: session.league,
    name: session.name,
    groups: groupTeamsByLeague(teams)
  });
}

function getTeamBySession(e, playersSheet, teamsSheet, sessionsSheet, auditSheet) {
  const sessionCheck = validateSession(e, sessionsSheet);

  if (!sessionCheck.success) {
    return output(sessionCheck);
  }

  const slug = String(e.parameter.slug || "").trim();

  if (!slug) {
    return output({ success: false, error: "Missing team slug" });
  }

  const session = sessionCheck.session;
  const team = findTeamBySlug(teamsSheet, slug);

  if (!team) {
    return output({ success: false, error: "Team not found or disabled" });
  }

  if (session.role === "Team Admin" && team.teamId !== session.subjectId) {
    logAudit(auditSheet, session.subjectId, "ACCESS_DENIED", team.name);
    return output({ success: false, error: "You do not have access to this team." });
  }

  if (session.role !== "Team Admin") {
    const canAccessAll = session.league === "ALL" || session.role === "Super Admin";
    const canAccessLeague = normaliseLeague(team.league) === normaliseLeague(session.league);

    if (!canAccessAll && !canAccessLeague) {
      logAudit(auditSheet, session.subjectId, "ACCESS_DENIED", team.name);
      return output({ success: false, error: "You do not have access to this team." });
    }
  }

  const players = getPlayersForTeam(playersSheet, team.name);

  logAudit(auditSheet, session.subjectId, "VIEW_TEAM", team.name);

  return output({
    success: true,
    mode: "team",
    role: session.role,
    team,
    players
  });
}

function createSession(sheet, data) {
  const sessionId = Utilities.getUuid();
  const created = new Date();
  const expires = new Date(created.getTime() + 12 * 60 * 60 * 1000);

  sheet.appendRow([
    sessionId,
    data.subjectId,
    data.league,
    created,
    expires,
    data.role,
    data.scope,
    data.name
  ]);

  return sessionId;
}

function validateSession(e, sessionsSheet) {
  const sessionId = String(e.parameter.sessionId || "").trim();

  if (!sessionId) {
    return { success: false, error: "Missing session" };
  }

  const rows = sessionsSheet.getDataRange().getValues().slice(1);
  const row = rows.find(r => String(r[0]) === sessionId);

  if (!row) {
    return { success: false, error: "Invalid session" };
  }

  if (new Date() > new Date(row[4])) {
    return { success: false, error: "Session expired" };
  }

  return {
    success: true,
    session: {
      sessionId: row[0],
      subjectId: row[1],
      league: row[2],
      created: row[3],
      expires: row[4],
      role: row[5] || "Team Admin",
      scope: row[6] || "",
      name: row[7] || ""
    }
  };
}

function findTeamByTeamId(sheet, teamId) {
  const rows = sheet.getDataRange().getValues().slice(1);

  const row = rows.find(r =>
    String(r[5]).trim().toLowerCase() === String(teamId).trim().toLowerCase() &&
    r[9] === true
  );

  return row ? teamFromRow(row) : null;
}

function findTeamBySlug(sheet, slug) {
  const rows = sheet.getDataRange().getValues().slice(1);

  const row = rows.find(r =>
    String(r[1]).trim() === String(slug).trim() &&
    r[9] === true
  );

  return row ? teamFromRow(row) : null;
}

function findTeamByName(sheet, teamName) {
  const rows = sheet.getDataRange().getValues().slice(1);

  const row = rows.find(r =>
    String(r[0]).trim() === String(teamName).trim() &&
    r[9] === true
  );

  return row ? teamFromRow(row) : null;
}

function getTeamsForAdmin(sheet, league, role) {
  const rows = sheet.getDataRange().getValues().slice(1);
  const canAccessAll = league === "ALL" || role === "Super Admin";

  return rows
    .filter(r => r[9] === true)
    .filter(r => canAccessAll || normaliseLeague(r[2]) === normaliseLeague(league))
    .map(teamFromRow)
    .sort((a, b) => a.order - b.order);
}

function teamFromRow(r) {
  return {
    name: r[0],
    slug: r[1],
    league: r[2],
    active: r[3],
    order: Number(r[4]) || 999,
    teamId: String(r[5] || "").trim(),
    logo: logoUrl(r[1])
  };
}

function groupTeamsByLeague(teams) {
  const groups = [];

  teams.forEach(team => {
    let group = groups.find(g => g.league === team.league);

    if (!group) {
      group = { league: team.league, teams: [] };
      groups.push(group);
    }

    group.teams.push(team);
  });

  return groups;
}

function getTodaysOpposition(fixturesSheet, teamsSheet, playersSheet, teamName) {
  if (!fixturesSheet) return null;

  const todayKey = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  const rows = fixturesSheet.getDataRange().getValues().slice(1);

  const fixture = rows.find(r => {
    const fixtureKey = getFixtureDateKey(r[0]);

    const active =
      r[4] === true ||
      String(r[4]).toUpperCase() === "TRUE" ||
      String(r[4]).toUpperCase() === "YES";

    const homeTeam = String(r[2] || "").trim();
    const awayTeam = String(r[3] || "").trim();

    return (
      fixtureKey === todayKey &&
      active &&
      (homeTeam === teamName || awayTeam === teamName)
    );
  });

  if (!fixture) return null;

  const homeTeam = String(fixture[2] || "").trim();
  const awayTeam = String(fixture[3] || "").trim();
  const oppositionName = homeTeam === teamName ? awayTeam : homeTeam;

  const opposition = findTeamByName(teamsSheet, oppositionName);

  if (!opposition) return null;

  return {
    team: opposition,
    players: getPlayersForTeam(playersSheet, oppositionName)
  };
}

function getFixtureDateKey(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
  }

  const text = String(value || "").trim();
  const parts = text.split("/");

  if (parts.length === 3) {
    return (
      parts[2] + "-" +
      parts[1].padStart(2, "0") + "-" +
      parts[0].padStart(2, "0")
    );
  }

  return text;
}

function getPlayersForTeam(playersSheet, teamName) {
  return playersSheet
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(r => r[6] === teamName)
    .map(r => ({
      playerId: r[0],
      name: r[5],
      photo: hostedPlayerImage(r[0], r[7]),
      route: r[8],
      showBMI: String(r[8]).startsWith("Route 3"),
      bmi: r[9],
      note: r[14] || ""
    }));
}

function hostedPlayerImage(playerId, fallbackDriveUrl) {
  if (!playerId) return driveImage(fallbackDriveUrl);

  return "https://portal.mffl.co.uk/assets/players/" + playerId + ".jpg";
}

function driveImage(url) {
  if (!url) return "";

  const match = String(url).match(/[-\w]{25,}/);

  if (!match) return "";

  return "https://drive.google.com/thumbnail?id=" + match[0] + "&sz=w800";
}

function uploadPlayerPhotoToHostinger(playerId, photoUrl) {
  const uploadUrl = "https://portal.mffl.co.uk/upload-player-photo.php";
  const uploadKey = "MFFL2026_UPLOAD_9X7K2P";

  const fileIdMatch = String(photoUrl).match(/[-\w]{25,}/);

  if (!fileIdMatch) return;

  const file = DriveApp.getFileById(fileIdMatch[0]);
  const blob = file.getBlob().setName(playerId + ".jpg");

  const payload = {
    key: uploadKey,
    playerId: playerId,
    photo: blob
  };

  UrlFetchApp.fetch(uploadUrl, {
    method: "post",
    payload: payload,
    muteHttpExceptions: true
  });
}

function updateLeagueAdminLastLogin(sheet, rowNumber) {
  sheet.getRange(rowNumber, 7).setValue(new Date());
}

function updateTeamLastLogin(sheet, rowNumber) {
  sheet.getRange(rowNumber, 9).setValue(new Date());
}

function normaliseLeague(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace("championship", "champ");
}

function logoUrl(slug) {
  return "https://portal.mffl.co.uk/assets/logos/" + slug + ".png";
}

function logAudit(sheet, teamId, action, details) {
  sheet.appendRow([
    new Date(),
    teamId || "",
    action,
    details || ""
  ]);
}

function output(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}