/**
 * Circa Pool Manager — Google Apps Script API
 *
 * Deploy as Web App:
 *   1. Open Google Apps Script (script.google.com)
 *   2. Create a new project, paste this code
 *   3. Deploy → New deployment → Web app
 *   4. Execute as: Me | Who has access: Anyone
 *   5. Copy the deployment URL into the Pool Manager Settings
 *
 * Spreadsheet: Vegas Circa Million VIII - Pool Manager
 */

const SHEET_ID = '1EuGOC-89rfE8O3AGvCYTsDOVt7UpNEWlGSs4b7MfgYs';
const MEMBERS = ['Errol', 'Marlon', 'DeMarlo', 'Alan', "D'"];
const MEMBER_ROWS = { 'Errol': 5, 'Marlon': 6, 'DeMarlo': 7, 'Alan': 8, "D'": 9 };

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';
    let result;

    switch (action) {
      case 'getWeekPicks':
        result = getWeekPicks(parseInt(e.parameter.week) || 1);
        break;
      case 'getAllWeeks':
        result = getAllWeeks();
        break;
      case 'getDashboard':
        result = getDashboard();
        break;
      case 'getResults':
        result = getResults(parseInt(e.parameter.week) || 1);
        break;
      case 'ping':
        result = { ok: true, timestamp: new Date().toISOString(), sheetId: SHEET_ID };
        break;
      default:
        result = { error: 'Unknown action. Use: getWeekPicks, getAllWeeks, getDashboard, getResults, ping' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;

    switch (data.action) {
      case 'savePicks':
        result = savePicks(data.week, data.member, data.picks);
        break;
      case 'saveResults':
        result = saveResults(data.week, data.results);
        break;
      case 'saveAllPicks':
        result = saveAllPicks(data.week, data.allPicks);
        break;
      default:
        result = { error: 'Unknown action' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Team name mapping (spreadsheet uses short names like "Chiefs") ───

const TEAM_NAME_TO_ID = {
  'Cardinals': 'ARI', 'Falcons': 'ATL', 'Ravens': 'BAL', 'Bills': 'BUF',
  'Panthers': 'CAR', 'Bears': 'CHI', 'Bengals': 'CIN', 'Browns': 'CLE',
  'Cowboys': 'DAL', 'Broncos': 'DEN', 'Lions': 'DET', 'Packers': 'GB',
  'Texans': 'HOU', 'Colts': 'IND', 'Jaguars': 'JAX', 'Chiefs': 'KC',
  'Chargers': 'LAC', 'Rams': 'LAR', 'Raiders': 'LV', 'Dolphins': 'MIA',
  'Vikings': 'MIN', 'Patriots': 'NE', 'Saints': 'NO', 'Giants': 'NYG',
  'Jets': 'NYJ', 'Eagles': 'PHI', 'Steelers': 'PIT', '49ers': 'SF',
  'Seahawks': 'SEA', 'Buccaneers': 'TB', 'Titans': 'TEN', 'Commanders': 'WAS'
};

const TEAM_ID_TO_NAME = {};
for (const [name, id] of Object.entries(TEAM_NAME_TO_ID)) {
  TEAM_ID_TO_NAME[id] = name;
}

function teamNameToId(name) {
  if (!name || typeof name !== 'string') return null;
  name = name.trim();
  if (TEAM_NAME_TO_ID[name]) return TEAM_NAME_TO_ID[name];
  if (TEAM_ID_TO_NAME[name]) return name;
  for (const [tName, tId] of Object.entries(TEAM_NAME_TO_ID)) {
    if (name.toLowerCase().includes(tName.toLowerCase())) return tId;
  }
  return name || null;
}

function teamIdToName(id) {
  if (!id) return '';
  return TEAM_ID_TO_NAME[id] || id;
}

// ─── Data Reading ───

function getWeekPicks(weekNum) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Week ' + weekNum);
  if (!sheet) return { error: 'Sheet "Week ' + weekNum + '" not found' };

  const picks = {};
  const range = sheet.getRange('A5:F9');
  const values = range.getValues();

  values.forEach(function(row) {
    const member = (row[0] || '').toString().trim();
    if (member && MEMBERS.indexOf(member) !== -1) {
      picks[member] = [];
      for (var i = 1; i <= 5; i++) {
        var val = (row[i] || '').toString().trim();
        picks[member].push(teamNameToId(val));
      }
    }
  });

  var submissionPicks = [];
  try {
    var subRange = sheet.getRange('N5:O9');
    var subValues = subRange.getValues();
    subValues.forEach(function(row) {
      submissionPicks.push({
        team: teamNameToId((row[0] || '').toString().trim()),
        score: row[1] || 0
      });
    });
  } catch (e) {}

  return {
    week: weekNum,
    picks: picks,
    submissionPicks: submissionPicks,
    timestamp: new Date().toISOString()
  };
}

function getAllWeeks() {
  var allWeeks = {};
  for (var w = 1; w <= 18; w++) {
    var data = getWeekPicks(w);
    if (!data.error) {
      allWeeks[w] = data;
    }
  }
  return { weeks: allWeeks, timestamp: new Date().toISOString() };
}

function getDashboard() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Dashboard');
  if (!sheet) return { error: 'Dashboard sheet not found' };

  var result = {
    wins: 0, losses: 0, pushes: 0, winPct: 0,
    avgPts: 0, totalPts: 0, timestamp: new Date().toISOString()
  };

  try {
    result.wins = sheet.getRange('B11').getValue() || 0;
    result.losses = sheet.getRange('B12').getValue() || 0;
    result.pushes = sheet.getRange('B13').getValue() || 0;
    result.winPct = sheet.getRange('E11').getValue() || 0;
    result.avgPts = sheet.getRange('E12').getValue() || 0;
    result.totalPts = sheet.getRange('E13').getValue() || 0;
  } catch (e) {}

  return result;
}

function getResults(weekNum) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var resultsSheet = ss.getSheetByName('Results');
  if (!resultsSheet) return { error: 'Results sheet not found', week: weekNum };

  var results = {};
  try {
    var data = resultsSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var week = parseInt(row[0]);
      if (week === weekNum) {
        var team = teamNameToId((row[1] || '').toString().trim());
        var result = (row[2] || '').toString().trim().toLowerCase();
        if (team && result) {
          results[team] = result;
        }
      }
    }
  } catch (e) {}

  return { week: weekNum, results: results, timestamp: new Date().toISOString() };
}

// ─── Data Writing ───

function savePicks(weekNum, member, picks) {
  if (!member || MEMBERS.indexOf(member) === -1) {
    return { error: 'Unknown member: ' + member };
  }
  if (!picks || !Array.isArray(picks) || picks.length !== 5) {
    return { error: 'Picks must be an array of 5 team IDs' };
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Week ' + weekNum);
  if (!sheet) return { error: 'Sheet "Week ' + weekNum + '" not found' };

  var row = MEMBER_ROWS[member];
  if (!row) return { error: 'No row mapping for member: ' + member };

  for (var i = 0; i < 5; i++) {
    var teamName = teamIdToName(picks[i]) || '';
    sheet.getRange(row, 2 + i).setValue(teamName);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    week: weekNum,
    member: member,
    timestamp: new Date().toISOString()
  };
}

function saveAllPicks(weekNum, allPicks) {
  if (!allPicks || typeof allPicks !== 'object') {
    return { error: 'allPicks must be an object of member → picks' };
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Week ' + weekNum);
  if (!sheet) return { error: 'Sheet "Week ' + weekNum + '" not found' };

  var saved = [];
  for (var member in allPicks) {
    if (MEMBERS.indexOf(member) === -1) continue;
    var picks = allPicks[member];
    if (!Array.isArray(picks)) continue;

    var row = MEMBER_ROWS[member];
    for (var i = 0; i < 5; i++) {
      var teamName = teamIdToName(picks[i]) || '';
      sheet.getRange(row, 2 + i).setValue(teamName);
    }
    saved.push(member);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    week: weekNum,
    savedMembers: saved,
    timestamp: new Date().toISOString()
  };
}

function saveResults(weekNum, results) {
  if (!results || typeof results !== 'object') {
    return { error: 'results must be an object of teamId → result' };
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Dashboard');
  if (!sheet) return { error: 'Dashboard not found' };

  return {
    success: true,
    week: weekNum,
    note: 'Results are calculated from the sheet formulas. Manual dashboard update may be needed.',
    timestamp: new Date().toISOString()
  };
}
