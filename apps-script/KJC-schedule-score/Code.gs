function getRoster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var divisions = ['Venti', 'Grande', 'Tall'];
  var roster = {}; // name -> {division, group, logo}
  divisions.forEach(function (divName) {
    var sheet = ss.getSheetByName(divName);
    if (!sheet) return;
    var display = sheet.getDataRange().getDisplayValues();
    var headers = display.shift();
    var idx = {};
    headers.forEach(function (h, i) { idx[h] = i; });
    display.forEach(function (row) {
      var name = row[idx['Team Name']];
      if (!name) return;
      roster[name] = { division: divName, group: row[idx['Group']], logo: row[idx['Logo']] };
    });
  });
  return roster;
}

function publishTeams() {
  var roster = getRoster();
  var teams = Object.keys(roster).map(function (name) {
    var r = roster[name];
    return { name: name, division: r.division, group: r.group, logo: r.logo };
  });
  pushToGithub(JSON.stringify(teams, null, 2), 'teams.json');
}

function logStatus(message) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Status') || ss.insertSheet('Status');
  sheet.getRange('A1').setValue('Last publishSchedule run:');
  sheet.getRange('A2').setValue(message);
}

function publishSchedule() {
  var roster = getRoster();
  var settings = getSettings();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var display = range.getDisplayValues();
  var headers = values.shift();
  display.shift();
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  var errors = [];
  var games = [];

  values.forEach(function (row, i) {
    var disp = display[i];
    var homeTeam = disp[idx['Home Team']];
    if (!homeTeam) return;

    var rowNum = i + 2;
    var division = disp[idx['Division']];
    var group = disp[idx['Group']];
    var awayTeam = disp[idx['Away Team']];

    if (homeTeam === awayTeam) {
      errors.push('Row ' + rowNum + ': Home Team and Away Team are the same (' + homeTeam + ')');
    }
    [['Home Team', homeTeam], ['Away Team', awayTeam]].forEach(function (pair) {
      var label = pair[0], name = pair[1];
      if (!name || name === 'TBD') return;
      var r = roster[name];
      if (!r) {
        errors.push('Row ' + rowNum + ': ' + label + ' "' + name + '" not found in any division roster tab');
        return;
      }
      if (r.division !== division) {
        errors.push('Row ' + rowNum + ': ' + label + ' "' + name + '" is rostered in ' + r.division + ', but row says Division = ' + division);
      }
      if (r.group !== group) {
        errors.push('Row ' + rowNum + ': ' + label + ' "' + name + '" is rostered in ' + r.group + ', but row says Group = ' + group);
      }
    });

    games.push({
      division: division,
      group: group,
      time: disp[idx['Time']],
      field: disp[idx['Field']],
      home_team: homeTeam,
      away_team: awayTeam,
      home_score: row[idx['Home Score']] === '' ? null : Number(row[idx['Home Score']]),
      away_score: row[idx['Away Score']] === '' ? null : Number(row[idx['Away Score']]),
      status: disp[idx['Status']],
    });
  });

  if (errors.length > 0) {
    logStatus('FAILED at ' + new Date().toLocaleString() + ':\n' + errors.join('\n'));
    return;
  }

  // publishSchedule()
  pushToGithub(
    JSON.stringify({
      updated_at: new Date().toISOString(),
      enabled: settings.enabled,
      games: games
    }, null, 2),
    'schedule.json'
  );
  logStatus('OK — published at ' + new Date().toLocaleString());
}

function pushToGithub(content, path) {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  var owner = 'sidbai';
  var repo = 'kingjuan-assets';
  var branch = 'main';
  var apiUrl = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + path + '?ref=' + branch;

  var sha = null;
  var getRes = UrlFetchApp.fetch(apiUrl, { headers: { Authorization: 'token ' + token }, muteHttpExceptions: true });
  if (getRes.getResponseCode() === 200) sha = JSON.parse(getRes.getContentText()).sha;

  var body = {
    message: 'Update ' + path,
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch: branch,
  };
  if (sha) body.sha = sha;

  var putRes = UrlFetchApp.fetch('https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + path, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: 'token ' + token },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  if (putRes.getResponseCode() >= 300) throw new Error('GitHub push failed: ' + putRes.getContentText());
}

function formatTime(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'h:mm a');
  }
  return String(value);
}

function getSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return { enabled: true }; // no Settings tab yet = show by default
  var value = sheet.getRange('B1').getValue();
  return { enabled: value === true || value === 'TRUE' };
}
