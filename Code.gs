const RATINGS_SHEET = 'ratings';
const WEIGHTS_SHEET = 'weights';
const RATINGS_LOG_SHEET = 'ratings_log';
const WEIGHTS_LOG_SHEET = 'weights_log';
const RATINGS_ROLLUP_SHEET = 'ratings_rollup';
const WEIGHTS_ROLLUP_SHEET = 'weights_rollup';
const CONTRIBUTION_ALERT_RECIPIENTS = ['tomrholman@gmail.com'];
const CONTRIBUTION_ALERT_PREVIEW_LIMIT = 10;

const HEADER_ROW = 2;
const DATA_START_ROW = 3;
const MAX_VISIBLE_ROWS = 100;

const REQUIRED_META_HEADERS = ['name', 'team', 'position'];

const ATTRIBUTE_GROUPS = {
  'Show all': [
    'speed',
    'agility',
    'strength',
    'stamina',
    'tackling',
    'breakdown',
    'defensiveIq',
    'ballHandling',
    'contactSkill',
    'attackingIq',
    'inPlayKicking',
    'goalKicking',
    'scrummaging',
    'lineoutThrowing',
    'lineoutJumping'
  ],
  'Physical': [
    'speed',
    'agility',
    'strength',
    'stamina'
  ],
  'Defence': [
    'tackling',
    'breakdown',
    'defensiveIq'
  ],
  'Attack': [
    'ballHandling',
    'contactSkill',
    'attackingIq'
  ],
  'Technical skills': [
    'inPlayKicking',
    'goalKicking',
    'scrummaging',
    'lineoutThrowing',
    'lineoutJumping'
  ]
};

const ATTRIBUTE_LABELS = {
  speed: 'Speed',
  agility: 'Agility',
  strength: 'Strength',
  stamina: 'Stamina',
  tackling: 'Tackling',
  breakdown: 'Breakdown',
  defensiveIq: 'Defensive IQ',
  ballHandling: 'Ball handling',
  contactSkill: 'Contact skill',
  attackingIq: 'Attacking IQ',
  inPlayKicking: 'In-play kicking',
  goalKicking: 'Goal kicking',
  scrummaging: 'Scrummaging',
  lineoutThrowing: 'Lineout throwing',
  lineoutJumping: 'Lineout jumping'
};

const ATTRIBUTE_LABEL_TO_KEY = Object.keys(ATTRIBUTE_LABELS)
  .reduce((acc, key) => {
    acc[ATTRIBUTE_LABELS[key]] = key;
    return acc;
  }, {});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Forge Rugby')
    .addItem('Open rating workspace', 'showSidebar')
    // .addItem('Initialise submission sheets', 'initialiseSheets')
    // .addItem('Rebuild submissions summary', 'rebuildSummary')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Forge Rugby Rating Workspace');
  SpreadsheetApp.getUi().showSidebar(html);
}

function initialiseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const submissions = ss.getSheetByName(RATINGS_LOG_SHEET);
  if (!submissions) throw new Error(`Sheet "${RATINGS_LOG_SHEET}" not found.`);

  submissions.clear();
  submissions.getRange(1, 1, 1, 10).setValues([[
    'submittedAt',
    'userEmail',
    'playerName',
    'team',
    'position',
    'attribute',
    'baselineValue',
    'suggestedValue',
    'delta',
    'sessionId'
  ]]);

  const summary = ensurePlayerSummarySheet_(ss);
  summary.clear();
  summary.getRange(1, 1, 1, 10).setValues([[
    'playerName',
    'team',
    'position',
    'attribute',
    'submissionCount',
    'baselineValue',
    'avgSuggestedValue',
    'medianSuggestedValue',
    'avgDelta',
    'latestSubmittedAt'
  ]]);

  const attributeWeightSubmissions = ensureAttributeWeightSubmissionsSheet_(ss);
  attributeWeightSubmissions.clear();
  attributeWeightSubmissions.getRange(1, 1, 1, 9).setValues([[
    'submittedAt',
    'userEmail',
    'attribute',
    'attributeKey',
    'positionGroup',
    'baselineValue',
    'suggestedValue',
    'delta',
    'sessionId'
  ]]);

  const attributeWeightSummary = ensureAttributeWeightSummarySheet_(ss);
  attributeWeightSummary.clear();
  attributeWeightSummary.getRange(1, 1, 1, 9).setValues([[
    'attribute',
    'attributeKey',
    'positionGroup',
    'submissionCount',
    'baselineValue',
    'avgSuggestedValue',
    'medianSuggestedValue',
    'avgDelta',
    'latestSubmittedAt'
  ]]);

  SpreadsheetApp.getUi().alert('Submission sheets initialised.');
}

function getMasterData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(RATINGS_SHEET);
  if (!sheet) throw new Error(`Sheet "${RATINGS_SHEET}" not found.`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < DATA_START_ROW) {
    throw new Error(`Sheet "${RATINGS_SHEET}" has no player data.`);
  }

  const headerValues = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0]
    .map(h => String(h).trim());

  const dataRows = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, lastCol).getValues();

  const missing = REQUIRED_META_HEADERS.filter(h => !headerValues.includes(h));
  if (missing.length) {
    throw new Error(`Missing required columns in "${RATINGS_SHEET}": ${missing.join(', ')}`);
  }

  const metaIndexes = {};
  REQUIRED_META_HEADERS.forEach(h => metaIndexes[h] = headerValues.indexOf(h));

  const headerIndexMap = {};
  headerValues.forEach((h, idx) => {
    headerIndexMap[h] = idx;
  });

  return {
    headers: headerValues,
    dataRows,
    metaIndexes,
    headerIndexMap
  };
}

function getWorkspaceConfig() {
  const { dataRows, metaIndexes, headerIndexMap } = getMasterData_();

  const teamSet = new Set();
  const positionSet = new Set();
  const nationalitySet = new Set();
  const nationalityIndex = Object.prototype.hasOwnProperty.call(headerIndexMap, 'nationality')
    ? headerIndexMap.nationality
    : -1;

  dataRows.forEach(row => {
    const team = row[metaIndexes.team];
    const position = row[metaIndexes.position];
    const nationality = nationalityIndex >= 0 ? row[nationalityIndex] : '';
    if (team !== '') teamSet.add(String(team));
    if (position !== '') positionSet.add(String(position));
    if (nationality !== '') nationalitySet.add(String(nationality));
  });

  const availableAttributes = ATTRIBUTE_GROUPS['Show all']
    .filter(attr => Object.prototype.hasOwnProperty.call(headerIndexMap, attr));

  return {
    totalPlayers: dataRows.length,
    teams: Array.from(teamSet).sort(),
    positions: Array.from(positionSet).sort(),
    nationalities: Array.from(nationalitySet).sort(),
    attributeGroups: Object.keys(ATTRIBUTE_GROUPS),
    availableAttributes: availableAttributes.map(attr => ({
      key: attr,
      label: ATTRIBUTE_LABELS[attr] || attr
    }))
  };
}

function loadWorkspaceRows(filters) {
  const { dataRows, metaIndexes, headerIndexMap } = getMasterData_();

  const teamFilter = filters && filters.team ? String(filters.team) : 'All';
  const positionFilter = filters && filters.position ? String(filters.position) : 'All';
  const nationalityFilter = filters && filters.nationality ? String(filters.nationality) : 'All';
  const nameSearch = filters && filters.nameSearch ? String(filters.nameSearch).trim().toLowerCase() : '';
  const attributeGroup = filters && filters.attributeGroup ? String(filters.attributeGroup) : 'Show all';
  const nationalityIndex = Object.prototype.hasOwnProperty.call(headerIndexMap, 'nationality')
    ? headerIndexMap.nationality
    : -1;

  const selectedAttributes = (ATTRIBUTE_GROUPS[attributeGroup] || ATTRIBUTE_GROUPS['Show all'])
    .filter(attr => Object.prototype.hasOwnProperty.call(headerIndexMap, attr));

  let filteredRows = dataRows.filter(row => {
    const name = String(row[metaIndexes.name] || '');
    const team = String(row[metaIndexes.team] || '');
    const position = String(row[metaIndexes.position] || '');
    const nationality = nationalityIndex >= 0 ? String(row[nationalityIndex] || '') : '';

    const matchesTeam = teamFilter === 'All' || team === teamFilter;
    const matchesPosition = positionFilter === 'All' || position === positionFilter;
    const matchesNationality = nationalityFilter === 'All' || nationality === nationalityFilter;
    const matchesName = !nameSearch || name.toLowerCase().includes(nameSearch);

    return matchesTeam && matchesPosition && matchesNationality && matchesName;
  });

  const totalFilteredPlayers = filteredRows.length;
  const wasCapped = totalFilteredPlayers > MAX_VISIBLE_ROWS;
  filteredRows = filteredRows.slice(0, MAX_VISIBLE_ROWS);

  const rows = filteredRows.map(row => {
    const obj = {
      name: row[metaIndexes.name],
      team: row[metaIndexes.team],
      position: row[metaIndexes.position]
    };

    selectedAttributes.forEach(attr => {
      obj[attr] = row[headerIndexMap[attr]];
    });

    return obj;
  });

  return {
    rows,
    totalFilteredPlayers,
    visibleCount: rows.length,
    wasCapped,
    maxVisibleRows: MAX_VISIBLE_ROWS,
    attributes: selectedAttributes.map(attr => ({
      key: attr,
      label: ATTRIBUTE_LABELS[attr] || attr
    }))
  };
}

function saveWorkspaceSubmission(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const submissionsSheet = ss.getSheetByName(RATINGS_LOG_SHEET);
  if (!submissionsSheet) throw new Error(`Sheet "${RATINGS_LOG_SHEET}" not found.`);

  const rows = payload.rows || [];
  const originalRows = payload.originalRows || [];
  const attributes = payload.attributes || [];
  const sessionId = payload.sessionId || Utilities.getUuid();

  if (!rows.length || !originalRows.length || !attributes.length) {
    return { savedChanges: 0, message: 'No rows submitted.' };
  }

  const originalMap = {};
  originalRows.forEach(r => {
    originalMap[String(r.name)] = r;
  });

  const submittedAt = new Date();
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  const logHeaders = submissionsSheet.getRange(1, 1, 1, submissionsSheet.getLastColumn()).getValues()[0]
    .map(value => String(value || '').trim());

  const output = [];

  rows.forEach(row => {
    const original = originalMap[String(row.name)];
    if (!original) return;

    attributes.forEach(attr => {
      const attrKey = typeof attr === 'string' ? attr : attr.key;
      const baseline = normaliseNumber(original[attrKey]);
      const suggested = normaliseNumber(row[attrKey]);

      if (baseline === null || suggested === null) return;
      if (baseline === suggested) return;

      output.push(logHeaders.map(header => {
        switch (header) {
          case 'submittedAt': return submittedAt;
          case 'userEmail': return userEmail;
          case 'playerName': return row.name;
          case 'team': return row.team;
          case 'position': return row.position;
          case 'age': return '';
          case 'attribute': return attrKey;
          case 'baselineValue': return baseline;
          case 'suggestedValue': return suggested;
          case 'delta': return suggested - baseline;
          case 'sessionId': return sessionId;
          default: return '';
        }
      }));
    });
  });

  if (!output.length) {
    return { savedChanges: 0, message: 'No changes detected.' };
  }

  submissionsSheet.getRange(
    submissionsSheet.getLastRow() + 1,
    1,
    output.length,
    output[0].length
  ).setValues(output);

  rebuildSummary();
  sendContributionAlert_({
    type: 'ratings',
    userEmail,
    submittedAt,
    sessionId,
    savedChanges: output.length,
    previewLines: output
      .slice(0, CONTRIBUTION_ALERT_PREVIEW_LIMIT)
      .map(row => {
        const rowData = mapLogRow_(logHeaders, row);
        return `${rowData.playerName} (${rowData.team}, ${rowData.position}) - ${rowData.attribute}: ${rowData.baselineValue} -> ${rowData.suggestedValue}`;
      })
  });

  return {
    savedChanges: output.length,
    message: `${output.length} change(s) saved.`,
    sessionId
  };
}

function getAttributeWeightsConfig() {
  const { rows, columns, totals } = getAttributeWeightData_();

  return {
    totalAttributes: rows.length,
    positionGroups: columns.map(column => column.label),
    totals
  };
}

function loadAttributeWeightRows() {
  const { rows, columns, totals } = getAttributeWeightData_();

  return {
    rows,
    columns,
    totalAttributes: rows.length,
    totalPositionGroups: columns.length,
    totals
  };
}

function saveAttributeWeightSubmission(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const submissionsSheet = ensureAttributeWeightSubmissionsSheet_(ss);

  const rows = payload.rows || [];
  const originalRows = payload.originalRows || [];
  const columns = payload.columns || [];
  const sessionId = payload.sessionId || Utilities.getUuid();

  if (!rows.length || !originalRows.length || !columns.length) {
    return { savedChanges: 0, message: 'No weight rows submitted.' };
  }

  const invalidTotals = getInvalidWeightTotals_(rows, columns);
  if (invalidTotals.length) {
    throw new Error(`Attribute weight totals must equal 100 for: ${invalidTotals.join(', ')}`);
  }

  const originalMap = {};
  originalRows.forEach(row => {
    originalMap[String(row.attribute)] = row;
  });

  const submittedAt = new Date();
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';

  const output = [];

  rows.forEach(row => {
    const original = originalMap[String(row.attribute)];
    if (!original) return;

    columns.forEach(column => {
      const baseline = normaliseNumber(original[column.key]);
      const suggested = normaliseNumber(row[column.key]);

      if (baseline === null || suggested === null) return;
      if (baseline === suggested) return;

      output.push([
        submittedAt,
        userEmail,
        row.attribute,
        row.attributeKey || '',
        column.label,
        baseline,
        suggested,
        suggested - baseline,
        sessionId
      ]);
    });
  });

  if (!output.length) {
    return { savedChanges: 0, message: 'No weight changes detected.' };
  }

  submissionsSheet.getRange(
    submissionsSheet.getLastRow() + 1,
    1,
    output.length,
    output[0].length
  ).setValues(output);

  rebuildAttributeWeightSummary();
  sendContributionAlert_({
    type: 'weights',
    userEmail,
    submittedAt,
    sessionId,
    savedChanges: output.length,
    previewLines: output
      .slice(0, CONTRIBUTION_ALERT_PREVIEW_LIMIT)
      .map(row => `${row[2]} (${row[4]}) - ${row[5]} -> ${row[6]}`)
  });

  return {
    savedChanges: output.length,
    message: `${output.length} weight change(s) saved.`,
    sessionId
  };
}

function rebuildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const submissionsSheet = ss.getSheetByName(RATINGS_LOG_SHEET);
  const summarySheet = ensurePlayerSummarySheet_(ss);

  if (!submissionsSheet) throw new Error(`Sheet "${RATINGS_LOG_SHEET}" not found.`);

  const data = submissionsSheet.getDataRange().getValues();

  summarySheet.clear();
  summarySheet.getRange(1, 1, 1, 10).setValues([[
    'playerName',
    'team',
    'position',
    'attribute',
    'submissionCount',
    'baselineValue',
    'avgSuggestedValue',
    'medianSuggestedValue',
    'avgDelta',
    'latestSubmittedAt'
  ]]);

  if (data.length <= 1) return;

  const rows = data.slice(1);
  const headers = data[0].map(value => String(value || '').trim());
  const submittedAtIndex = headers.indexOf('submittedAt');
  const playerNameIndex = headers.indexOf('playerName');
  const teamIndex = headers.indexOf('team');
  const positionIndex = headers.indexOf('position');
  const attributeIndex = headers.indexOf('attribute');
  const baselineValueIndex = headers.indexOf('baselineValue');
  const suggestedValueIndex = headers.indexOf('suggestedValue');
  const deltaIndex = headers.indexOf('delta');
  const grouped = {};

  rows.forEach(r => {
    const submittedAt = r[submittedAtIndex];
    const playerName = r[playerNameIndex];
    const team = r[teamIndex];
    const position = r[positionIndex];
    const attribute = r[attributeIndex];
    const baselineValue = normaliseNumber(r[baselineValueIndex]);
    const suggestedValue = normaliseNumber(r[suggestedValueIndex]);
    const delta = normaliseNumber(r[deltaIndex]);

    if (!playerName || !attribute || suggestedValue === null) return;

    const key = `${playerName}|||${attribute}`;
    if (!grouped[key]) {
      grouped[key] = {
        playerName,
        team,
        position,
        attribute,
        baselineValue,
        suggestedValues: [],
        deltas: [],
        latestSubmittedAt: submittedAt
      };
    }

    grouped[key].suggestedValues.push(suggestedValue);
    if (delta !== null) grouped[key].deltas.push(delta);

    if (submittedAt && submittedAt > grouped[key].latestSubmittedAt) {
      grouped[key].latestSubmittedAt = submittedAt;
    }
  });

  const output = Object.values(grouped)
    .map(g => [
      g.playerName,
      g.team,
      g.position,
      g.attribute,
      g.suggestedValues.length,
      g.baselineValue,
      average(g.suggestedValues),
      median(g.suggestedValues),
      average(g.deltas),
      g.latestSubmittedAt
    ])
    .sort((a, b) => {
      if (a[0] === b[0]) return String(a[3]).localeCompare(String(b[3]));
      return String(a[0]).localeCompare(String(b[0]));
    });

  if (output.length) {
    summarySheet.getRange(2, 1, output.length, output[0].length).setValues(output);
  }
}

function rebuildAttributeWeightSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const submissionsSheet = ensureAttributeWeightSubmissionsSheet_(ss);
  const summarySheet = ensureAttributeWeightSummarySheet_(ss);

  const data = submissionsSheet.getDataRange().getValues();

  summarySheet.clear();
  summarySheet.getRange(1, 1, 1, 9).setValues([[
    'attribute',
    'attributeKey',
    'positionGroup',
    'submissionCount',
    'baselineValue',
    'avgSuggestedValue',
    'medianSuggestedValue',
    'avgDelta',
    'latestSubmittedAt'
  ]]);

  if (data.length <= 1) return;

  const rows = data.slice(1);
  const grouped = {};

  rows.forEach(r => {
    const submittedAt = r[0];
    const attribute = r[2];
    const attributeKey = r[3];
    const positionGroup = r[4];
    const baselineValue = normaliseNumber(r[5]);
    const suggestedValue = normaliseNumber(r[6]);
    const delta = normaliseNumber(r[7]);

    if (!attribute || !positionGroup || suggestedValue === null) return;

    const key = `${attribute}|||${positionGroup}`;
    if (!grouped[key]) {
      grouped[key] = {
        attribute,
        attributeKey,
        positionGroup,
        baselineValue,
        suggestedValues: [],
        deltas: [],
        latestSubmittedAt: submittedAt
      };
    }

    grouped[key].suggestedValues.push(suggestedValue);
    if (delta !== null) grouped[key].deltas.push(delta);

    if (submittedAt && submittedAt > grouped[key].latestSubmittedAt) {
      grouped[key].latestSubmittedAt = submittedAt;
    }
  });

  const output = Object.values(grouped)
    .map(g => [
      g.attribute,
      g.attributeKey,
      g.positionGroup,
      g.suggestedValues.length,
      g.baselineValue,
      average(g.suggestedValues),
      median(g.suggestedValues),
      average(g.deltas),
      g.latestSubmittedAt
    ])
    .sort((a, b) => {
      if (a[0] === b[0]) return String(a[2]).localeCompare(String(b[2]));
      return String(a[0]).localeCompare(String(b[0]));
    });

  if (output.length) {
    summarySheet.getRange(2, 1, output.length, output[0].length).setValues(output);
  }
}

function normaliseNumber(value) {
  if (value === '' || value === null || typeof value === 'undefined') return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function average(arr) {
  if (!arr || !arr.length) return null;
  return arr.reduce((sum, x) => sum + x, 0) / arr.length;
}

function median(arr) {
  if (!arr || !arr.length) return null;
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function sendContributionAlert_(payload) {
  if (!CONTRIBUTION_ALERT_RECIPIENTS.length) return;

  const subjectPrefix = payload.type === 'weights' ? 'Weight contribution' : 'Rating contribution';
  const previewLines = (payload.previewLines || []).filter(Boolean);
  const extraChanges = Math.max(0, Number(payload.savedChanges || 0) - previewLines.length);
  const emailBody = [
    `${subjectPrefix} received for Forge Rugby.`,
    '',
    `Contributor: ${payload.userEmail || 'unknown'}`,
    `Saved changes: ${payload.savedChanges || 0}`,
    `Submitted at: ${formatAlertDate_(payload.submittedAt)}`,
    `Session ID: ${payload.sessionId || ''}`,
    '',
    'Changes:',
    previewLines.length ? previewLines.map(line => `- ${line}`).join('\n') : '- No preview available',
    extraChanges ? `- ...and ${extraChanges} more change(s)` : '',
    '',
    `Spreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`
  ].filter(Boolean).join('\n');

  try {
    MailApp.sendEmail({
      to: CONTRIBUTION_ALERT_RECIPIENTS.join(','),
      subject: `[Forge Rugby] ${subjectPrefix} by ${payload.userEmail || 'unknown'}`,
      body: emailBody
    });
  } catch (error) {
    console.error(`Failed to send contribution alert: ${error && error.message ? error.message : error}`);
  }
}

function mapLogRow_(headers, values) {
  return headers.reduce((acc, header, index) => {
    acc[header] = values[index];
    return acc;
  }, {});
}

function formatAlertDate_(value) {
  if (!(value instanceof Date)) return String(value || '');
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function getAttributeWeightData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(WEIGHTS_SHEET);
  if (!sheet) throw new Error(`Sheet "${WEIGHTS_SHEET}" not found.`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 2) {
    throw new Error(`Sheet "${WEIGHTS_SHEET}" has no weight data.`);
  }

  const headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(value => String(value || '').trim());

  const columns = headerValues.slice(1)
    .map((label, index) => ({
      key: `position_${index}`,
      label,
      columnIndex: index + 2
    }))
    .filter(column => column.label !== '');

  const dataRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = [];
  const totals = {};

  dataRows.forEach(values => {
    const attributeLabel = String(values[0] || '').trim();
    if (!attributeLabel) return;

    if (attributeLabel.toLowerCase() === 'total') {
      columns.forEach(column => {
        totals[column.label] = normaliseNumber(values[column.columnIndex - 1]);
      });
      return;
    }

    const row = {
      attribute: attributeLabel,
      attributeKey: ATTRIBUTE_LABEL_TO_KEY[attributeLabel] || ''
    };

    columns.forEach(column => {
      row[column.key] = values[column.columnIndex - 1];
    });

    rows.push(row);
  });

  return {
    rows,
    columns,
    totals
  };
}

function ensureAttributeWeightSubmissionsSheet_(ss) {
  let sheet = ss.getSheetByName(WEIGHTS_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(WEIGHTS_LOG_SHEET);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 10).setValues([[
      'submittedAt',
      'userEmail',
      'attribute',
      'attributeKey',
      'positionGroup',
      'baselineValue',
      'suggestedValue',
      'delta',
      'sessionId'
    ]]);
  }

  return sheet;
}

function ensurePlayerSummarySheet_(ss) {
  let sheet = ss.getSheetByName(RATINGS_ROLLUP_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RATINGS_ROLLUP_SHEET);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 9).setValues([[
      'playerName',
      'team',
      'position',
      'attribute',
      'submissionCount',
      'baselineValue',
      'avgSuggestedValue',
      'medianSuggestedValue',
      'avgDelta',
      'latestSubmittedAt'
    ]]);
  }

  return sheet;
}

function ensureAttributeWeightSummarySheet_(ss) {
  let sheet = ss.getSheetByName(WEIGHTS_ROLLUP_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(WEIGHTS_ROLLUP_SHEET);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 9).setValues([[
      'attribute',
      'attributeKey',
      'positionGroup',
      'submissionCount',
      'baselineValue',
      'avgSuggestedValue',
      'medianSuggestedValue',
      'avgDelta',
      'latestSubmittedAt'
    ]]);
  }

  return sheet;
}

function getInvalidWeightTotals_(rows, columns) {
  return columns
    .map(column => {
      const total = rows.reduce((sum, row) => sum + (normaliseNumber(row[column.key]) || 0), 0);
      return total === 100 ? null : `${column.label} (${total})`;
    })
    .filter(Boolean);
}
