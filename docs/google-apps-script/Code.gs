const HEADERS = Object.freeze({
  timestamp: ['Zeitstempel', 'Timestamp'],
  roundId: ['Beurteilungsrunde', 'Runden-ID'],
  name: ['Vorname, Nachname', 'Vorname/Nachname', 'Vorname und Nachname', 'Name'],
  email: ['E-Mail Adresse', 'E-Mail-Adresse', 'Email Address'],
  additionalQuality: [
    'Zusätzliche Qualität meiner Äusserungen',
    'Ich sehe in meinen Äusserungen zusätzlich folgende Qualität, die oben nicht erfragt worden ist:'
  ],
  comment: [
    'Sonstige Bemerkungen',
    'Sonstige Bemerkungen (z.B. zum Klima/ Gruppendynamik im PH-Mündlichunterricht):'
  ],
  criteria: Object.freeze({
    participation_frequency: [
      'Beteiligung mit Äusserungen',
      'In den PHI-Lektionen beteiligte ich mich mit Äusserungen ...'
    ],
    insufficient_contributions: [
      'Falsch, unbefriedigend oder nicht ausreichend',
      'Falsch, unbefriedigend oder nicht ausreichend (als Antwort auf Aufträge/ gestellte Fragen)?'
    ],
    original_unsuitable_contributions: [
      'Originell, aber unpassend im Lektionsverlauf',
      'Originell, aber unpassend im Lektionsverlauf?'
    ],
    correct_brief_contributions: ['Korrekt, aber stichwortartig', 'Korrekt, aber stichwortartig?'],
    advancing_contributions: [
      'Passend und im Lektionsverlauf weiterführend',
      'Passend und im Lektionsverlauf weiterführend?'
    ],
    independent_contributions: [
      'Passend, eigenständig und in eine neue, interessante Richtung führend',
      'Passend, eigenständig und in eine neue, interessante Richtung führend?'
    ],
    overall_participation: [
      'Teilnahme insgesamt',
      'Meine Teilnahme im letzten Semester im Fach PH schätze ich insgesamt ein als ...'
    ]
  })
});

function doGet(event) {
  try {
    const parameters = event && event.parameter || {};
    const properties = PropertiesService.getScriptProperties();
    if (!parameters.apiKey || parameters.apiKey !== properties.getProperty('ASSESSMENT_API_KEY')) {
      return json({ success: false, error: 'unauthorized' });
    }
    if (!parameters.roundId) return json({ success: false, error: 'missing_round_id' });
    const responses = readResponses(parameters.roundId, properties);
    return json({ success: true, roundId: parameters.roundId, count: responses.length, responses: responses });
  } catch (error) {
    return json({ success: false, error: 'internal_error' });
  }
}

function readResponses(roundId, properties) {
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  const spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Spreadsheet not found');
  const configuredSheet = properties.getProperty('RESPONSES_SHEET_NAME');
  const sheet = configuredSheet ? spreadsheet.getSheetByName(configuredSheet) : spreadsheet.getSheets()[0];
  if (!sheet) throw new Error('Response sheet not found');
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const columns = mapColumns(values[0]);
  return values.slice(1).map(function(row) {
    return mapResponse(row, columns);
  }).filter(function(response) {
    return response.roundId === roundId;
  });
}

function mapColumns(headerRow) {
  const columns = {
    timestamp: findColumn(headerRow, HEADERS.timestamp),
    roundId: findColumn(headerRow, HEADERS.roundId),
    name: findColumn(headerRow, HEADERS.name),
    email: findColumn(headerRow, HEADERS.email),
    additionalQuality: findColumn(headerRow, HEADERS.additionalQuality),
    comment: findColumn(headerRow, HEADERS.comment),
    criteria: {}
  };
  Object.keys(HEADERS.criteria).forEach(function(id) {
    columns.criteria[id] = findColumn(headerRow, HEADERS.criteria[id]);
  });
  return columns;
}

function mapResponse(row, columns) {
  const answers = {};
  Object.keys(columns.criteria).forEach(function(id) {
    answers[id] = parseRating(valueAt(row, columns.criteria[id]));
  });
  const timestamp = valueAt(row, columns.timestamp);
  const roundId = valueAt(row, columns.roundId);
  const name = valueAt(row, columns.name);
  const email = valueAt(row, columns.email).toLowerCase();
  return {
    responseId: stableId([timestamp, roundId, email, name].join('\n')),
    timestamp: new Date(timestamp).toISOString(),
    roundId: roundId,
    name: name,
    email: email,
    answers: answers,
    additionalQuality: valueAt(row, columns.additionalQuality),
    comment: valueAt(row, columns.comment)
  };
}

function findColumn(headers, acceptedNames) {
  const normalized = headers.map(normalize);
  const index = acceptedNames.map(normalize).map(function(name) { return normalized.indexOf(name); }).find(function(value) { return value >= 0; });
  if (index === undefined) throw new Error('Required column missing');
  return index;
}

function valueAt(row, index) {
  return String(row[index] === undefined || row[index] === null ? '' : row[index]).trim();
}

function parseRating(value) {
  const text = String(value || '').trim();
  const numeric = text.match(/^([1-5])(?:\D|$)/);
  if (numeric) return Number(numeric[1]);
  const labels = {
    '(fast) nie': 1,
    'fast nie': 1,
    'ab und zu': 2,
    'manchmal': 3,
    'häufig': 4,
    'sehr häufig': 5,
    'mangelhaft': 1,
    'genügend': 2,
    'recht': 3,
    'gut': 4,
    'sehr gut': 5
  };
  return labels[text.toLowerCase()] || 0;
}

function normalize(value) {
  return String(value || '').trim().toLowerCase();
}

function stableId(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8).map(function(byte) {
    return ('0' + ((byte + 256) % 256).toString(16)).slice(-2);
  }).join('');
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
