/**
 * ═══════════════════════════════════════════════════════════════
 *  TutorMetrics + TutorTrack + Tutor — Combined Google Apps Script
 * ═══════════════════════════════════════════════════════════════
 *
 *  CARA KERJA:
 *  - Deploy sekali sebagai Web App
 *  - Dashboard Package  panggil:  https://.../exec?source=metrics
 *  - Dashboard Timesheet panggil: https://.../exec?source=tracker
 *  - Dashboard Tutor     panggil: https://.../exec?source=tutor
 *
 *  KONFIGURASI:
 *  1. Ganti SPREADSHEET_ID dengan ID spreadsheet kamu
 *  2. Ganti nama sheet di objek SHEETS di bawah
 *  3. Deploy → Web App → Anyone → Deploy
 *
 * ═══════════════════════════════════════════════════════════════
 */

const SPREADSHEET_ID = '1BCbh-J7Rwxv3bbVtUnMATnxvqwDR8tF5TEb2zaW8hWw';

/**
 * Daftar sheet dan kolom tanggal untuk masing-masing source.
 * Ganti nama sheet sesuai tab di spreadsheet kamu.
 */
const SHEETS = {
  metrics: {
    name: 'all_tutoring_crm',
    dateField: 'success_date',
  },
  tracker: {
    name: 'all_tutoring_timesheet',
    dateField: 'session_date',
  },
  tutor: {
    name: 'tutor_update',
    dateField: 'updated_at',
  },
  session_qc: {
    name: 'session_qc',
    dateField: 'updated_at',
  },
  tutor_qc: {
    name: 'tutor_qc',
    dateField: 'updated_at',
  },
};


function doGet(e) {
  try {
    var source = (e && e.parameter && e.parameter.source) || '';

    if (!SHEETS[source]) {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var allSheets = ss.getSheets().map(function(s) { return s.getName().toLowerCase(); });

      if (source === '') {
        // Fallback: tebak dari nama sheet yang ada
        if (allSheets.indexOf('all_tutoring_crm') !== -1) source = 'metrics';
        else if (allSheets.indexOf('all_tutoring_timesheet') !== -1) source = 'tracker';
        else if (allSheets.indexOf('all_tutoring_tutor') !== -1) source = 'tutor';
        else if (allSheets.indexOf('crm') !== -1) source = 'metrics';
        else if (allSheets.indexOf('timesheet') !== -1) source = 'tracker';
        else if (allSheets.indexOf('tutor') !== -1) source = 'tutor';
        else if (allSheets.indexOf('session_qc') !== -1) source = 'session_qc';
        else if (allSheets.indexOf('tutor_qc') !== -1) source = 'tutor_qc';
        else source = 'metrics';
      } else {
        return jsonResponse(
          { error: 'Unknown source "' + source + '". Use: metrics, tracker, tutor or session_qc.' },
          400
        );
      }
    }

    var config = SHEETS[source];
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(config.name);

    if (!sheet) {
      return jsonResponse(
        { error: 'Sheet "' + config.name + '" not found. Available: ' + ss.getSheets().map(function(s){return s.getName()}).join(', ') },
        404
      );
    }

    var data = sheet.getDataRange().getValues();

    if (!data || data.length <= 1) {
      return jsonResponse([]);
    }

    // Baris pertama = header, normalize ke snake_case
    var headers = data[0].map(function(h) {
      return String(h)
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '_')
        .replace(/^_+|_+$/g, '');
    });

    // Mapping setiap baris menjadi object
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = (data[i][j] !== undefined && data[i][j] !== null) ? data[i][j] : null;
      }
      rows.push(row);
    }

    var response = {
      source: source,
      filter_date_field: config.dateField,
      data: rows,
    };

    return jsonResponse(response);

  } catch (err) {
    return jsonResponse({ error: err.toString() }, 500);
  }
}


/**
 * Helper: kirim respons JSON
 */
function jsonResponse(data, statusCode) {
  statusCode = statusCode || 200;
  var payload = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return payload;
}


/* ══════════════════════════════════════════════════════════════
   HELPER: Custom menu + test functions
   ══════════════════════════════════════════════════════════════ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TutorTools')
    .addItem('Test Metrics endpoint', 'testMetrics')
    .addItem('Test Tracker endpoint', 'testTracker')
    .addItem('Test Tutor endpoint', 'testTutor')
    .addItem('Test Session QC endpoint', 'testSessionQC')
    .addItem('Test Tutor QC endpoint', 'testTutorQC')
    .addItem('List all sheets', 'listSheets')
    .addToUi();
}

function testMetrics() {
  var result = doGet({ parameter: { source: 'metrics' } });
  var parsed = JSON.parse(result.getContent());
  Logger.log('METRICS: ' + parsed.data.length + ' rows');
  if (parsed.error) Logger.log('ERROR: ' + parsed.error);
  else if (parsed.data.length > 0) Logger.log('Sample: ' + JSON.stringify(parsed.data[0]).substring(0, 200));
  SpreadsheetApp.getUi().alert(
    'Metrics OK',
    parsed.error || (parsed.data.length + ' rows. Cek log untuk detail.'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function testTracker() {
  var result = doGet({ parameter: { source: 'tracker' } });
  var parsed = JSON.parse(result.getContent());
  Logger.log('TRACKER: ' + parsed.data.length + ' rows');
  if (parsed.error) Logger.log('ERROR: ' + parsed.error);
  else if (parsed.data.length > 0) Logger.log('Sample: ' + JSON.stringify(parsed.data[0]).substring(0, 200));
  SpreadsheetApp.getUi().alert(
    'Tracker OK',
    parsed.error || (parsed.data.length + ' rows. Cek log untuk detail.'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function testTutor() {
  var result = doGet({ parameter: { source: 'tutor' } });
  var parsed = JSON.parse(result.getContent());
  Logger.log('TUTOR: ' + parsed.data.length + ' rows');
  if (parsed.error) Logger.log('ERROR: ' + parsed.error);
  else if (parsed.data.length > 0) Logger.log('Sample: ' + JSON.stringify(parsed.data[0]).substring(0, 300));
  SpreadsheetApp.getUi().alert(
    'Tutor OK',
    parsed.error || (parsed.data.length + ' rows. Cek log untuk detail.'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function testSessionQC() {
  var result = doGet({ parameter: { source: 'session_qc' } });
  var parsed = JSON.parse(result.getContent());

  Logger.log('SESSION_QC: ' + parsed.data.length + ' rows');

  if (parsed.error) {
    Logger.log('ERROR: ' + parsed.error);
  } else if (parsed.data.length > 0) {
    Logger.log('Sample: ' + JSON.stringify(parsed.data[0]).substring(0, 300));
  }

  SpreadsheetApp.getUi().alert(
    'Session QC OK',
    parsed.error || (parsed.data.length + ' rows. Cek log untuk detail.'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function testTutorQC() {
  var result = doGet({ parameter: { source: 'tutor_qc' } });
  var parsed = JSON.parse(result.getContent());

  Logger.log('TUTOR_QC: ' + parsed.data.length + ' rows');

  if (parsed.error) {
    Logger.log('ERROR: ' + parsed.error);
  } else if (parsed.data.length > 0) {
    Logger.log('Sample: ' + JSON.stringify(parsed.data[0]).substring(0, 300));
  }

  SpreadsheetApp.getUi().alert(
    'Tutor QC OK',
    parsed.error || (parsed.data.length + ' rows. Cek log untuk detail.'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function listSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var names = ss.getSheets().map(function(s) { return s.getName(); });
  Logger.log('Sheets: ' + names.join(', '));
  SpreadsheetApp.getUi().alert('Sheet names:\n' + names.join('\n'));
}
