const APP_COLUMNS = [
  'created_at_utc',
  'source',
  'company',
  'role',
  'stage',
  'status',
  'applied_date',
  'external_id',
  'email_subject',
  'email_from',
  'thread_url',
  'notes'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Job Funnel')
    .addItem('Setup Sheets', 'setupSheets')
    .addItem('Sync Job Emails', 'syncJobEmails')
    .addItem('Rebuild Metrics', 'rebuildMetrics')
    .addItem('Build Dashboard', 'buildDashboard')
    .addItem('Build Follow-Up Queue', 'buildDefaultFollowUpQueue')
    .addToUi();
}

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: props.getProperty('SPREADSHEET_ID') || '',
    sourceLabel: props.getProperty('SOURCE_LABEL') || 'jobs/applications/inbox',
    processedLabel: props.getProperty('PROCESSED_LABEL') || 'jobs/applications/processed',
    searchLimit: Number(props.getProperty('SEARCH_LIMIT') || 150),
    ghostedDays: Number(props.getProperty('GHOSTED_DAYS') || 45)
  };
}

function getSpreadsheet_() {
  const cfg = getConfig();
  if (cfg.spreadsheetId) {
    return SpreadsheetApp.openById(cfg.spreadsheetId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function setupSheets() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, 'Applications', APP_COLUMNS);
  ensureSheet_(ss, 'Metrics', ['metric', 'value']);
  ensureSheet_(ss, 'MetricsByYear', ['year', 'metric', 'value']);
  ensureSheet_(ss, 'FollowUpQueue', ['company', 'role', 'source', 'applied_date', 'days_since_apply', 'thread_url']);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function syncJobEmails() {
  const cfg = getConfig();
  const ss = getSpreadsheet_();
  setupSheets();

  const appSheet = ss.getSheetByName('Applications');
  const existingIds = getExistingExternalIds_(appSheet);

  ensureLabel_(cfg.sourceLabel);
  ensureLabel_(cfg.processedLabel);

  const query = `label:"${cfg.sourceLabel}" -label:"${cfg.processedLabel}"`;
  const threads = GmailApp.search(query, 0, cfg.searchLimit);

  const processedLabel = GmailApp.getUserLabelByName(cfg.processedLabel);
  const rowsToAppend = [];

  threads.forEach((thread) => {
    const messages = thread.getMessages();
    messages.forEach((message) => {
      const parsed = parseJobEmail_(message, thread);
      if (!parsed.externalId || existingIds.has(parsed.externalId)) {
        return;
      }
      rowsToAppend.push([
        isoUtc_(new Date()),
        parsed.source,
        parsed.company,
        parsed.role,
        parsed.stage,
        parsed.status,
        parsed.appliedDate,
        parsed.externalId,
        parsed.subject,
        parsed.from,
        parsed.threadUrl,
        ''
      ]);
      existingIds.add(parsed.externalId);
    });

    // Some threads (e.g., restricted/system conversations) may reject label operations.
    // Continue sync instead of failing the full run.
    try {
      thread.addLabel(processedLabel);
    } catch (error) {
      Logger.log(`Could not add processed label to thread ${thread.getId()}: ${error}`);
    }
  });

  if (rowsToAppend.length > 0) {
    const start = appSheet.getLastRow() + 1;
    appSheet
      .getRange(start, 1, rowsToAppend.length, APP_COLUMNS.length)
      .setValues(rowsToAppend);
  }

  rebuildMetrics();
}

function parseJobEmail_(message, thread) {
  const subject = sanitize_(message.getSubject());
  const from = sanitize_(message.getFrom());
  const body = sanitize_(message.getPlainBody() || '');
  const lower = `${subject}\n${body}`.toLowerCase();

  const stage = inferStage_(lower);
  const company = inferCompany_(subject, from, body);
  const role = inferRole_(subject, body);
  const appliedDate = formatDate_(message.getDate());

  let externalId = sanitize_(message.getHeader('Message-ID'));
  if (!externalId) {
    externalId = `${thread.getId()}_${message.getId()}`;
  }

  return {
    source: 'Gmail',
    company: company || 'Unknown',
    role: role || 'Unknown',
    stage,
    status: stage,
    appliedDate,
    externalId,
    subject,
    from,
    threadUrl: `https://mail.google.com/mail/u/0/#inbox/${thread.getId()}`
  };
}

function inferStage_(lowerText) {
  if (/(interview|availability|schedule call|phone screen|technical screen)/.test(lowerText)) {
    return 'Interview';
  }
  if (/(assessment|codesignal|hackerrank|take-home|take home|oa)/.test(lowerText)) {
    return 'Assessment';
  }
  if (/(offer|congratulations|we are excited to offer)/.test(lowerText)) {
    return 'Offer';
  }
  if (/(rejected|unfortunately|not moving forward|we have decided to proceed with other candidates)/.test(lowerText)) {
    return 'Rejected';
  }
  if (/(received your application|thanks for applying|application submitted|application received)/.test(lowerText)) {
    return 'Applied';
  }
  return 'Updated';
}

function inferCompany_(subject, from, body) {
  const subjectMatch = subject.match(/(?:at|with)\s+([A-Z][A-Za-z0-9&.\- ]{2,})/);
  if (subjectMatch) {
    return cleanEntity_(subjectMatch[1]);
  }

  const fromMatch = from.match(/<[^@]+@([A-Za-z0-9.-]+)>/);
  if (fromMatch) {
    const host = fromMatch[1].split('.').slice(-2, -1)[0];
    if (host) {
      return toTitleCase_(host.replace(/[-_]/g, ' '));
    }
  }

  const bodyMatch = body.match(/(?:company|employer)\s*:\s*([A-Za-z0-9&.\- ]{2,})/i);
  if (bodyMatch) {
    return cleanEntity_(bodyMatch[1]);
  }

  return '';
}

function inferRole_(subject, body) {
  const rolePatterns = [
    /(?:for|as)\s+(?:the\s+)?([A-Z][A-Za-z0-9\-\/+ ]{3,})\s+(?:role|position)/,
    /position\s*:\s*([A-Za-z0-9\-\/+ ]{3,})/i,
    /role\s*:\s*([A-Za-z0-9\-\/+ ]{3,})/i
  ];

  for (let i = 0; i < rolePatterns.length; i += 1) {
    const fromSubject = subject.match(rolePatterns[i]);
    if (fromSubject) {
      return cleanEntity_(fromSubject[1]);
    }
    const fromBody = body.match(rolePatterns[i]);
    if (fromBody) {
      return cleanEntity_(fromBody[1]);
    }
  }

  return '';
}

function cleanEntity_(value) {
  return sanitize_(value).replace(/[|,;]+$/, '').trim();
}

function sanitize_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function toTitleCase_(value) {
  return value
    .split(' ')
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(' ');
}

function isoUtc_(date) {
  return Utilities.formatDate(date, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function ensureLabel_(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

function getExistingExternalIds_(appSheet) {
  const last = appSheet.getLastRow();
  const set = new Set();

  if (last <= 1) {
    return set;
  }

  const values = appSheet.getRange(2, 8, last - 1, 1).getValues();
  values.forEach((row) => {
    const key = sanitize_(row[0]);
    if (key) {
      set.add(key);
    }
  });

  return set;
}

function importSimplifyCsvFromDrive(fileId) {
  const file = DriveApp.getFileById(fileId);
  const csvText = file.getBlob().getDataAsString();
  importSimplifyCsv_(csvText, file.getName());
}

function importSimplifyCsv_(csvText, sourceName) {
  const ss = getSpreadsheet_();
  setupSheets();

  const appSheet = ss.getSheetByName('Applications');
  const existingIds = getExistingExternalIds_(appSheet);

  const rows = Utilities.parseCsv(csvText);
  if (rows.length < 2) {
    return;
  }

  const header = rows[0].map((h) => sanitize_(h).toLowerCase());
  const idx = {
    company: indexOfAny_(header, ['company', 'company name']),
    role: indexOfAny_(header, ['role', 'job title', 'title', 'position']),
    status: indexOfAny_(header, ['status', 'application status']),
    appliedDate: indexOfAny_(header, ['date applied', 'applied date', 'application date']),
    url: indexOfAny_(header, ['url', 'job url', 'posting url'])
  };

  const toAppend = [];
  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const company = fromIndex_(row, idx.company) || 'Unknown';
    const role = fromIndex_(row, idx.role) || 'Unknown';
    const status = fromIndex_(row, idx.status) || 'Applied';
    const appliedDate = normalizeDate_(fromIndex_(row, idx.appliedDate));
    const url = fromIndex_(row, idx.url);
    const ext = `simplify_${company}_${role}_${appliedDate}`.toLowerCase().replace(/\s+/g, '_');

    if (existingIds.has(ext)) {
      continue;
    }

    toAppend.push([
      isoUtc_(new Date()),
      sourceName || 'Simplify',
      company,
      role,
      status,
      status,
      appliedDate,
      ext,
      '',
      '',
      url,
      ''
    ]);

    existingIds.add(ext);
  }

  if (toAppend.length > 0) {
    const start = appSheet.getLastRow() + 1;
    appSheet.getRange(start, 1, toAppend.length, APP_COLUMNS.length).setValues(toAppend);
  }

  rebuildMetrics();
}

function indexOfAny_(header, names) {
  for (let i = 0; i < names.length; i += 1) {
    const idx = header.indexOf(names[i]);
    if (idx >= 0) {
      return idx;
    }
  }
  return -1;
}

function fromIndex_(row, idx) {
  if (idx < 0 || idx >= row.length) {
    return '';
  }
  return sanitize_(row[idx]);
}

function normalizeDate_(value) {
  if (!value) {
    return '';
  }

  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return sanitize_(value);
}

function normalizeMetricToken_(value) {
  const token = sanitize_(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  return token || 'unknown';
}

function extractYear_(value) {
  const text = sanitize_(value);
  if (!text) {
    return '';
  }

  const yearMatch = text.match(/\b(19\d{2}|20\d{2})\b/);
  if (yearMatch) {
    return yearMatch[1];
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return String(parsed.getFullYear());
  }

  return '';
}

function inferYearBucket_(appliedDate, createdAtUtc) {
  return extractYear_(appliedDate) || extractYear_(createdAtUtc) || 'unknown';
}

function extractMonth_(value) {
  const text = sanitize_(value);
  if (!text) {
    return '';
  }

  const monthMatch = text.match(/\b(19\d{2}|20\d{2})-(0[1-9]|1[0-2])\b/);
  if (monthMatch) {
    return `${monthMatch[1]}-${monthMatch[2]}`;
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM');
  }

  return '';
}

function inferMonthBucket_(appliedDate, createdAtUtc) {
  return extractMonth_(appliedDate) || extractMonth_(createdAtUtc) || 'unknown';
}

function incrementMap_(map, key, amount) {
  map.set(key, (map.get(key) || 0) + amount);
}

function incrementYearMetric_(bucketMap, year, metric, amount) {
  if (!bucketMap.has(year)) {
    bucketMap.set(year, new Map());
  }
  const yearMap = bucketMap.get(year);
  yearMap.set(metric, (yearMap.get(metric) || 0) + amount);
}

function parseDateSafe_(value) {
  const text = sanitize_(value);
  if (!text) {
    return null;
  }
  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function resolveStageForMetrics_(rawStage, appliedDate, createdAtUtc, now, ghostedDays) {
  const normalized = normalizeMetricToken_(rawStage);
  if (normalized !== 'updated') {
    return normalized;
  }

  const referenceDate = parseDateSafe_(appliedDate) || parseDateSafe_(createdAtUtc);
  if (!referenceDate) {
    return 'applied';
  }

  const ageDays = Math.floor((now.getTime() - referenceDate.getTime()) / (1000 * 60 * 60 * 24));
  return ageDays >= ghostedDays ? 'ghosted' : 'applied';
}

function toTitleLabel_(token) {
  return sanitize_(token)
    .split('_')
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(' ') || 'Unknown';
}

function sortMapByValueDesc_(map) {
  return [...map.entries()].sort((a, b) => b[1] - a[1]);
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function rebuildMetrics() {
  const cfg = getConfig();
  const ss = getSpreadsheet_();
  setupSheets();

  const appSheet = ss.getSheetByName('Applications');
  const metricsSheet = ss.getSheetByName('Metrics');
  const byYearSheet = ss.getSheetByName('MetricsByYear');

  const last = appSheet.getLastRow();
  const metrics = new Map();
  const yearBuckets = new Map();
  const now = new Date();
  metrics.set('total_rows', Math.max(0, last - 1));
  metrics.set('ghosted_days_threshold', cfg.ghostedDays);

  if (last > 1) {
    const rows = appSheet.getRange(2, 1, last - 1, APP_COLUMNS.length).getValues();

    rows.forEach((row) => {
      const createdAtUtc = sanitize_(row[0]);
      const source = normalizeMetricToken_(row[1]);
      const company = sanitize_(row[2]);
      const role = sanitize_(row[3]);
      const rawStage = normalizeMetricToken_(row[4]);
      const appliedDate = sanitize_(row[6]);
      const year = inferYearBucket_(appliedDate, createdAtUtc);
      const stage = resolveStageForMetrics_(rawStage, appliedDate, createdAtUtc, now, cfg.ghostedDays);

      if (rawStage === 'updated') {
        incrementMap_(metrics, 'stage_updated_raw', 1);
        incrementYearMetric_(yearBuckets, year, 'stage_updated_raw', 1);
      }

      incrementMap_(metrics, `stage_${stage}`, 1);
      incrementMap_(metrics, `source_${source}`, 1);
      incrementYearMetric_(yearBuckets, year, 'total_rows', 1);
      incrementYearMetric_(yearBuckets, year, `stage_${stage}`, 1);
      incrementYearMetric_(yearBuckets, year, `source_${source}`, 1);

      if (!company || company.toLowerCase() === 'unknown') {
        incrementMap_(metrics, 'unknown_company_rows', 1);
        incrementYearMetric_(yearBuckets, year, 'unknown_company_rows', 1);
      }

      if (!role || role.toLowerCase() === 'unknown') {
        incrementMap_(metrics, 'unknown_role_rows', 1);
        incrementYearMetric_(yearBuckets, year, 'unknown_role_rows', 1);
      }
    });
  }

  const years = [...yearBuckets.keys()].sort();
  years.forEach((year) => {
    const values = yearBuckets.get(year);
    values.forEach((value, metric) => {
      metrics.set(`year_${year}_${metric}`, value);
    });
  });

  const byYearOutput = [['year', 'metric', 'value']];
  years.forEach((year) => {
    const values = yearBuckets.get(year);
    [...values.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .forEach((entry) => byYearOutput.push([year, entry[0], entry[1]]));
  });

  const output = [['metric', 'value']];
  [...metrics.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach((entry) => output.push([entry[0], entry[1]]));

  metricsSheet.clear();
  metricsSheet.getRange(1, 1, output.length, 2).setValues(output);
  metricsSheet.setFrozenRows(1);

  byYearSheet.clear();
  byYearSheet.getRange(1, 1, byYearOutput.length, 3).setValues(byYearOutput);
  byYearSheet.setFrozenRows(1);

  buildDashboard();
}

function buildDashboard() {
  const cfg = getConfig();
  const ss = getSpreadsheet_();
  setupSheets();

  const appSheet = ss.getSheetByName('Applications');
  const dashboardSheet = getOrCreateSheet_(ss, 'Dashboard');
  const dataSheet = getOrCreateSheet_(ss, 'DashboardData');

  const last = appSheet.getLastRow();
  dataSheet.clear();
  dashboardSheet.clear();
  dashboardSheet.getCharts().forEach((chart) => dashboardSheet.removeChart(chart));

  dashboardSheet.getRange('A1').setValue('Job Funnel Dashboard').setFontWeight('bold').setFontSize(16);
  dashboardSheet.getRange('A2').setValue(`Last rebuilt: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`);

  if (last <= 1) {
    dashboardSheet.getRange('A4').setValue('No application rows found. Run syncJobEmails first.');
    dataSheet.getRange(1, 1).setValue('No data');
    dataSheet.hideSheet();
    return;
  }

  const rows = appSheet.getRange(2, 1, last - 1, APP_COLUMNS.length).getValues();
  const stageTotals = new Map();
  const sourceTotals = new Map();
  const monthlyTotals = new Map();
  const yearStage = new Map();
  const now = new Date();

  rows.forEach((row) => {
    const createdAtUtc = sanitize_(row[0]);
    const source = normalizeMetricToken_(row[1]);
    const rawStage = normalizeMetricToken_(row[4]);
    const appliedDate = sanitize_(row[6]);
    const year = inferYearBucket_(appliedDate, createdAtUtc);
    const month = inferMonthBucket_(appliedDate, createdAtUtc);
    const stage = resolveStageForMetrics_(rawStage, appliedDate, createdAtUtc, now, cfg.ghostedDays);

    incrementMap_(stageTotals, stage, 1);
    incrementMap_(sourceTotals, source, 1);
    incrementMap_(monthlyTotals, month, 1);
    incrementYearMetric_(yearStage, year, stage, 1);
  });

  const stageRows = [['stage', 'count']];
  sortMapByValueDesc_(stageTotals).forEach((entry) => stageRows.push([toTitleLabel_(entry[0]), entry[1]]));

  const sourceRows = [['source', 'count']];
  sortMapByValueDesc_(sourceTotals).forEach((entry) => sourceRows.push([toTitleLabel_(entry[0]), entry[1]]));

  const stages = [...stageTotals.keys()].sort();
  const years = [...yearStage.keys()].sort();
  const yearStageRows = [['year', ...stages.map((stage) => toTitleLabel_(stage))]];
  years.forEach((year) => {
    const values = yearStage.get(year) || new Map();
    yearStageRows.push([year, ...stages.map((stage) => values.get(stage) || 0)]);
  });

  const yearAppliedInterviewRows = [['year', 'applied', 'interview']];
  years.forEach((year) => {
    const values = yearStage.get(year) || new Map();
    yearAppliedInterviewRows.push([
      year,
      values.get('applied') || 0,
      values.get('interview') || 0
    ]);
  });

  const monthRows = [['month', 'applications']];
  [...monthlyTotals.entries()]
    .filter((entry) => entry[0] !== 'unknown')
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach((entry) => monthRows.push([entry[0], entry[1]]));

  dataSheet.getRange(1, 1, stageRows.length, 2).setValues(stageRows);
  dataSheet.getRange(1, 4, sourceRows.length, 2).setValues(sourceRows);
  dataSheet.getRange(1, 7, yearStageRows.length, yearStageRows[0].length).setValues(yearStageRows);
  dataSheet.getRange(1, 12, monthRows.length, 2).setValues(monthRows);
  dataSheet.getRange(1, 15, yearAppliedInterviewRows.length, 3).setValues(yearAppliedInterviewRows);
  dataSheet.getRange('A1:B1').setFontWeight('bold');
  dataSheet.getRange('D1:E1').setFontWeight('bold');
  dataSheet.getRange(1, 7, 1, yearStageRows[0].length).setFontWeight('bold');
  dataSheet.getRange('L1:M1').setFontWeight('bold');
  dataSheet.getRange('O1:Q1').setFontWeight('bold');

  const stageChart = dashboardSheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataSheet.getRange(1, 1, stageRows.length, 2))
    .setPosition(4, 1, 0, 0)
    .setOption('title', 'Stage Distribution')
    .setOption('legend', { position: 'right' })
    .build();
  dashboardSheet.insertChart(stageChart);

  const sourceChart = dashboardSheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataSheet.getRange(1, 4, sourceRows.length, 2))
    .setPosition(4, 8, 0, 0)
    .setOption('title', 'Source Distribution')
    .setOption('legend', { position: 'right' })
    .build();
  dashboardSheet.insertChart(sourceChart);

  const yearStageChart = dashboardSheet
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataSheet.getRange(1, 7, yearStageRows.length, yearStageRows[0].length))
    .setPosition(22, 1, 0, 0)
    .setOption('title', 'Stage Trend by Year')
    .setOption('isStacked', true)
    .build();
  dashboardSheet.insertChart(yearStageChart);

  if (monthRows.length > 1) {
    const monthChart = dashboardSheet
      .newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 12, monthRows.length, 2))
      .setPosition(22, 8, 0, 0)
      .setOption('title', 'Applications Over Time (Monthly)')
      .build();
    dashboardSheet.insertChart(monthChart);
  }

  if (yearAppliedInterviewRows.length > 1) {
    const appliedInterviewChart = dashboardSheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataSheet.getRange(1, 15, yearAppliedInterviewRows.length, 3))
      .setPosition(40, 1, 0, 0)
      .setOption('title', 'Interview vs Applied by Year')
      .build();
    dashboardSheet.insertChart(appliedInterviewChart);
  }

  dataSheet.hideSheet();
}

function buildDefaultFollowUpQueue() {
  buildFollowUpQueue(7, 25);
}

function buildFollowUpQueue(daysWithoutTouch, maxItems) {
  const ss = getSpreadsheet_();
  setupSheets();

  const appSheet = ss.getSheetByName('Applications');
  const queueSheet = ss.getSheetByName('FollowUpQueue');

  const last = appSheet.getLastRow();
  const output = [['company', 'role', 'source', 'applied_date', 'days_since_apply', 'thread_url']];

  if (last > 1) {
    const rows = appSheet.getRange(2, 1, last - 1, APP_COLUMNS.length).getValues();
    const now = new Date();

    const candidates = rows
      .map((r) => {
        const stage = sanitize_(r[4]).toLowerCase();
        const applied = sanitize_(r[6]);
        const appliedDate = applied ? new Date(applied) : null;
        const days = appliedDate && !isNaN(appliedDate.getTime())
          ? Math.floor((now.getTime() - appliedDate.getTime()) / (1000 * 60 * 60 * 24))
          : 0;

        return {
          company: sanitize_(r[2]),
          role: sanitize_(r[3]),
          source: sanitize_(r[1]),
          appliedDate: applied,
          daysSince: days,
          threadUrl: sanitize_(r[10]),
          stage
        };
      })
      .filter((row) => row.stage === 'applied' && row.daysSince >= daysWithoutTouch)
      .sort((a, b) => b.daysSince - a.daysSince)
      .slice(0, maxItems);

    candidates.forEach((c) => {
      output.push([c.company, c.role, c.source, c.appliedDate, c.daysSince, c.threadUrl]);
    });
  }

  queueSheet.clear();
  queueSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  queueSheet.setFrozenRows(1);
}
