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
const DASHBOARD_VERSION = 'v2026-02-23-03';

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
  ensureSheet_(ss, 'CurrentYearStats', ['month', 'applied', 'assessment', 'interview', 'offer', 'rejected', 'ghosted', 'total']);
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
  const index = buildApplicationIndex_(appSheet);
  const existingIds = index.byExternalId;

  ensureLabel_(cfg.sourceLabel);
  ensureLabel_(cfg.processedLabel);

  const query = `label:"${cfg.sourceLabel}" -label:"${cfg.processedLabel}"`;
  const threads = GmailApp.search(query, 0, cfg.searchLimit);

  const processedLabel = GmailApp.getUserLabelByName(cfg.processedLabel);
  const rowsToAppend = [];
  let nextRow = Math.max(2, appSheet.getLastRow() + 1);

  threads.forEach((thread) => {
    const messages = thread.getMessages();
    messages.forEach((message) => {
      const parsed = parseJobEmail_(message, thread);
      if (!parsed.externalId || existingIds.has(parsed.externalId)) {
        return;
      }
      const existingRow = findApplicationRow_(index, parsed.company, parsed.role, parsed.appliedDate, parsed.createdAtUtc);
      if (existingRow) {
        existingIds.set(parsed.externalId, existingRow);
        return;
      }
      rowsToAppend.push([
        parsed.createdAtUtc,
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
      registerApplicationIndexRow_(index, nextRow, parsed.externalId, parsed.company, parsed.role, parsed.appliedDate, parsed.createdAtUtc);
      nextRow += 1;
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
    createdAtUtc: isoUtc_(message.getDate()),
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

function normalizeCompanyKey_(value) {
  const normalized = normalizeMetricToken_(value)
    .replace(/\b(inc|llc|ltd|corp|corporation|company|co|careers|recruiting|jobs)\b/g, '')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  return normalized || '';
}

function normalizeRoleKey_(value) {
  const normalized = normalizeMetricToken_(value)
    .replace(/\b(position|role|opening|job)\b/g, '')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  return normalized || '';
}

function normalizeDateKey_(appliedDate, createdAtUtc) {
  const normalized = normalizeDate_(appliedDate);
  if (normalized) {
    return normalized;
  }
  const fallback = parseDateSafe_(createdAtUtc);
  if (!fallback) {
    return '';
  }
  return Utilities.formatDate(fallback, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function buildApplicationPrimaryKey_(company, role, appliedDate, createdAtUtc) {
  const companyKey = normalizeCompanyKey_(company);
  const roleKey = normalizeRoleKey_(role) || 'unknown';
  const dateKey = normalizeDateKey_(appliedDate, createdAtUtc);
  if (!companyKey || !dateKey) {
    return '';
  }
  return `${companyKey}|${roleKey}|${dateKey}`;
}

function buildApplicationFallbackKey_(company, appliedDate, createdAtUtc) {
  const companyKey = normalizeCompanyKey_(company);
  const dateKey = normalizeDateKey_(appliedDate, createdAtUtc);
  if (!companyKey || !dateKey) {
    return '';
  }
  return `${companyKey}|${dateKey}`;
}

function buildApplicationIndex_(appSheet) {
  const byExternalId = new Map();
  const byPrimaryKey = new Map();
  const byFallbackKey = new Map();
  const rowSource = new Map();
  const last = appSheet.getLastRow();

  if (last <= 1) {
    return {
      byExternalId,
      byPrimaryKey,
      byFallbackKey,
      rowSource
    };
  }

  const rows = appSheet.getRange(2, 1, last - 1, APP_COLUMNS.length).getValues();
  rows.forEach((row, idx) => {
    const rowIndex = idx + 2;
    const createdAtUtc = sanitize_(row[0]);
    const source = sanitize_(row[1]);
    const company = sanitize_(row[2]);
    const role = sanitize_(row[3]);
    const appliedDate = sanitize_(row[6]);
    const externalId = sanitize_(row[7]);

    if (externalId) {
      byExternalId.set(externalId, rowIndex);
    }

    const primary = buildApplicationPrimaryKey_(company, role, appliedDate, createdAtUtc);
    if (primary) {
      byPrimaryKey.set(primary, rowIndex);
    }

    const fallback = buildApplicationFallbackKey_(company, appliedDate, createdAtUtc);
    if (fallback) {
      byFallbackKey.set(fallback, rowIndex);
    }

    rowSource.set(rowIndex, source);
  });

  return {
    byExternalId,
    byPrimaryKey,
    byFallbackKey,
    rowSource
  };
}

function registerApplicationIndexRow_(index, rowIndex, externalId, company, role, appliedDate, createdAtUtc, source) {
  const id = sanitize_(externalId);
  if (id) {
    index.byExternalId.set(id, rowIndex);
  }

  const primary = buildApplicationPrimaryKey_(company, role, appliedDate, createdAtUtc);
  if (primary) {
    index.byPrimaryKey.set(primary, rowIndex);
  }

  const fallback = buildApplicationFallbackKey_(company, appliedDate, createdAtUtc);
  if (fallback) {
    index.byFallbackKey.set(fallback, rowIndex);
  }

  if (source) {
    index.rowSource.set(rowIndex, source);
  }
}

function findApplicationRow_(index, company, role, appliedDate, createdAtUtc) {
  const primary = buildApplicationPrimaryKey_(company, role, appliedDate, createdAtUtc);
  if (primary && index.byPrimaryKey.has(primary)) {
    return index.byPrimaryKey.get(primary);
  }

  const fallback = buildApplicationFallbackKey_(company, appliedDate, createdAtUtc);
  if (fallback && index.byFallbackKey.has(fallback)) {
    return index.byFallbackKey.get(fallback);
  }

  return 0;
}

function isSimplifySource_(source) {
  return /simplify/i.test(sanitize_(source));
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
  const index = buildApplicationIndex_(appSheet);
  const existingIds = index.byExternalId;
  const sourceLabel = sourceName || 'Simplify';

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
  const rowUpdates = [];
  let nextRow = Math.max(2, appSheet.getLastRow() + 1);
  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const company = fromIndex_(row, idx.company) || 'Unknown';
    const role = fromIndex_(row, idx.role) || 'Unknown';
    const status = fromIndex_(row, idx.status) || 'Applied';
    const appliedDate = normalizeDate_(fromIndex_(row, idx.appliedDate));
    const url = fromIndex_(row, idx.url);
    const ext = `simplify_${company}_${role}_${appliedDate}`.toLowerCase().replace(/\s+/g, '_');

    const existingRow = existingIds.get(ext) || findApplicationRow_(index, company, role, appliedDate, '');

    if (existingRow) {
      const current = appSheet.getRange(existingRow, 1, 1, APP_COLUMNS.length).getValues()[0];
      const existingStage = normalizeMetricToken_(current[4]);
      const incomingStage = normalizeMetricToken_(status || 'applied');
      const preferredStage = stageRank_(incomingStage) > stageRank_(existingStage) ? incomingStage : existingStage;
      current[1] = sourceLabel;
      current[2] = company;
      current[3] = role;
      current[4] = toTitleLabel_(preferredStage);
      current[5] = toTitleLabel_(preferredStage);
      if (appliedDate) {
        current[6] = appliedDate;
      }
      current[7] = ext;
      if (url) {
        current[10] = url;
      }
      rowUpdates.push({ rowIndex: existingRow, values: current });
      registerApplicationIndexRow_(index, existingRow, ext, company, role, appliedDate, current[0], sourceLabel);
      continue;
    }

    if (existingIds.has(ext)) {
      continue;
    }

    toAppend.push([
      isoUtc_(new Date()),
      sourceLabel,
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

    registerApplicationIndexRow_(index, nextRow, ext, company, role, appliedDate, isoUtc_(new Date()), sourceLabel);
    nextRow += 1;
  }

  rowUpdates.forEach((update) => {
    appSheet.getRange(update.rowIndex, 1, 1, APP_COLUMNS.length).setValues([update.values]);
  });

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

function stageRank_(stageToken) {
  const stage = normalizeMetricToken_(stageToken);
  const ranks = {
    unknown: 0,
    updated: 1,
    applied: 2,
    ghosted: 2,
    assessment: 3,
    interview: 4,
    rejected: 5,
    offer: 6
  };
  return ranks[stage] || 0;
}

function isUnknownEntity_(value) {
  const normalized = normalizeMetricToken_(value);
  return !normalized || normalized === 'unknown' || normalized === 'na' || normalized === 'n_a';
}

function pickPreferredEntity_(first, second) {
  if (!isUnknownEntity_(first)) {
    return first;
  }
  if (!isUnknownEntity_(second)) {
    return second;
  }
  return first || second || 'Unknown';
}

function pickPreferredDate_(first, second) {
  const firstDate = normalizeDate_(first);
  if (firstDate) {
    return firstDate;
  }
  const secondDate = normalizeDate_(second);
  if (secondDate) {
    return secondDate;
  }
  return sanitize_(first) || sanitize_(second) || '';
}

function mergeCanonicalApplication_(left, right) {
  const leftIsSimplify = isSimplifySource_(left.sourceRaw);
  const rightIsSimplify = isSimplifySource_(right.sourceRaw);

  let base = left;
  if (rightIsSimplify && !leftIsSimplify) {
    base = right;
  } else if (!rightIsSimplify && !leftIsSimplify) {
    const leftRank = stageRank_(left.stage);
    const rightRank = stageRank_(right.stage);
    if (rightRank > leftRank) {
      base = right;
    } else if (rightRank === leftRank) {
      const leftDate = parseDateSafe_(left.createdAtUtc);
      const rightDate = parseDateSafe_(right.createdAtUtc);
      if (rightDate && leftDate && rightDate.getTime() > leftDate.getTime()) {
        base = right;
      }
    }
  }

  const strongerStage = stageRank_(right.stage) > stageRank_(left.stage) ? right.stage : left.stage;
  return {
    ...base,
    source: (leftIsSimplify || rightIsSimplify) ? 'simplify' : base.source,
    sourceRaw: (leftIsSimplify || rightIsSimplify) ? 'Simplify' : base.sourceRaw,
    company: pickPreferredEntity_(left.company, right.company),
    role: pickPreferredEntity_(left.role, right.role),
    appliedDate: pickPreferredDate_(left.appliedDate, right.appliedDate),
    stage: strongerStage,
    threadUrl: sanitize_(left.threadUrl) || sanitize_(right.threadUrl),
    emailSubject: sanitize_(left.emailSubject) || sanitize_(right.emailSubject),
    emailFrom: sanitize_(left.emailFrom) || sanitize_(right.emailFrom),
    externalId: sanitize_(left.externalId) || sanitize_(right.externalId)
  };
}

function buildCanonicalApplications_(rows, cfg) {
  const now = new Date();
  const deduped = new Map();

  rows.forEach((row, idx) => {
    const createdAtUtc = sanitize_(row[0]);
    const sourceRaw = sanitize_(row[1]) || 'Unknown';
    const source = isSimplifySource_(sourceRaw) ? 'simplify' : normalizeMetricToken_(sourceRaw);
    const company = sanitize_(row[2]) || 'Unknown';
    const role = sanitize_(row[3]) || 'Unknown';
    const rawStage = sanitize_(row[4]);
    const appliedDate = sanitize_(row[6]);
    const externalId = sanitize_(row[7]);
    const emailSubject = sanitize_(row[8]);
    const emailFrom = sanitize_(row[9]);
    const threadUrl = sanitize_(row[10]);
    const stage = resolveStageForMetrics_(rawStage, appliedDate, createdAtUtc, now, cfg.ghostedDays);
    const year = inferYearBucket_(appliedDate, createdAtUtc);
    const month = inferMonthBucket_(appliedDate, createdAtUtc);
    const primary = buildApplicationPrimaryKey_(company, role, appliedDate, createdAtUtc);
    const fallback = buildApplicationFallbackKey_(company, appliedDate, createdAtUtc);
    const key = primary ? `p:${primary}` : (fallback ? `f:${fallback}` : (externalId ? `e:${externalId}` : `r:${idx}`));

    const candidate = {
      rowIndex: idx + 2,
      createdAtUtc,
      sourceRaw,
      source,
      company,
      role,
      stage,
      appliedDate,
      year,
      month,
      externalId,
      emailSubject,
      emailFrom,
      threadUrl
    };

    if (!deduped.has(key)) {
      deduped.set(key, candidate);
      return;
    }

    deduped.set(key, mergeCanonicalApplication_(deduped.get(key), candidate));
  });

  return [...deduped.values()];
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

function chartTextStyle_() {
  return { color: '#202124', fontSize: 12 };
}

function chartThemeOptions_(title) {
  return {
    title,
    titleTextStyle: { color: '#202124', fontSize: 18, bold: true },
    legend: { position: 'right', textStyle: chartTextStyle_() },
    backgroundColor: '#ffffff',
    chartArea: { backgroundColor: '#ffffff', left: 70, top: 60, width: '70%', height: '65%' },
    hAxis: { textStyle: chartTextStyle_(), titleTextStyle: chartTextStyle_() },
    vAxis: { textStyle: chartTextStyle_(), titleTextStyle: chartTextStyle_() }
  };
}

function chartPalette_(count) {
  const palette = [
    '#1a73e8', '#ea4335', '#fbbc04', '#34a853',
    '#9334e6', '#00acc1', '#e37400', '#5f6368'
  ];
  const colors = [];
  for (let i = 0; i < count; i += 1) {
    colors.push(palette[i % palette.length]);
  }
  return colors;
}

function listMonthsForYear_(year) {
  const months = [];
  for (let month = 1; month <= 12; month += 1) {
    months.push(`${year}-${String(month).padStart(2, '0')}`);
  }
  return months;
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
  metrics.set('raw_total_rows', Math.max(0, last - 1));
  metrics.set('ghosted_days_threshold', cfg.ghostedDays);

  if (last > 1) {
    const rawRows = appSheet.getRange(2, 1, last - 1, APP_COLUMNS.length).getValues();
    const canonicalRows = buildCanonicalApplications_(rawRows, cfg);
    metrics.set('total_rows', canonicalRows.length);

    canonicalRows.forEach((row) => {
      const source = normalizeMetricToken_(row.source);
      const company = sanitize_(row.company);
      const role = sanitize_(row.role);
      const stage = normalizeMetricToken_(row.stage);
      const year = row.year || 'unknown';

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
  const currentYearSheet = getOrCreateSheet_(ss, 'CurrentYearStats');

  const last = appSheet.getLastRow();
  dataSheet.clear();
  currentYearSheet.clear();
  dashboardSheet.clear();
  dashboardSheet.getCharts().forEach((chart) => dashboardSheet.removeChart(chart));
  dashboardSheet.getRange(1, 1, 120, 30).setBackground('#ffffff').setFontColor('#202124');

  dashboardSheet.getRange('A1').setValue(`Job Funnel Dashboard ${DASHBOARD_VERSION}`).setFontWeight('bold').setFontSize(16);
  dashboardSheet.getRange('A2').setValue(`Last rebuilt: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`);
  dashboardSheet.getRange('A3').setValue('Stage Trend by Year = yearly stacked count of stage-classified application records.');

  if (last <= 1) {
    dashboardSheet.getRange('A4').setValue('No application rows found. Run syncJobEmails first.');
    dataSheet.getRange(1, 1).setValue('No data');
    currentYearSheet.getRange(1, 1).setValue('No data');
    dataSheet.hideSheet();
    return;
  }

  const rawRows = appSheet.getRange(2, 1, last - 1, APP_COLUMNS.length).getValues();
  const rows = buildCanonicalApplications_(rawRows, cfg);
  const stageTotals = new Map();
  const sourceTotals = new Map();
  const monthlyTotals = new Map();
  const yearStage = new Map();
  const currentYearMonthly = new Map();
  const currentYear = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  dashboardSheet.getRange('A4').setValue(`Canonical rows used: ${rows.length}`);
  dashboardSheet.getRange('A5').setValue(`Current year: ${currentYear}`);

  rows.forEach((row) => {
    const source = normalizeMetricToken_(row.source);
    const stage = normalizeMetricToken_(row.stage);
    const year = row.year || 'unknown';
    const month = row.month || 'unknown';

    incrementMap_(stageTotals, stage, 1);
    incrementMap_(sourceTotals, source, 1);
    incrementMap_(monthlyTotals, month, 1);
    incrementYearMetric_(yearStage, year, stage, 1);

    if (year === currentYear && month !== 'unknown') {
      incrementYearMetric_(currentYearMonthly, month, stage, 1);
      incrementYearMetric_(currentYearMonthly, month, 'total', 1);
    }
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

  const currentYearRows = [['month', 'applied', 'assessment', 'interview', 'offer', 'rejected', 'ghosted', 'total']];
  listMonthsForYear_(currentYear).forEach((month) => {
    const bucket = currentYearMonthly.get(month) || new Map();
    currentYearRows.push([
      month,
      bucket.get('applied') || 0,
      bucket.get('assessment') || 0,
      bucket.get('interview') || 0,
      bucket.get('offer') || 0,
      bucket.get('rejected') || 0,
      bucket.get('ghosted') || 0,
      bucket.get('total') || 0
    ]);
  });

  dataSheet.getRange(1, 1, stageRows.length, 2).setValues(stageRows);
  dataSheet.getRange(1, 4, sourceRows.length, 2).setValues(sourceRows);
  dataSheet.getRange(1, 7, yearStageRows.length, yearStageRows[0].length).setValues(yearStageRows);
  dataSheet.getRange(1, 12, monthRows.length, 2).setValues(monthRows);
  dataSheet.getRange(1, 15, yearAppliedInterviewRows.length, 3).setValues(yearAppliedInterviewRows);
  dataSheet.getRange(1, 19, currentYearRows.length, currentYearRows[0].length).setValues(currentYearRows);
  dataSheet.getRange('A1:B1').setFontWeight('bold');
  dataSheet.getRange('D1:E1').setFontWeight('bold');
  dataSheet.getRange(1, 7, 1, yearStageRows[0].length).setFontWeight('bold');
  dataSheet.getRange('L1:M1').setFontWeight('bold');
  dataSheet.getRange('O1:Q1').setFontWeight('bold');
  dataSheet.getRange(1, 19, 1, currentYearRows[0].length).setFontWeight('bold');

  currentYearSheet.getRange(1, 1, currentYearRows.length, currentYearRows[0].length).setValues(currentYearRows);
  currentYearSheet.setFrozenRows(1);
  currentYearSheet.getRange('A1:H1').setFontWeight('bold');
  currentYearSheet.getRange('J1').setValue('Current year monthly stats are the primary operational view.');
  dashboardSheet.getRange('A6').setValue(`Stage rows: ${Math.max(0, stageRows.length - 1)}`);
  dashboardSheet.getRange('A7').setValue(`Source rows: ${Math.max(0, sourceRows.length - 1)}`);
  dashboardSheet.getRange('A8').setValue(`Year-stage rows: ${Math.max(0, yearStageRows.length - 1)}`);
  dashboardSheet.getRange('A9').setValue(`Monthly rows: ${Math.max(0, monthRows.length - 1)}`);
  dashboardSheet.getRange('A10').setValue(`Current year rows: ${Math.max(0, currentYearRows.length - 1)}`);

  if (stageRows.length > 1) {
    const stageChart = dashboardSheet
      .newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dataSheet.getRange(1, 1, stageRows.length, 2))
      .setNumHeaders(1)
      .setPosition(4, 1, 0, 0)
      .setOption('title', chartThemeOptions_('Stage Distribution').title)
      .setOption('titleTextStyle', chartThemeOptions_('Stage Distribution').titleTextStyle)
      .setOption('legend', chartThemeOptions_('Stage Distribution').legend)
      .setOption('backgroundColor', chartThemeOptions_('Stage Distribution').backgroundColor)
      .setOption('chartArea', chartThemeOptions_('Stage Distribution').chartArea)
      .setOption('colors', chartPalette_(Math.max(1, stageRows[0].length - 1)))
      .build();
    dashboardSheet.insertChart(stageChart);
  } else {
    dashboardSheet.getRange('A12').setValue('Stage chart not built: no stage data rows.');
  }

  if (sourceRows.length > 1) {
    const sourceChart = dashboardSheet
      .newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dataSheet.getRange(1, 4, sourceRows.length, 2))
      .setNumHeaders(1)
      .setPosition(4, 8, 0, 0)
      .setOption('title', chartThemeOptions_('Source Distribution').title)
      .setOption('titleTextStyle', chartThemeOptions_('Source Distribution').titleTextStyle)
      .setOption('legend', chartThemeOptions_('Source Distribution').legend)
      .setOption('backgroundColor', chartThemeOptions_('Source Distribution').backgroundColor)
      .setOption('chartArea', chartThemeOptions_('Source Distribution').chartArea)
      .setOption('colors', chartPalette_(Math.max(1, sourceRows[0].length - 1)))
      .build();
    dashboardSheet.insertChart(sourceChart);
  } else {
    dashboardSheet.getRange('H12').setValue('Source chart not built: no source data rows.');
  }

  if (yearStageRows.length > 1 && yearStageRows[0].length > 1) {
    const yearStageChart = dashboardSheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataSheet.getRange(1, 7, yearStageRows.length, yearStageRows[0].length))
      .setNumHeaders(1)
      .setPosition(22, 1, 0, 0)
      .setOption('title', chartThemeOptions_('Yearly Stage Counts (Stacked)').title)
      .setOption('titleTextStyle', chartThemeOptions_('Yearly Stage Counts (Stacked)').titleTextStyle)
      .setOption('legend', chartThemeOptions_('Yearly Stage Counts (Stacked)').legend)
      .setOption('backgroundColor', chartThemeOptions_('Yearly Stage Counts (Stacked)').backgroundColor)
      .setOption('chartArea', chartThemeOptions_('Yearly Stage Counts (Stacked)').chartArea)
      .setOption('hAxis', { title: 'Year', textStyle: chartTextStyle_(), titleTextStyle: chartTextStyle_() })
      .setOption('vAxis', { title: 'Count', textStyle: chartTextStyle_(), titleTextStyle: chartTextStyle_() })
      .setOption('colors', chartPalette_(Math.max(1, yearStageRows[0].length - 1)))
      .setOption('isStacked', true)
      .build();
    dashboardSheet.insertChart(yearStageChart);
  } else {
    dashboardSheet.getRange('A22').setValue('Yearly stage chart not built: no yearly stage data rows.');
  }

  if (monthRows.length > 1) {
    const monthChart = dashboardSheet
      .newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 12, monthRows.length, 2))
      .setNumHeaders(1)
      .setPosition(22, 8, 0, 0)
      .setOption('title', chartThemeOptions_('Applications Over Time (Monthly)').title)
      .setOption('titleTextStyle', chartThemeOptions_('Applications Over Time (Monthly)').titleTextStyle)
      .setOption('legend', chartThemeOptions_('Applications Over Time (Monthly)').legend)
      .setOption('backgroundColor', chartThemeOptions_('Applications Over Time (Monthly)').backgroundColor)
      .setOption('chartArea', chartThemeOptions_('Applications Over Time (Monthly)').chartArea)
      .setOption('hAxis', { title: 'Month', textStyle: chartTextStyle_(), titleTextStyle: chartTextStyle_() })
      .setOption('vAxis', { title: 'Applications', textStyle: chartTextStyle_(), titleTextStyle: chartTextStyle_() })
      .setOption('colors', ['#1a73e8'])
      .build();
    dashboardSheet.insertChart(monthChart);
  }

  if (yearAppliedInterviewRows.length > 1) {
    const appliedInterviewChart = dashboardSheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataSheet.getRange(1, 15, yearAppliedInterviewRows.length, 3))
      .setNumHeaders(1)
      .setPosition(40, 1, 0, 0)
      .setOption('title', chartThemeOptions_('Interview vs Applied by Year').title)
      .setOption('titleTextStyle', chartThemeOptions_('Interview vs Applied by Year').titleTextStyle)
      .setOption('legend', chartThemeOptions_('Interview vs Applied by Year').legend)
      .setOption('backgroundColor', chartThemeOptions_('Interview vs Applied by Year').backgroundColor)
      .setOption('chartArea', chartThemeOptions_('Interview vs Applied by Year').chartArea)
      .setOption('hAxis', { title: 'Year', textStyle: chartTextStyle_(), titleTextStyle: chartTextStyle_() })
      .setOption('vAxis', { title: 'Count', textStyle: chartTextStyle_(), titleTextStyle: chartTextStyle_() })
      .setOption('colors', ['#1a73e8', '#ea4335'])
      .build();
    dashboardSheet.insertChart(appliedInterviewChart);
  }

  if (currentYearRows.length > 1) {
    const currentYearChart = dashboardSheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataSheet.getRange(1, 19, currentYearRows.length, 8))
      .setNumHeaders(1)
      .setPosition(40, 8, 0, 0)
      .setOption('title', chartThemeOptions_(`${currentYear} Monthly Funnel`).title)
      .setOption('titleTextStyle', chartThemeOptions_(`${currentYear} Monthly Funnel`).titleTextStyle)
      .setOption('legend', chartThemeOptions_(`${currentYear} Monthly Funnel`).legend)
      .setOption('backgroundColor', chartThemeOptions_(`${currentYear} Monthly Funnel`).backgroundColor)
      .setOption('chartArea', chartThemeOptions_(`${currentYear} Monthly Funnel`).chartArea)
      .setOption('hAxis', { title: 'Month', textStyle: chartTextStyle_(), titleTextStyle: chartTextStyle_() })
      .setOption('vAxis', { title: 'Count', textStyle: chartTextStyle_(), titleTextStyle: chartTextStyle_() })
      .setOption('colors', ['#1a73e8', '#fbbc04', '#34a853', '#ea4335', '#c5221f', '#5f6368', '#00acc1'])
      .setOption('isStacked', false)
      .build();
    dashboardSheet.insertChart(currentYearChart);
  }

  // Keep DashboardData visible for troubleshooting when chart rendering looks blank.
}

function buildDefaultFollowUpQueue() {
  buildFollowUpQueue(7, 25);
}

function buildFollowUpQueue(daysWithoutTouch, maxItems) {
  const cfg = getConfig();
  const ss = getSpreadsheet_();
  setupSheets();

  const appSheet = ss.getSheetByName('Applications');
  const queueSheet = ss.getSheetByName('FollowUpQueue');

  const last = appSheet.getLastRow();
  const output = [['company', 'role', 'source', 'applied_date', 'days_since_apply', 'thread_url']];

  if (last > 1) {
    const rawRows = appSheet.getRange(2, 1, last - 1, APP_COLUMNS.length).getValues();
    const rows = buildCanonicalApplications_(rawRows, cfg);
    const now = new Date();

    const candidates = rows
      .map((row) => {
        const stage = normalizeMetricToken_(row.stage);
        const applied = sanitize_(row.appliedDate);
        const appliedDate = applied ? new Date(applied) : null;
        const days = appliedDate && !isNaN(appliedDate.getTime())
          ? Math.floor((now.getTime() - appliedDate.getTime()) / (1000 * 60 * 60 * 24))
          : 0;

        return {
          company: sanitize_(row.company),
          role: sanitize_(row.role),
          source: sanitize_(row.sourceRaw || row.source),
          appliedDate: applied,
          daysSince: days,
          threadUrl: sanitize_(row.threadUrl),
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
