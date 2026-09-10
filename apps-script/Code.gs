const CONFIG = {
  SPREADSHEET_ID: '',
  SHEET_NAME: 'Survey Responses'
};

const HEADERS = [
  'Timestamp',
  'Submission ID',
  'Customer Name',
  'Phone Number',
  'Last Purchase',
  'Has Repeat Order',
  'Repeat Platform',
  'Other Platform',
  'No Repeat Reason',
  'Why Other Platform',
  'Still Using',
  'Not Using Reason',
  'Product Experience',
  'Reorder Trigger',
  'Website Experience',
  'Win-back Idea'
];

function setupSurveySheet() {
  const sheet = getOrCreateSheet_();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#4f46e5')
    .setFontColor('#ffffff');
  sheet.autoResizeColumns(1, HEADERS.length);
  return 'Survey Responses sheet is ready.';
}

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '').trim();
    if (action !== 'dashboard') {
      return jsonResponse_({ success: true, message: 'NBDY Survey API is active.' });
    }

    const filters = {
      dateFrom: cleanText_(e.parameter.dateFrom, 10),
      dateTo: cleanText_(e.parameter.dateTo, 10),
      repeat: cleanText_(e.parameter.repeat, 3).toLowerCase(),
      using: cleanText_(e.parameter.using, 3).toLowerCase()
    };

    return jsonResponse_({
      success: true,
      dashboard: buildDashboard_(filters)
    });
  } catch (error) {
    console.error(error);
    return jsonResponse_({ success: false, message: error.message || 'Unable to load dashboard.' });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No submission data received.');
    }

    const data = JSON.parse(e.postData.contents);
    const submission = validateSubmission_(data);
    const sheet = getOrCreateSheet_();

    if (submissionExists_(sheet, submission.submissionId)) {
      return jsonResponse_({
        success: true,
        duplicate: true,
        submissionId: submission.submissionId,
        message: 'Submission already recorded.'
      });
    }

    sheet.appendRow([
      new Date(),
      submission.submissionId,
      safeCell_(submission.customerName),
      safeCell_(submission.customerPhone),
      safeCell_(submission.lastPurchase),
      safeCell_(submission.hasRepeatOrder),
      safeCell_(submission.repeatPlatform),
      safeCell_(submission.otherPlatformInput),
      safeCell_(submission.noRepeatReason),
      safeCell_(submission.whyOtherPlatform.join(' | ')),
      safeCell_(submission.stillUsing),
      safeCell_(submission.notUsingReason),
      safeCell_(submission.productExperience),
      safeCell_(submission.reorderTrigger.join(' | ')),
      safeCell_(submission.websiteExperience),
      safeCell_(submission.winBackIdea)
    ]);

    return jsonResponse_({
      success: true,
      duplicate: false,
      submissionId: submission.submissionId,
      message: 'Survey successfully recorded.'
    });
  } catch (error) {
    console.error(error);
    return jsonResponse_({ success: false, message: error.message || 'Unable to save survey.' });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function buildDashboard_(filters) {
  const sheet = getOrCreateSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return emptyDashboard_(filters);

  const headerMap = {};
  values[0].forEach(function(header, index) { headerMap[String(header).trim()] = index; });
  HEADERS.forEach(function(header) {
    if (headerMap[header] === undefined) throw new Error('Missing sheet column: ' + header);
  });

  const timezone = Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur';
  const start = parseDateFilter_(filters.dateFrom, false);
  const end = parseDateFilter_(filters.dateTo, true);
  const rows = values.slice(1).filter(function(row) {
    const timestamp = normaliseDate_(row[headerMap['Timestamp']]);
    const repeat = String(row[headerMap['Has Repeat Order']] || '').toLowerCase();
    const using = String(row[headerMap['Still Using']] || '').toLowerCase();
    if (start && (!timestamp || timestamp < start)) return false;
    if (end && (!timestamp || timestamp > end)) return false;
    if (filters.repeat && repeat !== filters.repeat) return false;
    if (filters.using && using !== filters.using) return false;
    return true;
  });

  const stats = emptyDashboard_(filters);
  stats.total = rows.length;

  rows.forEach(function(row) {
    const timestamp = normaliseDate_(row[headerMap['Timestamp']]);
    const repeat = String(row[headerMap['Has Repeat Order']] || '').toLowerCase();
    const using = String(row[headerMap['Still Using']] || '').toLowerCase();
    const lastPurchase = String(row[headerMap['Last Purchase']] || '').trim();
    const platform = String(row[headerMap['Repeat Platform']] || '').trim();

    if (repeat === 'yes') stats.repeatYes++;
    if (repeat === 'no') stats.repeatNo++;
    if (using === 'yes') stats.stillUsingYes++;
    if (using === 'no') stats.stillUsingNo++;
    increment_(stats.lastPurchaseCounts, lastPurchase);
    increment_(stats.platformCounts, platform);

    splitMulti_(row[headerMap['Reorder Trigger']]).forEach(function(item) {
      increment_(stats.triggerCounts, item);
    });
    splitMulti_(row[headerMap['Why Other Platform']]).forEach(function(item) {
      increment_(stats.otherPlatformReasonCounts, item);
    });

    if (String(row[headerMap['No Repeat Reason']] || '').trim()) stats.textResponseCounts.noRepeatReasons++;
    if (String(row[headerMap['Not Using Reason']] || '').trim()) stats.textResponseCounts.notUsingReasons++;
    if (String(row[headerMap['Product Experience']] || '').trim()) stats.textResponseCounts.productExperiences++;
    if (String(row[headerMap['Website Experience']] || '').trim()) stats.textResponseCounts.websiteExperiences++;
    if (String(row[headerMap['Win-back Idea']] || '').trim()) stats.textResponseCounts.winBackIdeas++;
  });

  stats.repeatRate = percentage_(stats.repeatYes, stats.repeatYes + stats.repeatNo);
  stats.stillUsingRate = percentage_(stats.stillUsingYes, stats.stillUsingYes + stats.stillUsingNo);
  stats.topPlatform = topLabel_(stats.platformCounts);
  stats.topTrigger = topLabel_(stats.triggerCounts);

  const latest = rows.reduce(function(found, row) {
    const value = normaliseDate_(row[headerMap['Timestamp']]);
    return value && (!found || value > found) ? value : found;
  }, null);
  stats.latestResponse = latest
    ? Utilities.formatDate(latest, timezone, 'dd MMM yyyy, HH:mm')
    : '';

  return stats;
}

function emptyDashboard_(filters) {
  const hasDateFilter = filters && (filters.dateFrom || filters.dateTo);
  return {
    total: 0,
    repeatYes: 0,
    repeatNo: 0,
    stillUsingYes: 0,
    stillUsingNo: 0,
    repeatRate: 0,
    stillUsingRate: 0,
    topPlatform: '',
    topTrigger: '',
    latestResponse: '',
    periodLabel: hasDateFilter
      ? ((filters.dateFrom || 'Awal') + ' hingga ' + (filters.dateTo || 'Kini'))
      : 'Semua tempoh',
    lastPurchaseCounts: {
      'Kurang dari sebulan': 0,
      '1-3 bulan lepas': 0,
      '3-6 bulan lepas': 0,
      'Lebih 6 bulan': 0,
      'Tak ingat': 0
    },
    platformCounts: { Website: 0, 'TikTok Shop': 0, Shopee: 0, Other: 0 },
    triggerCounts: {
      'Better price/promotion': 0,
      Voucher: 0,
      'Free shipping': 0,
      'TikTok Live': 0,
      Reminder: 0,
      'Product benefits': 0,
      'Stock running out': 0,
      'Easier ordering': 0
    },
    otherPlatformReasonCounts: {
      'TikTok voucher/discount': 0,
      'Free shipping': 0,
      'Easier/convenient': 0,
      'Saw TikTok Live': 0,
      'Affiliate/influencer': 0,
      'Platform preference': 0,
      Other: 0
    },
    textResponseCounts: {
      noRepeatReasons: 0,
      notUsingReasons: 0,
      productExperiences: 0,
      websiteExperiences: 0,
      winBackIdeas: 0
    }
  };
}

function validateSubmission_(data) {
  if (!data || typeof data !== 'object') throw new Error('Invalid submission format.');
  const submission = {
    submissionId: cleanText_(data.submissionId, 100) || Utilities.getUuid(),
    customerName: cleanText_(data.customerName, 100),
    customerPhone: cleanText_(data.customerPhone, 30),
    lastPurchase: cleanText_(data.lastPurchase, 50),
    hasRepeatOrder: cleanText_(data.hasRepeatOrder, 10).toLowerCase(),
    repeatPlatform: cleanText_(data.repeatPlatform, 50),
    otherPlatformInput: cleanText_(data.otherPlatformInput, 100),
    noRepeatReason: cleanText_(data.noRepeatReason, 1000),
    whyOtherPlatform: cleanArray_(data.whyOtherPlatform, 10, 100),
    stillUsing: cleanText_(data.stillUsing, 10).toLowerCase(),
    notUsingReason: cleanText_(data.notUsingReason, 1000),
    productExperience: cleanText_(data.productExperience, 2000),
    reorderTrigger: cleanArray_(data.reorderTrigger, 10, 100),
    websiteExperience: cleanText_(data.websiteExperience, 2000),
    winBackIdea: cleanText_(data.winBackIdea, 2000)
  };

  if (!submission.customerName) throw new Error('Customer name is required.');
  if (!submission.customerPhone) throw new Error('Phone number is required.');
  if (!/^[0-9+\-()\s]{7,25}$/.test(submission.customerPhone)) throw new Error('Invalid phone number.');
  if (!['Kurang dari sebulan', '1-3 bulan lepas', '3-6 bulan lepas', 'Lebih 6 bulan', 'Tak ingat'].includes(submission.lastPurchase)) throw new Error('Invalid last purchase selection.');
  if (!['yes', 'no'].includes(submission.hasRepeatOrder)) throw new Error('Invalid repeat-order answer.');

  if (submission.hasRepeatOrder === 'yes') {
    if (!['Website', 'TikTok Shop', 'Shopee', 'Other'].includes(submission.repeatPlatform)) throw new Error('Repeat platform is required.');
    if (submission.repeatPlatform === 'Other' && !submission.otherPlatformInput) throw new Error('Please specify the other platform.');
    submission.noRepeatReason = '';
  } else {
    submission.repeatPlatform = '';
    submission.otherPlatformInput = '';
    submission.whyOtherPlatform = [];
    if (!submission.noRepeatReason) throw new Error('No-repeat reason is required.');
  }

  if (!['yes', 'no'].includes(submission.stillUsing)) throw new Error('Invalid product-usage answer.');
  if (submission.stillUsing === 'no' && !submission.notUsingReason) throw new Error('Reason for stopping usage is required.');
  if (submission.stillUsing === 'yes') submission.notUsingReason = '';
  if (!submission.productExperience) throw new Error('Product experience is required.');
  if (!submission.websiteExperience) throw new Error('Website experience is required.');
  if (!submission.winBackIdea) throw new Error('Win-back suggestion is required.');
  return submission;
}

function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Spreadsheet not found. Add the Spreadsheet ID in CONFIG.');
  return spreadsheet;
}

function getOrCreateSheet_() {
  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sheet;
}

function submissionExists_(sheet, submissionId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  return sheet.getRange(2, 2, lastRow - 1, 1)
    .createTextFinder(submissionId).matchEntireCell(true).findNext() !== null;
}

function cleanText_(value, maximumLength) {
  if (value === undefined || value === null) return '';
  return String(value).replace(/\u0000/g, '').trim().slice(0, maximumLength);
}

function cleanArray_(value, maximumItems, maximumLength) {
  if (value === undefined || value === null || value === '') return [];
  const values = Array.isArray(value) ? value : [value];
  return values.slice(0, maximumItems).map(function(item) {
    return cleanText_(item, maximumLength);
  }).filter(Boolean);
}

function safeCell_(value) {
  const text = String(value || '');
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function splitMulti_(value) {
  return String(value || '').split('|').map(function(item) { return item.trim(); }).filter(Boolean);
}

function increment_(target, key) {
  if (!key) return;
  target[key] = Number(target[key] || 0) + 1;
}

function percentage_(part, total) {
  return total ? Math.round((part / total) * 1000) / 10 : 0;
}

function topLabel_(counts) {
  return Object.keys(counts).reduce(function(best, key) {
    return Number(counts[key] || 0) > Number(counts[best] || 0) ? key : best;
  }, '') || '';
}

function normaliseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function parseDateFilter_(value, endOfDay) {
  if (!value) return null;
  const date = new Date(value + 'T00:00:00');
  if (isNaN(date.getTime())) return null;
  if (endOfDay) date.setHours(23, 59, 59, 999);
  return date;
}

function jsonResponse_(result) {
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
