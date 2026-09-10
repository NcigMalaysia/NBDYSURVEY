const CONFIG = {
  // Biarkan kosong untuk Apps Script yang terikat pada Google Sheet.
  SPREADSHEET_ID: "",
  SHEET_NAME: "Survey Responses"
};

const HEADERS = [
  "Timestamp",
  "Submission ID",
  "Customer Name",
  "Phone Number",
  "Last Purchase",
  "Has Repeat Order",
  "Repeat Platform",
  "Other Platform",
  "No Repeat Reason",
  "Why Other Platform",
  "Still Using",
  "Not Using Reason",
  "Product Experience",
  "Reorder Trigger",
  "Website Experience",
  "Win-back Idea"
];

function setupSurveySheet() {
  const sheet = getOrCreateSheet_();

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight("bold")
    .setBackground("#4f46e5")
    .setFontColor("#ffffff");

  sheet.autoResizeColumns(1, HEADERS.length);

  return "Survey Responses sheet is ready.";
}

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || "");

    if (action !== "dashboard") {
      return jsonResponse_({
        success: true,
        message: "NBDY Survey API is active."
      });
    }

    return jsonResponse_({
      success: true,
      dashboard: buildPublicDashboard_()
    });

  } catch (error) {
    console.error(error);

    return jsonResponse_({
      success: false,
      message: error.message || "Unable to load dashboard."
    });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No submission data received.");
    }

    const data = JSON.parse(e.postData.contents);
    const submission = validateSubmission_(data);
    const sheet = getOrCreateSheet_();

    if (submissionExists_(sheet, submission.submissionId)) {
      return jsonResponse_({
        success: true,
        duplicate: true,
        submissionId: submission.submissionId,
        message: "Submission already recorded."
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
      safeCell_(submission.whyOtherPlatform.join(" | ")),
      safeCell_(submission.stillUsing),
      safeCell_(submission.notUsingReason),
      safeCell_(submission.productExperience),
      safeCell_(submission.reorderTrigger.join(" | ")),
      safeCell_(submission.websiteExperience),
      safeCell_(submission.winBackIdea)
    ]);

    return jsonResponse_({
      success: true,
      duplicate: false,
      submissionId: submission.submissionId,
      message: "Survey successfully recorded."
    });

  } catch (error) {
    console.error(error);

    return jsonResponse_({
      success: false,
      message: error.message || "Unable to save survey."
    });

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function buildPublicDashboard_() {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();

  const platformCounts = {
    Website: 0,
    "TikTok Shop": 0,
    Shopee: 0,
    Other: 0
  };

  const triggerCounts = {
    "Better price/promotion": 0,
    Voucher: 0,
    "Free shipping": 0,
    "TikTok Live": 0,
    Reminder: 0,
    "Product benefits": 0,
    "Stock running out": 0,
    "Easier ordering": 0
  };

  if (lastRow < 2) {
    return {
      total: 0,
      repeatCount: 0,
      repeatRate: 0,
      topPlatform: "-",
      platformCounts: platformCounts,
      triggerCounts: triggerCounts
    };
  }

  const rows = sheet
    .getRange(2, 1, lastRow - 1, HEADERS.length)
    .getDisplayValues()
    .filter(function(row) {
      return row[1] || row[2];
    });

  let repeatCount = 0;

  rows.forEach(function(row) {
    const hasRepeatOrder = row[5];
    const repeatPlatform = row[6];
    const reorderTriggers = row[13];

    if (hasRepeatOrder === "yes") {
      repeatCount++;

      if (Object.prototype.hasOwnProperty.call(platformCounts, repeatPlatform)) {
        platformCounts[repeatPlatform]++;
      }
    }

    if (reorderTriggers) {
      reorderTriggers.split(" | ").forEach(function(trigger) {
        if (Object.prototype.hasOwnProperty.call(triggerCounts, trigger)) {
          triggerCounts[trigger]++;
        }
      });
    }
  });

  let topPlatform = "-";
  let topCount = 0;

  Object.keys(platformCounts).forEach(function(platform) {
    if (platformCounts[platform] > topCount) {
      topPlatform = platform;
      topCount = platformCounts[platform];
    }
  });

  return {
    total: rows.length,
    repeatCount: repeatCount,
    repeatRate: rows.length
      ? Math.round((repeatCount / rows.length) * 100)
      : 0,
    topPlatform: topPlatform,
    platformCounts: platformCounts,
    triggerCounts: triggerCounts
  };
}

function validateSubmission_(data) {
  if (!data || typeof data !== "object") {
    throw new Error("Invalid submission format.");
  }

  const submission = {
    submissionId: cleanText_(data.submissionId, 100) || Utilities.getUuid(),
    customerName: cleanText_(data.customerName, 100),
    customerPhone: cleanText_(data.customerPhone, 30),
    lastPurchase: cleanText_(data.lastPurchase, 50),
    hasRepeatOrder: cleanText_(data.hasRepeatOrder, 10),
    repeatPlatform: cleanText_(data.repeatPlatform, 50),
    otherPlatformInput: cleanText_(data.otherPlatformInput, 100),
    noRepeatReason: cleanText_(data.noRepeatReason, 1000),
    whyOtherPlatform: cleanArray_(data.whyOtherPlatform, 10, 100),
    stillUsing: cleanText_(data.stillUsing, 10),
    notUsingReason: cleanText_(data.notUsingReason, 1000),
    productExperience: cleanText_(data.productExperience, 2000),
    reorderTrigger: cleanArray_(data.reorderTrigger, 10, 100),
    websiteExperience: cleanText_(data.websiteExperience, 2000),
    winBackIdea: cleanText_(data.winBackIdea, 2000)
  };

  if (!submission.customerName) {
    throw new Error("Customer name is required.");
  }

  if (!submission.customerPhone) {
    throw new Error("Phone number is required.");
  }

  if (!/^[0-9+\-()\s]{7,25}$/.test(submission.customerPhone)) {
    throw new Error("Invalid phone number.");
  }

  const allowedLastPurchase = [
    "Kurang dari sebulan",
    "1-3 bulan lepas",
    "3-6 bulan lepas",
    "Lebih 6 bulan",
    "Tak ingat"
  ];

  if (!allowedLastPurchase.includes(submission.lastPurchase)) {
    throw new Error("Invalid last purchase selection.");
  }

  if (!["yes", "no"].includes(submission.hasRepeatOrder)) {
    throw new Error("Invalid repeat-order answer.");
  }

  if (submission.hasRepeatOrder === "yes") {
    const allowedPlatforms = [
      "Website",
      "TikTok Shop",
      "Shopee",
      "Other"
    ];

    if (!allowedPlatforms.includes(submission.repeatPlatform)) {
      throw new Error("Repeat platform is required.");
    }

    if (
      submission.repeatPlatform === "Other" &&
      !submission.otherPlatformInput
    ) {
      throw new Error("Please specify the other platform.");
    }

    submission.noRepeatReason = "";

  } else {
    submission.repeatPlatform = "";
    submission.otherPlatformInput = "";
    submission.whyOtherPlatform = [];

    if (!submission.noRepeatReason) {
      throw new Error("No-repeat reason is required.");
    }
  }

  if (!["yes", "no"].includes(submission.stillUsing)) {
    throw new Error("Invalid product-usage answer.");
  }

  if (
    submission.stillUsing === "no" &&
    !submission.notUsingReason
  ) {
    throw new Error("Reason for stopping usage is required.");
  }

  if (submission.stillUsing === "yes") {
    submission.notUsingReason = "";
  }

  if (!submission.productExperience) {
    throw new Error("Product experience is required.");
  }

  if (!submission.winBackIdea) {
    throw new Error("Win-back suggestion is required.");
  }

  return submission;
}

function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      "Spreadsheet not found. Add the Spreadsheet ID in CONFIG."
    );
  }

  return spreadsheet;
}

function getOrCreateSheet_() {
  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function submissionExists_(sheet, submissionId) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return false;
  }

  return sheet
    .getRange(2, 2, lastRow - 1, 1)
    .createTextFinder(submissionId)
    .matchEntireCell(true)
    .findNext() !== null;
}

function cleanText_(value, maximumLength) {
  if (value === undefined || value === null) {
    return "";
  }

  return String(value)
    .replace(/\u0000/g, "")
    .trim()
    .slice(0, maximumLength);
}

function cleanArray_(value, maximumItems, maximumLength) {
  if (value === undefined || value === null || value === "") {
    return [];
  }

  const values = Array.isArray(value) ? value : [value];

  return values
    .slice(0, maximumItems)
    .map(function(item) {
      return cleanText_(item, maximumLength);
    })
    .filter(function(item) {
      return item !== "";
    });
}

function safeCell_(value) {
  const text = String(value || "");

  if (/^[=+\-@]/.test(text)) {
    return "'" + text;
  }

  return text;
}

function jsonResponse_(result) {
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
