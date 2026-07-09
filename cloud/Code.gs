const SPREADSHEET_ID = "1QXTqq1Q-ahUFwoaDH6bZaXJGuuvz-QyX7h3J9qPjjk8";
const TIME_ZONE = "Asia/Taipei";
const DEADLINE_DAY = 5;

function doPost(e) {
  const payloadText = e && e.parameter && e.parameter.payload
    ? e.parameter.payload
    : e && e.postData && e.postData.contents
      ? e.postData.contents
      : "";
  const payload = JSON.parse(payloadText || "{}");
  saveSubmission_(payload);
  return json_({ ok: true, submissionId: payload.submissionId || "" });
}

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = params.action || "latest";
  let data;
  if (action === "status") {
    requireAdmin_(params.adminKey);
    data = { ok: true, status: getMonthlyStatus_(params.month) };
  } else if (action === "aggregate") {
    requireAdmin_(params.adminKey);
    data = { ok: true, aggregate: getMonthlyAggregate_(params.month) };
  } else {
    data = { ok: true, latest: getLatestSubmission_() };
  }
  const callback = params.callback;
  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(data)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(data);
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function requireAdmin_(value) {
  const sheet = ss_().getSheetByName("系統設定");
  const expected = sheet ? String(sheet.getRange("B2").getDisplayValue()).trim() : "";
  if (!expected || String(value || "").trim() !== expected) {
    throw new Error("管理金鑰不正確。");
  }
}

function getExpected_() {
  const sheet = ss_().getSheetByName("系統設定");
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 4, sheet.getLastRow() - 1, 3).getDisplayValues()
    .filter((row) => row[1])
    .map((row, index) => ({
      office: row[0] || "",
      code: row[1] || "",
      vendor: row[2] || "",
      group: row[0] === "第二區工程處" ? 2 : 1,
      order: index,
    }));
}

function saveSubmission_(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    const submissionId = payload.submissionId || `${Date.now()}`;
    const submittedAt = payload.submittedAt || new Date().toISOString();
    const rows = (Array.isArray(payload.rows) ? payload.rows : []).filter((row) => String(row.code || "").trim());
    const managementRows = (Array.isArray(payload.managementRows) ? payload.managementRows : []).filter((row) => String(row.code || "").trim());
    const totals = payload.totals || {};
    const totalWorkers = Number(totals.thailand || 0) + Number(totals.indonesia || 0)
      + Number(totals.philippines || 0) + Number(totals.vietnam || 0);

    if (!payload.date) throw new Error("缺少統計日期。");
    if (!String(payload.reporter || "").trim()) throw new Error("請填寫填報人。");
    if (!rows.length) throw new Error("至少需要一個區段標。");

    ss.getSheetByName("填報紀錄").appendRow([
      submissionId,
      submittedAt,
      payload.date || "",
      payload.reporter || "",
      payload.note || "",
      rows.length,
      managementRows.length,
      Number(totals.thailand || 0),
      totalWorkers,
    ]);

    const personValues = rows.map((row) => [
      submissionId,
      payload.date || "",
      submittedAt,
      payload.reporter || "",
      row.office || "",
      row.code || "",
      row.vendor || "",
      numberOrBlank_(row.thailand),
      numberOrBlank_(row.indonesia),
      numberOrBlank_(row.philippines),
      numberOrBlank_(row.vietnam),
      Number(row.thailand || 0) + Number(row.indonesia || 0) + Number(row.philippines || 0) + Number(row.vietnam || 0),
      row.address || "",
      row.group || "",
      row.highlight ? "Y" : "",
    ]);
    if (personValues.length) {
      const sheet = ss.getSheetByName("人數明細");
      sheet.getRange(sheet.getLastRow() + 1, 1, personValues.length, personValues[0].length).setValues(personValues);
    }

    const managementValues = managementRows.map((row) => [
      submissionId,
      payload.date || "",
      submittedAt,
      payload.reporter || "",
      row.office || "",
      row.code || "",
      row.dates || "",
      row.group || "",
    ]);
    if (managementValues.length) {
      const sheet = ss.getSheetByName("管理稽查");
      sheet.getRange(sheet.getLastRow() + 1, 1, managementValues.length, managementValues[0].length).setValues(managementValues);
    }
  } finally {
    lock.releaseLock();
  }
}

function numberOrBlank_(value) {
  if (value === "" || value === null || value === undefined) return "";
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : "";
}

function monthKey_(value) {
  return String(value || "").slice(0, 7);
}

function latestPersonByCode_(month) {
  const values = ss_().getSheetByName("人數明細").getDataRange().getValues().slice(1);
  const result = {};
  values.forEach((row) => {
    const code = String(row[5] || "").trim();
    if (!code || monthKey_(row[1]) !== month) return;
    const submittedAt = String(row[2] || "");
    if (!result[code] || submittedAt >= result[code].submittedAt) {
      result[code] = {
        submissionId: row[0] || "",
        date: row[1] || "",
        submittedAt,
        reporter: row[3] || "",
        office: row[4] || "",
        code,
        vendor: row[6] || "",
        thailand: row[7] === "" ? "" : String(row[7]),
        indonesia: row[8] === "" ? "" : String(row[8]),
        philippines: row[9] === "" ? "" : String(row[9]),
        vietnam: row[10] === "" ? "" : String(row[10]),
        address: row[12] || "",
        group: Number(row[13] || 1),
        highlight: row[14] === "Y",
      };
    }
  });
  return result;
}

function latestManagementByCode_(month) {
  const values = ss_().getSheetByName("管理稽查").getDataRange().getValues().slice(1);
  const result = {};
  values.forEach((row) => {
    const code = String(row[5] || "").trim();
    if (!code || monthKey_(row[1]) !== month) return;
    const submittedAt = String(row[2] || "");
    if (!result[code] || submittedAt >= result[code].submittedAt) {
      result[code] = {
        submissionId: row[0] || "",
        date: row[1] || "",
        submittedAt,
        reporter: row[3] || "",
        office: row[4] || "",
        code,
        dates: row[6] || "",
        group: Number(row[7] || 1),
      };
    }
  });
  return result;
}

function submittedDay_(submittedAt) {
  const value = new Date(submittedAt);
  if (Number.isNaN(value.getTime())) return null;
  return Number(Utilities.formatDate(value, TIME_ZONE, "d"));
}

function getMonthlyStatus_(monthInput) {
  const month = monthKey_(monthInput || Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM"));
  const people = latestPersonByCode_(month);
  const management = latestManagementByCode_(month);
  const items = getExpected_().map((expected) => {
    const person = people[expected.code] || null;
    const inspection = management[expected.code] || null;
    const complete = Boolean(person && inspection && person.reporter && person.vendor);
    const day = person ? submittedDay_(person.submittedAt) : null;
    return {
      office: expected.office,
      code: expected.code,
      vendor: expected.vendor,
      reporter: person ? person.reporter : "",
      submittedAt: person ? person.submittedAt : "",
      submittedDay: day,
      complete,
      onTime: complete && day !== null && day <= DEADLINE_DAY,
      status: !person ? "missing" : complete ? (day <= DEADLINE_DAY ? "on-time" : "late") : "incomplete",
    };
  });
  return {
    month,
    deadlineDay: DEADLINE_DAY,
    total: items.length,
    completed: items.filter((item) => item.complete).length,
    onTime: items.filter((item) => item.onTime).length,
    items,
  };
}

function getMonthlyAggregate_(monthInput) {
  const month = monthKey_(monthInput || Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM"));
  const people = latestPersonByCode_(month);
  const management = latestManagementByCode_(month);
  const expected = getExpected_();
  return {
    month,
    date: `${month}-01`,
    rows: expected.map((item) => people[item.code] || {
      office: item.office,
      code: item.code,
      vendor: item.vendor,
      thailand: "",
      indonesia: "",
      philippines: "",
      vietnam: "",
      address: "",
      group: item.group,
      highlight: false,
    }),
    managementRows: expected.map((item) => management[item.code] || {
      office: item.office,
      code: item.code,
      dates: "",
      group: item.group,
    }),
  };
}

function getLatestSubmission_() {
  const ss = ss_();
  const submissions = ss.getSheetByName("填報紀錄").getDataRange().getValues();
  if (submissions.length < 2) return null;
  const latest = submissions[submissions.length - 1];
  const submissionId = latest[0];
  const date = latest[2];
  const reporter = latest[3];
  const note = latest[4];

  const personRows = ss.getSheetByName("人數明細").getDataRange().getValues().slice(1)
    .filter((row) => row[0] === submissionId)
    .map((row) => ({
      office: row[4] || "",
      code: row[5] || "",
      vendor: row[6] || "",
      thailand: row[7] === "" ? "" : String(row[7]),
      indonesia: row[8] === "" ? "" : String(row[8]),
      philippines: row[9] === "" ? "" : String(row[9]),
      vietnam: row[10] === "" ? "" : String(row[10]),
      address: row[12] || "",
      group: Number(row[13] || 1),
      highlight: row[14] === "Y",
    }));

  const managementRows = ss.getSheetByName("管理稽查").getDataRange().getValues().slice(1)
    .filter((row) => row[0] === submissionId)
    .map((row) => ({
      office: row[4] || "",
      code: row[5] || "",
      dates: row[6] || "",
      group: Number(row[7] || 1),
    }));

  return { submissionId, date, reporter, note, rows: personRows, managementRows };
}
