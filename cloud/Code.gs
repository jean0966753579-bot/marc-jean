const SPREADSHEET_ID = "1QXTqq1Q-ahUFwoaDH6bZaXJGuuvz-QyX7h3J9qPjjk8";

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
  const data = { ok: true, latest: getLatestSubmission_() };
  const callback = e && e.parameter && e.parameter.callback;
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

function saveSubmission_(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    const submissionId = payload.submissionId || `${Date.now()}`;
    const submittedAt = payload.submittedAt || new Date().toISOString();
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const managementRows = Array.isArray(payload.managementRows) ? payload.managementRows : [];
    const totals = payload.totals || {};
    const totalWorkers = Number(totals.thailand || 0) + Number(totals.indonesia || 0) + Number(totals.philippines || 0) + Number(totals.vietnam || 0);

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
      ss.getSheetByName("人數明細").getRange(ss.getSheetByName("人數明細").getLastRow() + 1, 1, personValues.length, personValues[0].length).setValues(personValues);
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
      ss.getSheetByName("管理稽查").getRange(ss.getSheetByName("管理稽查").getLastRow() + 1, 1, managementValues.length, managementValues[0].length).setValues(managementValues);
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

  return {
    submissionId,
    date,
    reporter,
    note,
    rows: personRows,
    managementRows,
  };
}
