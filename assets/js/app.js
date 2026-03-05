/**
 * InCheck Lite Session Tracker API (Robust header matching)
 * Spreadsheet file name: "InCheck Lite Session Tracker"
 *
 * Endpoints:
 *  GET     /exec?action=summary
 *  GET     /exec?action=health
 *  POST    /exec?key=YOUR_KEY   (JSON body)
 *  OPTIONS /exec                (CORS preflight)
 */

const SPREADSHEET_NAME = "InCheck Lite Session Tracker";

// Required headers (logical names). Actual sheet headers may differ by case/spaces; we normalize.
const REQUIRED_RESPONSE_HEADERS = [
  "Timestamp",
  "CSM In charge",
  "CLIENT",
  "Account name",
  "Number of attendees",
  "Main Contact Name",
  "Date of Session",
  "Duration of session (Minutes)",
  "Upload Post-session brief (PDF)",
  "Additional notes"
];

// Commitments tab and headers
const SHEET_COMMIT = "Commitments";
const COMMIT_HEADERS = ["CLIENT", "Account name", "Committed Minutes"];

// API key stored in Script Properties
const API_KEY_PROP = "DASH_API_KEY";

// Near-limit threshold in minutes (120 = 2 hours)
const NEAR_LIMIT_THRESHOLD_MIN = 120;

/** -----------------------------
 * HTTP handlers
 * ----------------------------- */
function doGet(e) {
  const action = (((e || {}).parameter || {}).action || "summary").toLowerCase();

  if (action === "summary") {
    return jsonOutput_(buildSummary_());
  }

  if (action === "health") {
    const ss = getSpreadsheet_();
    const respSheet = findResponseSheet_();
    const commitSheet = getOrCreateCommitmentsSheet_();

    // Validate (robust) so health tells you immediately if something is off
    assertHeaders_(respSheet, REQUIRED_RESPONSE_HEADERS);
    ensureHeaders_(commitSheet, COMMIT_HEADERS);

    return jsonOutput_({
      ok: true,
      spreadsheet: ss.getName(),
      responseSheet: respSheet.getName(),
      commitmentsSheet: commitSheet.getName()
    });
  }

  return jsonOutput_({ error: "Unknown action" }, 400);
}

function doPost(e) {
  const apiKey = getApiKey_();
  const payload = safeJson_((e && e.postData && e.postData.contents) || "");

  // Key can be provided via query string (?key=) OR JSON body {"key": "..."}
  const providedKey = (((e || {}).parameter || {}).key) || (payload && payload.key);

  if (!apiKey || providedKey !== apiKey) {
    return jsonOutput_({ error: "Unauthorized" }, 401);
  }
  if (!payload) {
    return jsonOutput_({ error: "Invalid JSON" }, 400);
  }

  const sh = findResponseSheet_();

  // Validate headers (do not auto-modify response sheet)
  assertHeaders_(sh, REQUIRED_RESPONSE_HEADERS);

  // To append a row correctly even if header order differs,
  // we build the row as a full-width row and place values by header index.
  const headerRow = getHeaderRow_(sh);
  const idx = indexMap_(headerRow);
  const H = (name) => idx[normalizeHeader_(name)];

  const row = new Array(headerRow.length).fill("");

  row[H("Timestamp")] = new Date();
  row[H("CSM In charge")] = payload.csm || "";
  row[H("CLIENT")] = payload.client || "";
  row[H("Account name")] = payload.account || "";
  row[H("Number of attendees")] = Number(payload.attendees || 0);
  row[H("Main Contact Name")] = payload.mainContact || "";
  row[H("Date of Session")] = payload.sessionDate ? new Date(payload.sessionDate) : "";
  row[H("Duration of session (Minutes)")] = Number(payload.durationMinutes || 0);
  row[H("Upload Post-session brief (PDF)")] = payload.pdfUrl || "";
  row[H("Additional notes")] = payload.notes || "";

  sh.appendRow(row);

  return jsonOutput_({ ok: true });
}

/**
 * CORS preflight handler for browser-based requests (GitHub Pages).
 * Note: some Apps Script environments may not invoke doOptions; GET often still works.
 */
function doOptions(e) {
  return corsTextOutput_("");
}

/** -----------------------------
 * Summary builder
 * ----------------------------- */
function buildSummary_() {
  const respSheet = findResponseSheet_();
  const commitSheet = getOrCreateCommitmentsSheet_();

  // Responses sheet: validate headers (robust)
  assertHeaders_(respSheet, REQUIRED_RESPONSE_HEADERS);

  // Commitments sheet: ensure headers (safe to auto-create/append here)
  ensureHeaders_(commitSheet, COMMIT_HEADERS);

  const respValues = respSheet.getDataRange().getValues();
  const commitValues = commitSheet.getDataRange().getValues();

  const respHeader = respValues[0].map(v => (v ?? "").toString());
  const idx = indexMap_(respHeader);
  const H = (name) => idx[normalizeHeader_(name)];

  const commitHeader = commitValues[0].map(v => (v ?? "").toString());
  const cidx = indexMap_(commitHeader);
  const CH = (name) => cidx[normalizeHeader_(name)];

  // Commit map: key = client||account => committedMinutes
  const commitMap = new Map();
  for (let i = 1; i < commitValues.length; i++) {
    const client = (commitValues[i][CH("CLIENT")] || "").toString().trim();
    const account = (commitValues[i][CH("Account name")] || "").toString().trim();
    const committed = Number(commitValues[i][CH("Committed Minutes")] || 0);
    if (!client || !account) continue;
    commitMap.set(`${client}||${account}`, committed);
  }

  // Used minutes map from responses
  const usedMap = new Map();
  for (let i = 1; i < respValues.length; i++) {
    const client = (respValues[i][H("CLIENT")] || "").toString().trim();
    const account = (respValues[i][H("Account name")] || "").toString().trim();
    const dur = Number(respValues[i][H("Duration of session (Minutes)")] || 0);

    if (!client || !account) continue;

    const key = `${client}||${account}`;
    usedMap.set(key, (usedMap.get(key) || 0) + dur);
  }

  // Merge keys
  const keys = new Set([...commitMap.keys(), ...usedMap.keys()]);
  const rows = [];

  keys.forEach(key => {
    const parts = key.split("||");
    const client = parts[0] || "";
    const account = parts[1] || "";

    const committed = commitMap.get(key) || 0;
    const used = usedMap.get(key) || 0;
    const remaining = committed - used;

    const status =
      committed === 0 ? "NO COMMIT SET" :
      remaining < 0 ? "OVER LIMIT" :
      remaining <= NEAR_LIMIT_THRESHOLD_MIN ? "NEAR LIMIT" :
      "OK";

    rows.push({ client, account, committed, used, remaining, status });
  });

  // Sort by urgency (lowest remaining first)
  rows.sort((a, b) => a.remaining - b.remaining);

  return {
    generatedAt: new Date().toISOString(),
    thresholdNearLimitMinutes: NEAR_LIMIT_THRESHOLD_MIN,
    counts: {
      over: rows.filter(r => r.status === "OVER LIMIT").length,
      near: rows.filter(r => r.status === "NEAR LIMIT").length,
      ok: rows.filter(r => r.status === "OK").length,
      noCommit: rows.filter(r => r.status === "NO COMMIT SET").length
    },
    rows
  };
}

/** -----------------------------
 * Spreadsheet helpers
 * ----------------------------- */
function getSpreadsheet_() {
  // If bound script, active spreadsheet is best
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getName() === SPREADSHEET_NAME) return active;
  } catch (e) {}

  // Otherwise find by name in Drive
  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (!files.hasNext()) {
    throw new Error(`Spreadsheet not found by name: ${SPREADSHEET_NAME}`);
  }
  const file = files.next();
  return SpreadsheetApp.openById(file.getId());
}

function findResponseSheet_() {
  const ss = getSpreadsheet_();
  const sheets = ss.getSheets();

  // Find sheet whose header row contains all REQUIRED_RESPONSE_HEADERS (robust match)
  for (const sh of sheets) {
    const lastCol = sh.getLastColumn();
    const lastRow = sh.getLastRow();
    if (lastRow < 1 || lastCol < 1) continue;

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => (v ?? "").toString());
    if (containsAllNormalized_(header, REQUIRED_RESPONSE_HEADERS)) {
      return sh;
    }
  }

  // If not found, fallback to first sheet but validate (will throw if wrong)
  const fallback = sheets[0];
  assertHeaders_(fallback, REQUIRED_RESPONSE_HEADERS);
  return fallback;
}

function getOrCreateCommitmentsSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_COMMIT);
  if (!sh) sh = ss.insertSheet(SHEET_COMMIT);
  ensureHeaders_(sh, COMMIT_HEADERS);
  return sh;
}

/** -----------------------------
 * Robust header utilities
 * ----------------------------- */
function getHeaderRow_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => (v ?? "").toString());
}

function normalizeHeader_(s) {
  return (s ?? "")
    .toString()
    .trim()
    .replace(/\s+/g, " ") // collapse multiple spaces
    .toLowerCase();
}

function containsAllNormalized_(haveHeaders, requiredHeaders) {
  const have = new Set((haveHeaders || []).map(normalizeHeader_));
  return (requiredHeaders || []).every(r => have.has(normalizeHeader_(r)));
}

function assertHeaders_(sheet, requiredHeaders) {
  const headerRow = getHeaderRow_(sheet);
  const have = new Set(headerRow.map(normalizeHeader_));
  const missing = requiredHeaders.filter(h => !have.has(normalizeHeader_(h)));
  if (missing.length) {
    throw new Error(`Responses sheet missing headers: ${missing.join(", ")}`);
  }
}

function indexMap_(headerRow) {
  const map = {};
  (headerRow || []).forEach((h, i) => {
    map[normalizeHeader_(h)] = i;
  });
  return map;
}

/**
 * Only safe for Commitments sheet (not for responses sheet).
 * Creates headers if empty, appends missing headers if partially present.
 */
function ensureHeaders_(sheet, headers) {
  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => (v ?? "").toString());

  // If sheet empty, write headers
  const isEmpty = current.every(v => !v);
  if (isEmpty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  // Append missing headers
  const currentNorm = new Set(current.map(normalizeHeader_));
  const missing = headers.filter(h => !currentNorm.has(normalizeHeader_(h)));
  if (missing.length) {
    const start = current.length + 1;
    sheet.getRange(1, start, 1, missing.length).setValues([missing]);
  }
}

/** -----------------------------
 * Auth + output + CORS helpers
 * ----------------------------- */
function getApiKey_() {
  return PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
}

function setApiKey() {
  // Run once
  PropertiesService.getScriptProperties().setProperty(
    API_KEY_PROP,
    "zK7pQ2vH9cR1xT6mB8nL0sF5dG3yJ7uA1eW9qC2hV6tN8x"
  );
}

function jsonOutput_(obj, code) {
  // Apps Script doesn't reliably allow status codes; optionally include code in body
  if (code) obj = Object.assign({ httpStatus: code }, obj);
  return corsTextOutput_(JSON.stringify(obj), ContentService.MimeType.JSON);
}

function corsTextOutput_(text, mime) {
  const out = ContentService.createTextOutput(text);
  out.setMimeType(mime || ContentService.MimeType.TEXT);

  // CORS headers for GitHub Pages/browser fetch
  // If you get an error on setHeader, tell me the exact error and I’ll provide a fallback.
  out.setHeader("Access-Control-Allow-Origin", "*");
  out.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  out.setHeader("Access-Control-Allow-Headers", "Content-Type");

  return out;
}

function safeJson_(txt) {
  try { return JSON.parse(txt); } catch (e) { return null; }
}
