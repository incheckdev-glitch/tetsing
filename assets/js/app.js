/**
 * InCheck Lite Session Tracker API (2-tab fixed setup)
 *
 * Tabs expected:
 *  1) RESPONSES_SHEET_NAME: Timestamp | CSM In charge | CLIENT | Account name | Number of attendees |
 *                           Main Contact Name | Date of Session | Duration of session (Minutes) |
 *                           Upload Post-session brief (PDF) | Additional notes
 *
 *  2) COMMIT_SHEET_NAME: CLIENT | Account name | Committed Minutes
 *
 * Endpoints:
 *  GET     /exec?action=summary
 *  GET     /exec?action=health
 *  POST    /exec?key=YOUR_KEY   (JSON body)
 *  OPTIONS /exec                (CORS preflight)
 */

const SPREADSHEET_NAME = "InCheck Lite Session Tracker";

// ✅ Set these exactly to your tab names (bottom tabs in the sheet)
const RESPONSES_SHEET_NAME = "InCheck Lite Session Tracker"; // <-- CHANGE to your first tab exact name
const COMMIT_SHEET_NAME = "Commitments";         // <-- CHANGE to your second tab exact name

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

const COMMIT_HEADERS = ["CLIENT", "Account name", "Committed Minutes"];

const API_KEY_PROP = "DASH_API_KEY";
const NEAR_LIMIT_THRESHOLD_MIN = 120;

/** -----------------------------
 * HTTP handlers
 * ----------------------------- */
function doGet(e) {
  const action = (((e || {}).parameter || {}).action || "summary").toLowerCase();

  if (action === "summary") return jsonOutput_(buildSummary_());

  if (action === "health") {
    const ss = getSpreadsheet_();
    const resp = getSheet_(RESPONSES_SHEET_NAME);
    const com = getSheet_(COMMIT_SHEET_NAME);

    assertHeadersExact_(resp, REQUIRED_RESPONSE_HEADERS);
    assertHeadersExact_(com, COMMIT_HEADERS);

    return jsonOutput_({
      ok: true,
      spreadsheet: ss.getName(),
      responsesSheet: resp.getName(),
      commitmentsSheet: com.getName(),
      rowsInResponses: Math.max(0, resp.getLastRow() - 1),
      rowsInCommitments: Math.max(0, com.getLastRow() - 1)
    });
  }

  return jsonOutput_({ error: "Unknown action" }, 400);
}

function doPost(e) {
  const apiKey = getApiKey_();
  const payload = safeJson_((e && e.postData && e.postData.contents) || "");

  const providedKey = (((e || {}).parameter || {}).key) || (payload && payload.key);

  if (!apiKey || providedKey !== apiKey) return jsonOutput_({ error: "Unauthorized" }, 401);
  if (!payload) return jsonOutput_({ error: "Invalid JSON" }, 400);

  const sh = getSheet_(RESPONSES_SHEET_NAME);
  assertHeadersExact_(sh, REQUIRED_RESPONSE_HEADERS);

  // Append aligned with headers
  const row = [
    new Date(),                             // Timestamp
    payload.csm || "",                      // CSM In charge
    payload.client || "",                   // CLIENT
    payload.account || "",                  // Account name
    Number(payload.attendees || 0),         // Number of attendees
    payload.mainContact || "",              // Main Contact Name
    payload.sessionDate ? new Date(payload.sessionDate) : "", // Date of Session
    Number(payload.durationMinutes || 0),   // Duration of session (Minutes)
    payload.pdfUrl || "",                   // Upload Post-session brief (PDF) -> store link if any
    payload.notes || ""                     // Additional notes
  ];

  sh.appendRow(row);
  return jsonOutput_({ ok: true });
}

function doOptions(e) {
  return corsTextOutput_("");
}

/** -----------------------------
 * Summary builder
 * ----------------------------- */
function buildSummary_() {
  const respSheet = getSheet_(RESPONSES_SHEET_NAME);
  const commitSheet = getSheet_(COMMIT_SHEET_NAME);

  assertHeadersExact_(respSheet, REQUIRED_RESPONSE_HEADERS);
  assertHeadersExact_(commitSheet, COMMIT_HEADERS);

  const resp = respSheet.getDataRange().getValues();
  const com = commitSheet.getDataRange().getValues();

  // Indexes based on exact headers
  // Responses: A Timestamp, B CSM, C CLIENT, D Account, E Attendees, F Main Contact, G Date, H Duration, I PDF, J Notes
  const IDX_CLIENT = 2;
  const IDX_ACCOUNT = 3;
  const IDX_DURATION = 7;

  // Commitments: A CLIENT, B Account name, C Committed Minutes
  const C_CLIENT = 0;
  const C_ACCOUNT = 1;
  const C_COMMITTED = 2;

  const commitMap = new Map();
  for (let i = 1; i < com.length; i++) {
    const client = (com[i][C_CLIENT] || "").toString().trim();
    const account = (com[i][C_ACCOUNT] || "").toString().trim();
    const committed = Number(com[i][C_COMMITTED] || 0);
    if (!client || !account) continue;
    commitMap.set(`${client}||${account}`, committed);
  }

  const usedMap = new Map();
  for (let i = 1; i < resp.length; i++) {
    const client = (resp[i][IDX_CLIENT] || "").toString().trim();
    const account = (resp[i][IDX_ACCOUNT] || "").toString().trim();
    const dur = Number(resp[i][IDX_DURATION] || 0);

    if (!client || !account) continue;

    const key = `${client}||${account}`;
    usedMap.set(key, (usedMap.get(key) || 0) + dur);
  }

  const keys = new Set([...commitMap.keys(), ...usedMap.keys()]);
  const rows = [];

  keys.forEach(key => {
    const [client, account] = key.split("||");
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
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getName() === SPREADSHEET_NAME) return active;
  } catch (e) {}

  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (!files.hasNext()) throw new Error(`Spreadsheet not found by name: ${SPREADSHEET_NAME}`);
  return SpreadsheetApp.openById(files.next().getId());
}

function getSheet_(name) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet tab not found: "${name}"`);
  return sh;
}

/**
 * Exact header check (since you know your two tabs).
 * If any mismatch happens, it shows the actual headers for quick debugging.
 */
function assertHeadersExact_(sheet, expected) {
  const lastCol = Math.max(sheet.getLastColumn(), expected.length);
  const actual = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => (v ?? "").toString().trim());

  // Compare only up to expected length
  const actualShort = actual.slice(0, expected.length);

  const ok = expected.every((h, i) => (actualShort[i] || "") === h);
  if (!ok) {
    throw new Error(
      `Header mismatch in sheet "${sheet.getName()}".\n` +
      `Expected: ${expected.join(" | ")}\n` +
      `Actual:   ${actualShort.join(" | ")}`
    );
  }
}

/** -----------------------------
 * Auth + output + CORS helpers
 * ----------------------------- */
function getApiKey_() {
  return PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
}

function setApiKey() {
  PropertiesService.getScriptProperties().setProperty(
    API_KEY_PROP,
    "zK7pQ2vH9cR1xT6mB8nL0sF5dG3yJ7uA1eW9qC2hV6tN8x"
  );
}

function jsonOutput_(obj, code) {
  if (code) obj = Object.assign({ httpStatus: code }, obj);
  return corsTextOutput_(JSON.stringify(obj), ContentService.MimeType.JSON);
}

function corsTextOutput_(text, mime) {
  const out = ContentService.createTextOutput(text);
  out.setMimeType(mime || ContentService.MimeType.TEXT);
  out.setHeader("Access-Control-Allow-Origin", "*");
  out.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  out.setHeader("Access-Control-Allow-Headers", "Content-Type");
  return out;
}

function safeJson_(txt) {
  try { return JSON.parse(txt); } catch (e) { return null; }
}
