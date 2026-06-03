const PEOPLE_SOURCE_SPREADSHEET_ID = "1k79vEPkX1glnz5AQAROdLu_3e-FodeGRpSFMqw9JS_E";
const PEOPLE_SOURCE_GID = 1321198903;
const PEOPLE_STATE_SHEET_NAME = "人員資料庫";

const PEOPLE_HEADERS = [
  "id",
  "archived",
  "category",
  "store",
  "name",
  "englishName",
  "reportDate",
  "promotionMonth",
  "uniformSize",
  "leaveMonth",
  "expectedReturnDate",
  "promotedPassed",
  "innischool",
  "uniformReturned",
  "note",
  "source",
  "updatedAt",
];

function doGet(e) {
  const params = e.parameter || {};
  const callback = params.callback || "callback";
  let payload;

  try {
    const action = params.action || "list";
    if (action === "list") payload = { ok: true, people: listPeople_() };
    else if (action === "archive") payload = { ok: true, person: archivePerson_(params.id) };
    else if (action === "update") payload = { ok: true, person: updatePerson_(params.id, JSON.parse(params.patch || "{}")) };
    else if (action === "upsert") payload = { ok: true, person: upsertPerson_(JSON.parse(params.person || "{}")) };
    else payload = { ok: false, message: "Unknown action: " + action };
  } catch (error) {
    payload = { ok: false, message: String(error && error.message ? error.message : error) };
  }

  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(payload) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function listPeople_() {
  const sourcePeople = readSourcePeople_();
  const state = readStateMap_();
  const sourceIds = {};
  const merged = [];

  sourcePeople.forEach(person => {
    sourceIds[person.id] = true;
    const saved = state[person.id] || {};
    const mergedPerson = state[person.id] ? Object.assign({}, person, cleanState_(saved)) : person;
    if (!toBoolean_(mergedPerson.archived)) merged.push(mergedPerson);
  });

  Object.keys(state).forEach(id => {
    if (sourceIds[id]) return;
    const person = cleanState_(state[id]);
    if (!toBoolean_(person.archived)) merged.push(person);
  });

  return merged;
}

function archivePerson_(id) {
  if (!id) throw new Error("Missing id");
  return upsertPerson_(Object.assign(getPersonState_(id), {
    id: id,
    archived: true,
    updatedAt: new Date().toISOString(),
  }));
}

function updatePerson_(id, patch) {
  if (!id) throw new Error("Missing id");
  const current = getPersonState_(id);
  return upsertPerson_(Object.assign({}, current, patch, {
    id: id,
    updatedAt: new Date().toISOString(),
  }));
}

function upsertPerson_(person) {
  if (!person.id) person.id = "manual-" + Date.now();
  person.updatedAt = person.updatedAt || new Date().toISOString();

  const sheet = getStateSheet_();
  const values = sheet.getDataRange().getValues();
  const index = values.findIndex((row, rowIndex) => rowIndex > 0 && row[0] === person.id);
  const rowValues = PEOPLE_HEADERS.map(key => serializeValue_(person[key]));

  if (index >= 0) sheet.getRange(index + 1, 1, 1, PEOPLE_HEADERS.length).setValues([rowValues]);
  else sheet.appendRow(rowValues);

  return person;
}

function getPersonState_(id) {
  return readStateMap_()[id] || { id: id };
}

function readStateMap_() {
  const sheet = getStateSheet_();
  const values = sheet.getDataRange().getValues();
  const map = {};
  values.slice(1).forEach(row => {
    const person = {};
    PEOPLE_HEADERS.forEach((header, index) => person[header] = row[index]);
    if (person.id) map[person.id] = person;
  });
  return map;
}

function getStateSheet_() {
  const ss = SpreadsheetApp.openById(PEOPLE_SOURCE_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PEOPLE_STATE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PEOPLE_STATE_SHEET_NAME);
    sheet.getRange(1, 1, 1, PEOPLE_HEADERS.length).setValues([PEOPLE_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readSourcePeople_() {
  const ss = SpreadsheetApp.openById(PEOPLE_SOURCE_SPREADSHEET_ID);
  const sheet = getSheetByGid_(ss, PEOPLE_SOURCE_GID);
  const values = sheet.getDataRange().getDisplayValues();
  return values.slice(1).map((row, index) => mapSourceRow_(row, index + 2))
    .filter(person => person.store || person.reportDate || (person.name && person.name !== "未填姓名"));
}

function getSheetByGid_(ss, gid) {
  const sheets = ss.getSheets();
  const sheet = sheets.find(item => item.getSheetId() === Number(gid));
  if (!sheet) throw new Error("找不到人員來源分頁");
  return sheet;
}

function mapSourceRow_(row, rowNumber) {
  const rawCategory = String(row[0] || "").trim();
  const store = String(row[1] || "").trim();
  const reportDate = parseDate_(row[2]);
  const chineseName = String(row[3] || "").trim();
  const englishName = String(row[4] || "").trim();
  const promotedPassed = parseBoolean_(row[7]) || Boolean(String(row[6] || "").trim());
  let category = "new";
  if (rawCategory.indexOf("工讀") >= 0) category = "parttime";
  if (promotedPassed && category !== "parttime") category = "promoted";

  const noteParts = [
    rawCategory ? "原類別：" + rawCategory : "",
    row[8] ? "新生課：" + row[8] : "",
    row[10] ? "資深認證：" + row[10] : "",
    row[12] ? "備註：" + row[12] : "",
  ].filter(Boolean);

  return {
    id: "people-sheet-" + rowNumber,
    source: "peopleSheet",
    archived: false,
    category: category,
    store: store,
    name: chineseName || englishName || "未填姓名",
    englishName: englishName,
    reportDate: reportDate,
    promotionMonth: parseMonth_(row[5]) || parseMonth_(row[6]),
    uniformSize: String(row[11] || "").trim(),
    leaveMonth: "",
    expectedReturnDate: "",
    promotedPassed: promotedPassed,
    innischool: parseBoolean_(row[9]),
    uniformReturned: false,
    note: noteParts.join("；"),
  };
}

function cleanState_(person) {
  const cleaned = {};
  PEOPLE_HEADERS.forEach(key => cleaned[key] = normalizeValue_(key, person[key]));
  return cleaned;
}

function normalizeValue_(key, value) {
  if (["archived", "promotedPassed", "innischool", "uniformReturned"].includes(key)) return toBoolean_(value);
  return value == null ? "" : value;
}

function serializeValue_(value) {
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return value == null ? "" : value;
}

function toBoolean_(value) {
  const text = String(value || "").trim().toLowerCase();
  return ["true", "v", "yes", "y", "1", "已開通", "已通過"].includes(text);
}

function parseBoolean_(value) {
  return toBoolean_(value);
}

function parseMonth_(value) {
  const text = String(value || "").trim();
  const match = text.match(/(20\d{2})\D+(\d{1,2})/);
  if (match) return match[1] + "-" + String(Number(match[2])).padStart(2, "0");
  return "";
}

function parseDate_(value) {
  const text = String(value || "").trim();
  const match = text.match(/(20\d{2})\D+(\d{1,2})\D+(\d{1,2})/);
  if (!match) return "";
  return match[1] + "-" + String(Number(match[2])).padStart(2, "0") + "-" + String(Number(match[3])).padStart(2, "0");
}
