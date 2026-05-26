const CONFIG = {
  sheetName: "alerts",
  importedLabel: "industry-info-imported",
  gmailQuery:
    'from:(googlealerts-noreply@google.com) newer_than:30d -label:"industry-info-imported"',
  maxThreads: 50,
};

const HEADERS = [
  "id",
  "receivedAt",
  "alertKeyword",
  "title",
  "source",
  "url",
  "snippet",
  "status",
  "priority",
  "owner",
  "category",
  "tags",
  "memo",
  "publishedAt",
  "importedAt",
];

function setup() {
  const sheet = getOrCreateSheet_();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange("B:B").setNumberFormat("yyyy-mm-dd hh:mm");
    sheet.getRange("N:N").setNumberFormat("yyyy-mm-dd");
    sheet.getRange("O:O").setNumberFormat("yyyy-mm-dd hh:mm");
  }

  if (!GmailApp.getUserLabelByName(CONFIG.importedLabel)) {
    GmailApp.createLabel(CONFIG.importedLabel);
  }
}

function importGoogleAlerts() {
  setup();

  const sheet = getOrCreateSheet_();
  const importedIds = getImportedIds_(sheet);
  const label = GmailApp.getUserLabelByName(CONFIG.importedLabel);
  const threads = GmailApp.search(CONFIG.gmailQuery, 0, CONFIG.maxThreads);
  const rows = [];

  threads.forEach((thread) => {
    thread.getMessages().forEach((message) => {
      parseAlertMessage_(message).forEach((item) => {
        if (!importedIds.has(item.id)) {
          rows.push(HEADERS.map((header) => item[header] || ""));
          importedIds.add(item.id);
        }
      });
    });
    thread.addLabel(label);
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
  }

  return { imported: rows.length, checkedThreads: threads.length };
}

function createHourlyTrigger() {
  ScriptApp.newTrigger("importGoogleAlerts").timeBased().everyHours(1).create();
}

function doGet(e) {
  if (e.parameter.action === "list") {
    return json_(readRecords_());
  }

  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("業界情報チェック")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || "{}");
  const result = updateRecord_(payload.id, payload.fields || {});
  return json_(result);
}

function listRecords() {
  return readRecords_();
}

function saveRecordFields(id, fields) {
  return updateRecord_(id, fields);
}

function runImportNow() {
  return importGoogleAlerts();
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.sheetName) || ss.insertSheet(CONFIG.sheetName);
}

function getImportedIds_(sheet) {
  if (sheet.getLastRow() < 2) return new Set();
  return new Set(
    sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .filter(Boolean),
  );
}

function parseAlertMessage_(message) {
  const receivedAt = message.getDate();
  const alertKeyword = parseKeyword_(message.getSubject());
  const html = message.getBody();
  const plain = normalizeText_(message.getPlainBody());
  const candidates = extractLinks_(html);

  return candidates.map((candidate) => {
    const url = normalizeUrl_(candidate.url);
    const title = normalizeText_(candidate.title) || "(タイトルなし)";
    const snippet = findSnippet_(plain, title);

    return {
      id: makeId_(alertKeyword, url),
      receivedAt,
      alertKeyword,
      title,
      source: sourceFromUrl_(url),
      url,
      snippet,
      status: "未確認",
      priority: "中",
      owner: "未設定",
      category: "未分類",
      tags: "",
      memo: "",
      publishedAt: "",
      importedAt: new Date(),
    };
  });
}

function parseKeyword_(subject) {
  return normalizeText_(subject)
    .replace(/^Google\s*アラート\s*[-:：]\s*/i, "")
    .replace(/^Google\s*Alert\s*[-:：]\s*/i, "")
    .trim();
}

function extractLinks_(html) {
  const links = [];
  const seen = new Set();
  const anchorRegex = /<a\b[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let match;

  while ((match = anchorRegex.exec(html)) !== null) {
    const url = normalizeUrl_(decodeHtml_(match[1]));
    const title = normalizeText_(stripTags_(decodeHtml_(match[2])));

    if (!isArticleUrl_(url) || !title || title.length < 6 || seen.has(url)) continue;

    links.push({ url, title });
    seen.add(url);
  }

  return links;
}

function normalizeUrl_(url) {
  const decoded = decodeHtml_(url || "");
  const match = decoded.match(/[?&](?:url|q)=([^&]+)/);
  if (match) return decodeURIComponent(match[1]);
  return decoded;
}

function isArticleUrl_(url) {
  if (!/^https?:\/\//i.test(url)) return false;
  return ![
    "google.com/alerts",
    "google.co.jp/alerts",
    "accounts.google",
    "support.google",
    "policies.google",
    "mail.google",
  ].some((blocked) => url.includes(blocked));
}

function findSnippet_(plainBody, title) {
  const normalizedTitle = normalizeText_(title);
  const index = plainBody.indexOf(normalizedTitle);
  if (index < 0) return "";
  return normalizeText_(plainBody.slice(index + normalizedTitle.length, index + normalizedTitle.length + 240));
}

function readRecords_() {
  const sheet = getOrCreateSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();

  return values
    .filter((row) => row.some(Boolean))
    .map((row) => {
      const record = {};
      headers.forEach((header, index) => {
        const value = row[index];
        record[header] = value instanceof Date ? value.toISOString() : value;
      });
      return record;
    });
}

function updateRecord_(id, fields) {
  if (!id) return { ok: false, error: "Missing id" };

  const sheet = getOrCreateSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idColumn = headers.indexOf("id") + 1;
  if (idColumn < 1) return { ok: false, error: "Missing id column" };

  for (let row = 2; row <= values.length; row += 1) {
    if (values[row - 1][idColumn - 1] === id) {
      Object.keys(fields).forEach((key) => {
        const column = headers.indexOf(key) + 1;
        if (column > 0) sheet.getRange(row, column).setValue(fields[key]);
      });
      return { ok: true, id };
    }
  }

  return { ok: false, error: "Record not found" };
}

function makeId_(keyword, url) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, `${keyword}|${url}`);
  return Utilities.base64EncodeWebSafe(bytes).slice(0, 22);
}

function sourceFromUrl_(url) {
  const match = String(url || "").match(/^https?:\/\/([^/?#]+)/i);
  return match ? match[1].replace(/^www\./, "") : "";
}

function stripTags_(value) {
  return String(value || "").replace(/<[^>]*>/g, " ");
}

function decodeHtml_(value) {
  return String(value || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function normalizeText_(value) {
  return stripTags_(decodeHtml_(value))
    .replace(/\s+/g, " ")
    .trim();
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
