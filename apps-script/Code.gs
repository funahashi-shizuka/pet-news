const SHEET_NAME = "ニュース";
const IMPORTED_LABEL = "pet-news-imported";
const GMAIL_QUERY =
  'from:(googlealerts-noreply@google.com) newer_than:30d -label:"pet-news-imported"';
const HEADERS = ["受信日", "キーワード", "タイトル", "媒体", "URL", "概要"];

function setup() {
  const sheet = getSheet_();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm");
  }

  if (!GmailApp.getUserLabelByName(IMPORTED_LABEL)) {
    GmailApp.createLabel(IMPORTED_LABEL);
  }
}

function importGoogleAlerts() {
  setup();

  const sheet = getSheet_();
  const existingUrls = getExistingUrls_(sheet);
  const label = GmailApp.getUserLabelByName(IMPORTED_LABEL);
  const threads = GmailApp.search(GMAIL_QUERY, 0, 50);
  const rows = [];

  threads.forEach((thread) => {
    thread.getMessages().forEach((message) => {
      parseAlertMessage_(message).forEach((item) => {
        if (!existingUrls.has(item.url)) {
          rows.push([
            item.date,
            item.keyword,
            item.title,
            item.source,
            item.url,
            item.summary,
          ]);
          existingUrls.add(item.url);
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

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function getExistingUrls_(sheet) {
  if (sheet.getLastRow() < 2) return new Set();
  return new Set(
    sheet
      .getRange(2, 5, sheet.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .filter(Boolean),
  );
}

function parseAlertMessage_(message) {
  const date = message.getDate();
  const keyword = parseKeyword_(message.getSubject());
  const html = message.getBody();
  const plain = normalizeText_(message.getPlainBody());
  const candidates = extractLinks_(html);

  return candidates.map((candidate) => {
    const url = normalizeUrl_(candidate.url);
    const title = normalizeText_(candidate.title) || "(タイトルなし)";
    const snippet = findSnippet_(plain, title);

    return {
      date,
      keyword,
      title,
      source: sourceFromUrl_(url),
      url,
      summary: snippet,
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
