const sampleAlerts = [
  {
    id: "sample-001",
    receivedAt: "2026-05-25T09:20:00+09:00",
    alertKeyword: "ペット保険 市場",
    title: "ペット関連サービス、定期利用型の加入動向に変化",
    source: "業界ニュースサンプル",
    url: "https://example.com/pet-market-subscription",
    snippet:
      "ペット関連サービスでは、継続利用を前提にした商品の比較検討が広がっている。価格だけでなく、補償範囲や手続きの簡便さが選定要素になっている。",
    status: "未確認",
    priority: "高",
    owner: "佐藤",
    category: "市場動向",
    tags: "市場,加入動向",
    memo: "",
    publishedAt: "2026-05-25",
    importedAt: "2026-05-25T10:00:00+09:00",
  },
  {
    id: "sample-002",
    receivedAt: "2026-05-24T08:45:00+09:00",
    alertKeyword: "動物医療 DX",
    title: "動物病院向けオンライン受付、都市部で導入が進む",
    source: "ヘルスケアビジネスサンプル",
    url: "https://example.com/vet-dx-checkin",
    snippet:
      "オンライン受付や事前問診の導入により、来院前の情報整理と待ち時間短縮を目指す動きが見られる。",
    status: "確認中",
    priority: "中",
    owner: "田中",
    category: "医療DX",
    tags: "動物病院,DX",
    memo: "既存顧客向けサービス連携の参考にする。",
    publishedAt: "2026-05-24",
    importedAt: "2026-05-24T09:30:00+09:00",
  },
  {
    id: "sample-003",
    receivedAt: "2026-05-23T12:10:00+09:00",
    alertKeyword: "保険 金融庁",
    title: "少額短期保険分野の情報開示に関する資料が更新",
    source: "行政情報サンプル",
    url: "https://example.com/insurance-disclosure-update",
    snippet:
      "保険関連事業者の情報開示や顧客説明に関する資料が更新された。開示項目の確認が必要。",
    status: "共有済み",
    priority: "高",
    owner: "鈴木",
    category: "規制",
    tags: "規制,開示",
    memo: "週次会議で共有済み。",
    publishedAt: "2026-05-23",
    importedAt: "2026-05-23T13:00:00+09:00",
  },
  {
    id: "sample-004",
    receivedAt: "2026-05-21T17:35:00+09:00",
    alertKeyword: "ペット 高齢化",
    title: "高齢ペット向けケアサービス、相談窓口の需要が増加",
    source: "生活産業サンプル",
    url: "https://example.com/senior-pet-care",
    snippet:
      "高齢ペットの通院、食事、介護に関する相談が増えている。保険加入後のサポート領域にも示唆がある。",
    status: "未確認",
    priority: "中",
    owner: "未設定",
    category: "顧客ニーズ",
    tags: "高齢化,ケア",
    memo: "",
    publishedAt: "2026-05-21",
    importedAt: "2026-05-21T18:10:00+09:00",
  },
  {
    id: "sample-005",
    receivedAt: "2026-05-18T10:05:00+09:00",
    alertKeyword: "ペット保険 比較",
    title: "ペット保険比較サイト、補償条件の見せ方を刷新",
    source: "マーケティングサンプル",
    url: "https://example.com/pet-insurance-comparison-ui",
    snippet:
      "比較検討時に重視される項目を再整理し、免責条件や通院補償を分かりやすく表示する改修が進んでいる。",
    status: "確認中",
    priority: "低",
    owner: "佐藤",
    category: "競合調査",
    tags: "比較,UI",
    memo: "",
    publishedAt: "2026-05-18",
    importedAt: "2026-05-18T10:40:00+09:00",
  },
  {
    id: "sample-006",
    receivedAt: "2026-05-17T14:50:00+09:00",
    alertKeyword: "ペット テック",
    title: "見守りデバイスの利用データ、健康管理サービスへ展開",
    source: "テック市場サンプル",
    url: "https://example.com/pet-tech-health-data",
    snippet:
      "活動量や食事量などのデータをもとに、日常的な体調変化の把握を支援するサービスが広がる。",
    status: "保留",
    priority: "低",
    owner: "田中",
    category: "周辺サービス",
    tags: "データ,テック",
    memo: "短期より中長期の参考。",
    publishedAt: "2026-05-17",
    importedAt: "2026-05-17T15:20:00+09:00",
  },
];

const state = {
  alerts: [],
  selectedId: null,
  filters: {
    search: "",
    status: "all",
    priority: "all",
    category: "all",
    owner: "all",
  },
};

const els = {
  csvUrlInput: document.querySelector("#csvUrlInput"),
  loadCsvButton: document.querySelector("#loadCsvButton"),
  exportButton: document.querySelector("#exportButton"),
  searchInput: document.querySelector("#searchInput"),
  statusFilters: document.querySelector("#statusFilters"),
  priorityFilter: document.querySelector("#priorityFilter"),
  categoryFilter: document.querySelector("#categoryFilter"),
  ownerFilter: document.querySelector("#ownerFilter"),
  rows: document.querySelector("#alertRows"),
  resultCount: document.querySelector("#resultCount"),
  metricTotal: document.querySelector("#metricTotal"),
  metricOpen: document.querySelector("#metricOpen"),
  metricHigh: document.querySelector("#metricHigh"),
  metricWeek: document.querySelector("#metricWeek"),
  detailEmpty: document.querySelector("#detailEmpty"),
  detailPanel: document.querySelector("#detailPanel"),
  detailKeyword: document.querySelector("#detailKeyword"),
  detailTitle: document.querySelector("#detailTitle"),
  detailLink: document.querySelector("#detailLink"),
  detailSource: document.querySelector("#detailSource"),
  detailReceived: document.querySelector("#detailReceived"),
  detailCategory: document.querySelector("#detailCategory"),
  detailSnippet: document.querySelector("#detailSnippet"),
  detailPriority: document.querySelector("#detailPriority"),
  detailStatus: document.querySelector("#detailStatus"),
  detailOwner: document.querySelector("#detailOwner"),
  detailTags: document.querySelector("#detailTags"),
  detailMemo: document.querySelector("#detailMemo"),
};

const storageKey = "industry-info-hub-alerts";

init();

function init() {
  const saved = localStorage.getItem(storageKey);
  state.alerts = saved ? JSON.parse(saved) : structuredClone(sampleAlerts);
  state.selectedId = state.alerts[0]?.id ?? null;

  els.searchInput.addEventListener("input", () => {
    state.filters.search = els.searchInput.value.trim().toLowerCase();
    render();
  });

  els.statusFilters.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-status]");
    if (!button) return;
    state.filters.status = button.dataset.status;
    [...els.statusFilters.querySelectorAll("button")].forEach((item) => {
      item.classList.toggle("active", item === button);
    });
    render();
  });

  els.priorityFilter.addEventListener("change", () => {
    state.filters.priority = els.priorityFilter.value;
    render();
  });

  els.categoryFilter.addEventListener("change", () => {
    state.filters.category = els.categoryFilter.value;
    render();
  });

  els.ownerFilter.addEventListener("change", () => {
    state.filters.owner = els.ownerFilter.value;
    render();
  });

  els.loadCsvButton.addEventListener("click", loadCsvFromUrl);
  els.exportButton.addEventListener("click", exportCsv);

  [els.detailPriority, els.detailStatus, els.detailOwner, els.detailTags].forEach((input) => {
    input.addEventListener("change", updateSelectedFromDetail);
  });
  els.detailMemo.addEventListener("input", debounce(updateSelectedFromDetail, 300));

  render();
}

function render() {
  renderFilterOptions();
  renderMetrics();
  renderRows();
  renderDetail();
}

function renderFilterOptions() {
  renderSelectOptions(els.categoryFilter, uniqueValues("category"), state.filters.category);
  renderSelectOptions(els.ownerFilter, uniqueValues("owner"), state.filters.owner);
}

function renderSelectOptions(select, values, selectedValue) {
  const current = values.includes(selectedValue) ? selectedValue : "all";
  select.innerHTML = [
    `<option value="all">すべて</option>`,
    ...values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`),
  ].join("");
  select.value = current;
}

function uniqueValues(key) {
  return [...new Set(state.alerts.map((item) => item[key]).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, "ja"),
  );
}

function renderMetrics() {
  const now = new Date();
  const weekAgo = new Date(now);
  weekAgo.setDate(now.getDate() - 7);

  els.metricTotal.textContent = state.alerts.length;
  els.metricOpen.textContent = state.alerts.filter((item) => item.status === "未確認").length;
  els.metricHigh.textContent = state.alerts.filter((item) => item.priority === "高").length;
  els.metricWeek.textContent = state.alerts.filter((item) => new Date(item.receivedAt) >= weekAgo).length;
}

function renderRows() {
  const rows = filteredAlerts();
  els.resultCount.textContent = `${rows.length}件`;

  if (!rows.length) {
    els.rows.innerHTML = `<tr><td colspan="7">該当する記事がありません</td></tr>`;
    return;
  }

  els.rows.innerHTML = rows
    .map((item) => {
      const priorityClass = item.priority === "高" ? "high" : item.priority === "中" ? "medium" : "low";
      return `
        <tr data-id="${escapeHtml(item.id)}" class="${item.id === state.selectedId ? "selected" : ""}">
          <td>${formatDate(item.receivedAt)}</td>
          <td>${escapeHtml(item.alertKeyword)}</td>
          <td>
            <div class="title-cell">
              <strong>${escapeHtml(item.title)}</strong>
              <span>${escapeHtml(item.snippet)}</span>
            </div>
          </td>
          <td>${escapeHtml(item.source)}</td>
          <td><span class="pill ${priorityClass}">${escapeHtml(item.priority)}</span></td>
          <td><span class="pill">${escapeHtml(item.status)}</span></td>
          <td>${escapeHtml(item.owner || "未設定")}</td>
        </tr>
      `;
    })
    .join("");

  els.rows.querySelectorAll("tr[data-id]").forEach((row) => {
    row.addEventListener("click", () => {
      state.selectedId = row.dataset.id;
      renderRows();
      renderDetail();
    });
  });
}

function renderDetail() {
  const item = state.alerts.find((alert) => alert.id === state.selectedId);
  els.detailPanel.classList.toggle("hidden", !item);
  els.detailEmpty.classList.toggle("hidden", Boolean(item));
  if (!item) return;

  els.detailKeyword.textContent = item.alertKeyword;
  els.detailTitle.textContent = item.title;
  els.detailLink.href = item.url;
  els.detailSource.textContent = item.source;
  els.detailReceived.textContent = formatDateTime(item.receivedAt);
  els.detailCategory.textContent = item.category || "未分類";
  els.detailSnippet.textContent = item.snippet;
  els.detailPriority.value = item.priority || "中";
  els.detailStatus.value = item.status || "未確認";
  els.detailOwner.value = item.owner || "";
  els.detailTags.value = item.tags || "";
  els.detailMemo.value = item.memo || "";
}

function filteredAlerts() {
  const query = state.filters.search;
  return state.alerts
    .filter((item) => {
      if (state.filters.status !== "all" && item.status !== state.filters.status) return false;
      if (state.filters.priority !== "all" && item.priority !== state.filters.priority) return false;
      if (state.filters.category !== "all" && item.category !== state.filters.category) return false;
      if (state.filters.owner !== "all" && item.owner !== state.filters.owner) return false;
      if (!query) return true;
      return [
        item.alertKeyword,
        item.title,
        item.source,
        item.snippet,
        item.status,
        item.priority,
        item.owner,
        item.category,
        item.tags,
        item.memo,
      ]
        .join(" ")
        .toLowerCase()
        .includes(query);
    })
    .sort((a, b) => new Date(b.receivedAt) - new Date(a.receivedAt));
}

function updateSelectedFromDetail() {
  const index = state.alerts.findIndex((alert) => alert.id === state.selectedId);
  if (index < 0) return;

  state.alerts[index] = {
    ...state.alerts[index],
    priority: els.detailPriority.value,
    status: els.detailStatus.value,
    owner: els.detailOwner.value.trim() || "未設定",
    tags: els.detailTags.value.trim(),
    memo: els.detailMemo.value.trim(),
  };

  save();
  renderMetrics();
  renderRows();
}

async function loadCsvFromUrl() {
  const url = els.csvUrlInput.value.trim();
  if (!url) return;

  els.loadCsvButton.disabled = true;
  els.loadCsvButton.textContent = "読込中";

  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();
    const records = csvToRecords(text);
    if (!records.length) throw new Error("No records");

    state.alerts = normalizeRecords(records);
    state.selectedId = state.alerts[0]?.id ?? null;
    save();
    render();
  } catch (error) {
    alert("CSVを読み込めませんでした。公開CSVのURLか、ネットワーク設定を確認してください。");
    console.error(error);
  } finally {
    els.loadCsvButton.disabled = false;
    els.loadCsvButton.textContent = "読み込み";
  }
}

function csvToRecords(text) {
  const rows = parseCsv(text);
  const headers = rows.shift()?.map((header) => header.trim()) ?? [];
  return rows
    .filter((row) => row.some((value) => value.trim()))
    .map((row) =>
      Object.fromEntries(headers.map((header, index) => [header, row[index]?.trim() ?? ""])),
    );
}

function normalizeRecords(records) {
  return records.map((record, index) => ({
    id: record.id || record.ID || `row-${index + 1}`,
    receivedAt: record.receivedAt || record["受信日"] || new Date().toISOString(),
    alertKeyword: record.alertKeyword || record["キーワード"] || "",
    title: record.title || record["タイトル"] || "(タイトルなし)",
    source: record.source || record["媒体"] || hostname(record.url || record.URL || ""),
    url: record.url || record.URL || "#",
    snippet: record.snippet || record["概要"] || "",
    status: record.status || record["ステータス"] || "未確認",
    priority: record.priority || record["重要度"] || "中",
    owner: record.owner || record["担当"] || "未設定",
    category: record.category || record["カテゴリ"] || "未分類",
    tags: record.tags || record["タグ"] || "",
    memo: record.memo || record["メモ"] || "",
    publishedAt: record.publishedAt || record["公開日"] || "",
    importedAt: record.importedAt || record["取込日"] || "",
  }));
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"' && inQuotes && next === '"') {
      value += '"';
      i += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === "," && !inQuotes) {
      row.push(value);
      value = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
    } else {
      value += char;
    }
  }

  row.push(value);
  rows.push(row);
  return rows;
}

function exportCsv() {
  const headers = [
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
  const rows = [headers, ...state.alerts.map((item) => headers.map((key) => item[key] ?? ""))];
  const csv = rows.map((row) => row.map(csvCell).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `industry-alerts-${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function csvCell(value) {
  const text = String(value ?? "");
  return /[",\n\r]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function save() {
  localStorage.setItem(storageKey, JSON.stringify(state.alerts));
}

function formatDate(value) {
  return new Intl.DateTimeFormat("ja-JP", {
    month: "2-digit",
    day: "2-digit",
  }).format(new Date(value));
}

function formatDateTime(value) {
  return new Intl.DateTimeFormat("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(value));
}

function hostname(url) {
  try {
    return new URL(url).hostname;
  } catch {
    return "";
  }
}

function debounce(callback, wait) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => callback(...args), wait);
  };
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
