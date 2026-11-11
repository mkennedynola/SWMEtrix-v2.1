"use strict";

const REQUIRED_COLUMNS = [
  "State Name", "Brand", "Account Invoice Date", "On/Off Premise",
  "Store Name", "Store Address", "Store City",
  "Item Description w/Vintage", "Volume CY"
];

const CHAIN_PATTERNS = [
  ["Walmart", /walmart/i],
  ["Kroger", /kroger/i],
  ["Sam's Club", /sam'?s\s*club/i],
  ["Target", /target/i],
  ["Whole Foods", /whole\s*foods/i],
  ["Walgreens", /walgreens/i],
  ["Harps", /\bharps\b/i]
];

const CHANNEL_CATEGORIES = [
  {
    title: "Off Premise Independent Trade Channel",
    match: (row) => isOffPremise(row) &&
      hasTerm(row.tradeChannel, ["independent", "independents", "indep"])
  },
  {
    title: "Off Premise Club Channel",
    match: (row) => isOffPremise(row) &&
      hasTerm(row.tradeChannel, ["club", "clubs"])
  },
  {
    title: "On Premise Independent Channel",
    match: (row) => isOnPremise(row) &&
      hasTerm(row.tradeChannel, ["independent", "independents", "indep"])
  },
  {
    title: "On Premise National Account Channel",
    match: (row) => isOnPremise(row) &&
      hasTerm(row.tradeChannel, ["national account", "national accounts", "national"])
  }
];

let allRows = [];
let charts = {};
let expandedDetailSections = new Set();
let hasTradeChannel = false;

const $ = (id) => document.getElementById(id);
const numberFormat = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 });
const volumeFormat = (value) => numberFormat.format(value);

const filterState = {
  states: new Set(),
  cities: new Set(),
  premises: new Set(),
  tradeChannels: new Set(),
  entities: new Set(),
  items: new Set()
};

function clean(value) {
  return value == null ? "" : String(value).trim();
}

function normalized(value) {
  return clean(value).toLowerCase().replace(/\s+/g, " ");
}

function hasTerm(value, terms) {
  const text = normalized(value);
  return terms.some((term) => text.includes(term));
}

function isOffPremise(row) {
  return normalized(row.premise).includes("off");
}

function isOnPremise(row) {
  const premise = normalized(row.premise);
  return premise.includes("on") && !premise.includes("off");
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number" && window.XLSX) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
  }

  const text = clean(value);
  if (!text) return null;

  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

function uniqueSorted(values) {
  return [...new Set(values.map(clean).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

function accountKey(row) {
  return [row.store, row.address, row.city, row.state]
    .map(normalized)
    .join("|");
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[char]);
}

function validateRows(rows) {
  if (!rows.length) return ["The selected file contains no data rows."];

  const headers = Object.keys(rows[0] || {}).map(clean);
  return REQUIRED_COLUMNS.filter((column) => !headers.includes(column));
}

function normalizeRows(rows) {
  return rows.map((row) => ({
    state: clean(row["State Name"]),
    brand: clean(row.Brand),
    date: parseDate(row["Account Invoice Date"]),
    premise: clean(row["On/Off Premise"]),
    tradeChannel: clean(row["Trade Channel"]),
    store: clean(row["Store Name"]),
    address: clean(row["Store Address"]),
    city: clean(row["Store City"]),
    product: clean(row["Item Description w/Vintage"]),
    volume: Number(String(row["Volume CY"] ?? 0).replace(/,/g, "")) || 0,
    storeDescription: clean(row["Store Description"])
  })).filter((row) => row.state || row.store || row.brand || row.product);
}

function resetSelectionState() {
  Object.keys(filterState).forEach((key) => {
    filterState[key] = new Set();
  });
}

function setDataControlsEnabled(enabled) {
  [
    "state-search",
    "city-search",
    "premise-search",
    "trade-search",
    "entity-search",
    "item-search",
    "date-filter",
    "view-filter"
  ].forEach((id) => {
    const element = $(id);
    if (element) element.disabled = !enabled;
  });

  $("reset-filters").disabled = !enabled;
  $("export-button").disabled = !enabled;
}

function buildCheckMenu(menuId, values, selectedSet, allLabel) {
  const container = $(menuId);
  if (!container) return;

  const sorted = values.slice().sort((a, b) => a.localeCompare(b));
  const allOption = document.createElement("label");
  allOption.className = "check-option all-option";
  allOption.innerHTML = `
    <input type="checkbox" data-menu="${menuId}" data-value="__ALL__" ${selectedSet.size === 0 ? "checked" : ""}>
    <span>${escapeHtml(allLabel)}</span>
  `;
  container.replaceChildren(allOption);

  sorted.forEach((value) => {
    const option = document.createElement("label");
    option.className = "check-option";
    option.innerHTML = `
      <input type="checkbox" data-menu="${menuId}" data-value="${escapeHtml(value)}" ${selectedSet.has(value) ? "checked" : ""}>
      <span>${escapeHtml(value)}</span>
    `;
    container.append(option);
  });
}

function applyFilterSearches() {
  [
    ["state-search", "state-menu"],
    ["city-search", "city-menu"],
    ["premise-search", "premise-menu"],
    ["trade-search", "trade-menu"],
    ["entity-search", "entity-menu"],
    ["item-search", "item-menu"]
  ].forEach(([searchId, menuId]) => {
    const search = $(searchId);
    const menu = $(menuId);
    if (!search || !menu) return;

    const query = search.value.toLowerCase();
    menu.querySelectorAll(".check-option").forEach((option) => {
      option.style.display =
        !query || option.textContent.toLowerCase().includes(query) ? "" : "none";
    });
  });
}

function dateMatches(date) {
  const range = $("date-filter").value;
  if (!date || range === "all") return true;

  const today = new Date();
  today.setHours(23, 59, 59, 999);

  const compareDate = new Date(date);
  compareDate.setHours(0, 0, 0, 0);

  if (range === "ytd") {
    return compareDate >= new Date(today.getFullYear(), 0, 1) &&
      compareDate <= today;
  }

  const start = new Date(today);
  start.setHours(0, 0, 0, 0);
  start.setDate(start.getDate() - Number(range) + 1);

  return compareDate >= start && compareDate <= today;
}

function filteredRows() {
  return allRows.filter((row) => {
    if (filterState.states.size && !filterState.states.has(row.state)) return false;
    if (filterState.cities.size && !filterState.cities.has(row.city)) return false;
    if (filterState.premises.size && !filterState.premises.has(row.premise)) return false;
    if (filterState.tradeChannels.size && !filterState.tradeChannels.has(row.tradeChannel)) return false;

    if ($("view-filter").value === "brand") {
      if (filterState.entities.size && !filterState.entities.has(row.brand)) return false;
      if (filterState.items.size && !filterState.items.has(row.product)) return false;
    } else if (filterState.entities.size && !filterState.entities.has(row.store)) {
      return false;
    }

    return dateMatches(row.date);
  });
}

function groupVolume(rows, key) {
  const groups = new Map();

  rows.forEach((row) => {
    const name = row[key] || "Unknown";
    groups.set(name, (groups.get(name) || 0) + row.volume);
  });

  return [...groups.entries()].sort((a, b) => b[1] - a[1]);
}

function accountSummary(rows) {
  const accounts = new Map();

  rows.forEach((row) => {
    const key = accountKey(row);
    if (!accounts.has(key)) {
      accounts.set(key, {
        name: row.store || "Unknown",
        address: row.address,
        city: row.city,
        state: row.state,
        volume: 0
      });
    }
    accounts.get(key).volume += row.volume;
  });

  return [...accounts.values()].sort((a, b) => b.volume - a.volume);
}

function renderFilterMenus() {
  const states = uniqueSorted(allRows.map((row) => row.state));
  const cities = uniqueSorted(
    allRows
      .filter((row) => !filterState.states.size || filterState.states.has(row.state))
      .map((row) => row.city)
  );

  const entityKey = $("view-filter").value === "brand" ? "brand" : "store";
  const entities = uniqueSorted(allRows.map((row) => row[entityKey]));
  const items = uniqueSorted(allRows.map((row) => row.product));

  buildCheckMenu("state-menu", states, filterState.states, "All States");
  buildCheckMenu("city-menu", cities, filterState.cities, "All Cities");
  buildCheckMenu(
    "premise-menu",
    uniqueSorted(allRows.map((row) => row.premise)),
    filterState.premises,
    "All Premises"
  );

  $("trade-filter-control").hidden = !hasTradeChannel;
  if (hasTradeChannel) {
    buildCheckMenu(
      "trade-menu",
      uniqueSorted(allRows.map((row) => row.tradeChannel)),
      filterState.tradeChannels,
      "All Channels"
    );
  }

  $("entity-label").textContent = entityKey === "brand" ? "Brand" : "Account";
  buildCheckMenu(
    "entity-menu",
    entities,
    filterState.entities,
    entityKey === "brand" ? "All Brands" : "All Accounts"
  );

  $("item-filter-control").hidden = entityKey !== "brand";
  if (entityKey === "brand") {
    buildCheckMenu("item-menu", items, filterState.items, "All Items");
  }

  applyFilterSearches();
}

function setChart(id, type, labels, values, color = "#6ee7b7") {
  const canvas = $(id);
  if (!canvas || !window.Chart) return;

  if (charts[id]) charts[id].destroy();

  const palette = [
    "#6ee7b7", "#60a5fa", "#fbbf24", "#f472b6", "#a78bfa",
    "#fb7185", "#2dd4bf", "#c084fc", "#f97316", "#93c5fd"
  ];

  charts[id] = new Chart(canvas, {
    type,
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: type === "pie" || type === "doughnut" ? palette : color,
        borderColor: type === "line" ? color : "transparent",
        borderWidth: type === "line" ? 2 : 0,
        borderRadius: type === "bar" ? 5 : 0,
        tension: 0.3,
        fill: false
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: type === "pie" || type === "doughnut",
          labels: { color: "#cbd5e1" }
        }
      },
      scales: type === "pie" || type === "doughnut" ? {} : {
        x: { ticks: { color: "#9aa8bb" }, grid: { color: "#263244" } },
        y: { beginAtZero: true, ticks: { color: "#9aa8bb" }, grid: { color: "#263244" } }
      }
    }
  });
}

function renderCharts(rows) {
  const cities = groupVolume(rows, "city").slice(0, 10);
  const accounts = groupVolume(rows, "store").slice(0, 10);
  const products = groupVolume(rows, "product").slice(0, 8);
  const monthly = new Map();

  rows.forEach((row) => {
    if (!row.date) return;
    const key = `${row.date.getFullYear()}-${String(row.date.getMonth() + 1).padStart(2, "0")}`;
    monthly.set(key, (monthly.get(key) || 0) + row.volume);
  });

  setChart("city-chart", "bar", cities.map(([name]) => name), cities.map(([, value]) => value));
  setChart("account-chart", "bar", accounts.map(([name]) => name), accounts.map(([, value]) => value), "#60a5fa");
  setChart("product-chart", "pie", products.map(([name]) => name), products.map(([, value]) => value));

  const months = [...monthly.keys()].sort();
  setChart("trend-chart", "line", months, months.map((month) => monthly.get(month)));

  const premiseGroups = groupVolume(rows, "premise");
  setChart("premise-share-chart", "doughnut",
    premiseGroups.map(([name]) => name),
    premiseGroups.map(([, value]) => value));

  const tradeGroups = groupVolume(rows, "tradeChannel");
  setChart("trade-share-chart", "doughnut",
    tradeGroups.map(([name]) => name),
    tradeGroups.map(([, value]) => value));
}

function detailCard(title, html) {
  return `<article class="detail-item"><h3>${escapeHtml(title)}</h3>${html}</article>`;
}

function detailList(title, entries) {
  const expanded = expandedDetailSections.has(title);
  const visible = expanded ? entries : entries.slice(0, 5);

  const list = visible.length
    ? `<ul>${visible.map(([name, value]) =>
      `<li>${escapeHtml(name)} <strong>· ${volumeFormat(value)}</strong></li>`
    ).join("")}</ul>`
    : `<p class="muted">No data for this selection.</p>`;

  const button = entries.length > 5
    ? `<button type="button" class="more-button" data-expand-section="${escapeHtml(title)}">${expanded ? "Show less" : `More (${entries.length - 5} more)`}</button>`
    : "";

  return detailCard(title, list + button);
}

function renderDetails(rows) {
  const isBrandView = $("view-filter").value === "brand";
  const selectedEntities = [...filterState.entities];
  const selectedItems = [...filterState.items];

  const detailsRows = rows.filter((row) => {
    const entityMatches = !selectedEntities.length ||
      selectedEntities.includes(isBrandView ? row.brand : row.store);
    const itemMatches = !isBrandView || !selectedItems.length ||
      selectedItems.includes(row.product);
    return entityMatches && itemMatches;
  });

  $("detail-content").className = "detail-grid";

  if (isBrandView) {
    $("detail-eyebrow").textContent = "BRAND INSIGHTS";
    $("detail-title").textContent = selectedEntities[0] || (selectedItems.length ? "Item View" : "Brand Dashboard");

    $("detail-content").innerHTML =
      detailCard("Brand", `<p><strong>${escapeHtml(selectedEntities.join(", ") || "All Brands")}</strong></p>`) +
      detailCard("Total Volume", `<p><strong>${volumeFormat(detailsRows.reduce((sum, row) => sum + row.volume, 0))}</strong></p>`) +
      detailCard("Distribution", `<p>Accounts: <strong>${new Set(detailsRows.map(accountKey)).size}</strong></p><p>Cities: <strong>${uniqueSorted(detailsRows.map((row) => row.city)).length}</strong></p><p>Products: <strong>${uniqueSorted(detailsRows.map((row) => row.product)).length}</strong></p>`) +
      detailList("Top Accounts", accountSummary(detailsRows).map((account) => [
        `${account.name} — ${[account.address, account.city].filter(Boolean).join(", ")}`,
        account.volume
      ])) +
      detailList("Top Cities", groupVolume(detailsRows, "city")) +
      detailList("Top Products", groupVolume(detailsRows, "product"));
    return;
  }

  const dates = detailsRows.map((row) => row.date).filter(Boolean).sort((a, b) => b - a);
  $("detail-eyebrow").textContent = "ACCOUNT INSIGHTS";
  $("detail-title").textContent = selectedEntities.length ? "Selected Accounts" : "Account Dashboard";

  $("detail-content").innerHTML =
    detailCard("Account", `<p><strong>${escapeHtml(selectedEntities.join(", ") || "All Accounts")}</strong></p>`) +
    detailCard("Account Metrics", `<p>Total volume: <strong>${volumeFormat(detailsRows.reduce((sum, row) => sum + row.volume, 0))}</strong></p><p>Brands: <strong>${uniqueSorted(detailsRows.map((row) => row.brand)).length}</strong></p><p>Products: <strong>${uniqueSorted(detailsRows.map((row) => row.product)).length}</strong></p><p>Last purchase: <strong>${dates[0] ? dates[0].toLocaleDateString() : "—"}</strong></p>`) +
    detailList("Top Brands", groupVolume(detailsRows, "brand")) +
    detailList("Top Products", groupVolume(detailsRows, "product"));
}

function renderChainAnalytics(rows) {
  const groups = new Map();

  rows.forEach((row) => {
    const matchedName = CHAIN_PATTERNS.find(([, pattern]) => pattern.test(row.store));
    const offPremiseChain = isOffPremise(row) && hasTerm(row.tradeChannel, ["chain"]);

    if (!matchedName && !offPremiseChain) return;

    const label = matchedName ? matchedName[0] : "Off Premise Chain";
    const key = `${label}|${accountKey(row)}`;

    if (!groups.has(key)) {
      groups.set(key, { label, rows: [] });
    }
    groups.get(key).rows.push(row);
  });

  const summaries = new Map();
  groups.forEach((group) => {
    if (!summaries.has(group.label)) summaries.set(group.label, []);
    summaries.get(group.label).push(...group.rows);
  });

  const tbody = $("chain-table");
  tbody.replaceChildren();

  if (!summaries.size) {
    tbody.innerHTML = `<tr><td colspan="5" class="muted">No recognized chain accounts in this selection.</td></tr>`;
    return;
  }

  summaries.forEach((groupRows, label) => {
    const tr = document.createElement("tr");
    const values = [
      label,
      volumeFormat(groupRows.reduce((sum, row) => sum + row.volume, 0)),
      new Set(groupRows.map(accountKey)).size,
      uniqueSorted(groupRows.map((row) => row.brand)).length,
      uniqueSorted(groupRows.map((row) => row.product)).length
    ];

    values.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.append(td);
    });
    tbody.append(tr);
  });
}

function renderChannelAnalytics(rows) {
  const container = $("channel-analytics-grid");
  if (!container) return;

  // rows is already filtered by the dashboard. Apply the state selection
  // explicitly here as well, so channel totals cannot include other states.
  const selectedStateRows = rows.filter((row) =>
    filterState.states.size === 0 || filterState.states.has(row.state)
  );

  container.replaceChildren();

  CHANNEL_CATEGORIES.forEach((category) => {
    const matchingRows = selectedStateRows.filter(category.match);
    const accountCount = new Set(matchingRows.map(accountKey)).size;
    const volume = matchingRows.reduce((sum, row) => sum + row.volume, 0);

    const card = document.createElement("article");
    card.className = "analytics-card";
    card.innerHTML = `
      <p class="eyebrow">CHANNEL CATEGORY</p>
      <h3>${escapeHtml(category.title)}</h3>
      <div class="table-scroll">
        <table>
          <tbody>
            <tr><th>Volume</th><td>${volumeFormat(volume)}</td></tr>
            <tr><th>Accounts</th><td>${accountCount}</td></tr>
            <tr><th>Brands</th><td>${uniqueSorted(matchingRows.map((row) => row.brand)).length}</td></tr>
            <tr><th>Products</th><td>${uniqueSorted(matchingRows.map((row) => row.product)).length}</td></tr>
          </tbody>
        </table>
      </div>
    `;
    container.append(card);
  });
}

function renderPlacementRows(tbody, placements, emptyMessage) {
  tbody.replaceChildren();

  if (!placements.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 3;
    td.className = "muted";
    td.textContent = emptyMessage;
    tr.append(td);
    tbody.append(tr);
    return;
  }

  placements.forEach((placement) => {
    const tr = document.createElement("tr");

    [placement.account, placement.city || "—", placement.product].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.append(td);
    });

    tbody.append(tr);
  });
}

function renderNewPlacements() {
  const today = new Date();
  today.setHours(23, 59, 59, 999);

  const yearStart = new Date(today.getFullYear(), 0, 1);
  yearStart.setHours(0, 0, 0, 0);

  const windowStart = new Date(today);
  windowStart.setHours(0, 0, 0, 0);
  windowStart.setDate(windowStart.getDate() - 29);

  // Use all loaded rows, not dashboard-filtered rows, to identify genuine
  // first purchases for the year.
  const firstPurchases = new Map();

  allRows.forEach((row) => {
    if (!row.date || row.volume <= 0 || !row.store || !row.product) return;

    const date = new Date(row.date);
    if (date < yearStart || date > today) return;

    const premiseType = isOffPremise(row) ? "off" : isOnPremise(row) ? "on" : "";
    if (!premiseType) return;

    const key = `${accountKey(row)}|${normalized(row.product)}`;
    const current = firstPurchases.get(key);

    if (!current || date < current.date) {
      firstPurchases.set(key, {
        date,
        account: row.store,
        city: row.city,
        product: row.product,
        premiseType
      });
    }
  });

  const recentPlacements = [...firstPurchases.values()]
    .filter((placement) => placement.date >= windowStart)
    .sort((a, b) =>
      a.account.localeCompare(b.account) ||
      a.product.localeCompare(b.product)
    );

  renderPlacementRows(
    $("on-premise-placements"),
    recentPlacements.filter((placement) => placement.premiseType === "on"),
    "No new on-premise placements found in the last 30 days."
  );

  renderPlacementRows(
    $("off-premise-placements"),
    recentPlacements.filter((placement) => placement.premiseType === "off"),
    "No new off-premise placements found in the last 30 days."
  );
}

function refresh() {
  const rows = filteredRows();

  $("kpi-records").textContent = numberFormat.format(rows.length);
  $("kpi-accounts").textContent = numberFormat.format(new Set(rows.map(accountKey)).size);
  $("kpi-brands").textContent = numberFormat.format(uniqueSorted(rows.map((row) => row.brand)).length);
  $("kpi-cities").textContent = numberFormat.format(uniqueSorted(rows.map((row) => row.city)).length);
  $("kpi-volume").textContent = volumeFormat(rows.reduce((sum, row) => sum + row.volume, 0));

  renderDetails(rows);
  renderCharts(rows);
  renderChainAnalytics(rows);
  renderChannelAnalytics(rows);
  renderNewPlacements();
}

function resetFilters() {
  resetSelectionState();
  $("date-filter").value = "all";
  $("view-filter").value = "brand";

  ["state", "city", "premise", "trade", "entity", "item"].forEach((name) => {
    const search = $(`${name}-search`);
    if (search) search.value = "";
  });

  expandedDetailSections.clear();
  renderFilterMenus();
  refresh();
}

function exportCurrentView() {
  const rows = filteredRows();
  const output = rows.map((row) => ({
    "State Name": row.state,
    Brand: row.brand,
    "Account Invoice Date": row.date ? row.date.toISOString().slice(0, 10) : "",
    "On/Off Premise": row.premise,
    "Store Name": row.store,
    "Store Address": row.address,
    "Store City": row.city,
    "Item Description w/Vintage": row.product,
    "Volume CY": row.volume,
    "Trade Channel": row.tradeChannel,
    "Store Description": row.storeDescription
  }));

  const sheet = XLSX.utils.json_to_sheet(output);
  const csv = XLSX.utils.sheet_to_csv(sheet);
  const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = "market-placement-current-view.csv";
  link.click();
  URL.revokeObjectURL(url);
}

async function parseFile(file) {
  let workbook;

  if (/\.csv$/i.test(file.name)) {
    workbook = XLSX.read(await file.text(), { type: "string", cellDates: true });
  } else {
    workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  }

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: "", cellDates: true });
}

async function loadRows(sourceRows, filename) {
  const missing = validateRows(sourceRows);
  if (missing.length) {
    throw new Error(`Missing required column${missing.length > 1 ? "s" : ""}: ${missing.join(", ")}`);
  }

  hasTradeChannel = Object.keys(sourceRows[0] || {}).map(clean).includes("Trade Channel");
  allRows = normalizeRows(sourceRows);

  if (!allRows.length) {
    throw new Error("No usable data rows were found.");
  }

  resetSelectionState();
  renderFilterMenus();
  setDataControlsEnabled(true);
  $("file-status").textContent = `${filename} · ${numberFormat.format(allRows.length)} records loaded`;
  refresh();
}

async function importFile(file) {
  $("import-error").hidden = true;
  $("file-status").textContent = `Reading ${file.name}…`;

  try {
    const rows = await parseFile(file);
    await loadRows(rows, file.name);
  } catch (error) {
    allRows = [];
    setDataControlsEnabled(false);
    $("file-status").textContent = "Import failed.";
    $("import-error").textContent = error.message || "Unable to read this file.";
    $("import-error").hidden = false;
  }
}

async function autoLoadSalesCsv() {
  $("file-status").textContent = "Looking for sales.csv…";

  try {
    const response = await fetch("./sales.csv", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Could not load sales.csv (HTTP ${response.status}).`);
    }

    const workbook = XLSX.read(await response.text(), { type: "string", cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", cellDates: true });
    await loadRows(rows, "sales.csv");
  } catch (error) {
    $("file-status").textContent = "Automatic load unavailable. You can load a file manually.";
    $("import-error").textContent = error.message;
    $("import-error").hidden = false;
  }
}

function bindMenus() {
  const menuMap = {
    "state-menu": "states",
    "city-menu": "cities",
    "premise-menu": "premises",
    "trade-menu": "tradeChannels",
    "entity-menu": "entities",
    "item-menu": "items"
  };

  Object.entries(menuMap).forEach(([menuId, setName]) => {
    $(menuId).addEventListener("change", (event) => {
      const input = event.target;
      if (!input.matches('input[type="checkbox"]')) return;

      const selected = filterState[setName];
      if (input.dataset.value === "__ALL__") {
        selected.clear();
      } else if (input.checked) {
        selected.add(input.dataset.value);
      } else {
        selected.delete(input.dataset.value);
      }

      renderFilterMenus();
      refresh();
    });
  });
}

function init() {
  setDataControlsEnabled(false);
  bindMenus();

  [
    "state-search",
    "city-search",
    "premise-search",
    "trade-search",
    "entity-search",
    "item-search"
  ].forEach((id) => {
    $(id).addEventListener("input", applyFilterSearches);
  });

  $("date-filter").addEventListener("change", refresh);
  $("view-filter").addEventListener("change", () => {
    filterState.entities.clear();
    filterState.items.clear();
    renderFilterMenus();
    refresh();
  });

  $("reset-filters").addEventListener("click", resetFilters);
  $("export-button").addEventListener("click", exportCurrentView);

  $("file-input").addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (file) await importFile(file);
    event.target.value = "";
  });

  $("detail-content").addEventListener("click", (event) => {
    const button = event.target.closest("[data-expand-section]");
    if (!button) return;

    const title = button.dataset.expandSection;
    if (expandedDetailSections.has(title)) {
      expandedDetailSections.delete(title);
    } else {
      expandedDetailSections.add(title);
    }
    refresh();
  });

  autoLoadSalesCsv();
}

init();