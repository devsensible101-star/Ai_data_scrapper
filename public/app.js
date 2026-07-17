const state = {
  activeKind: "latest",
  page: 1,
  pageSize: 50,
  pages: 0,
  query: "",
  workbooks: { previous: null, latest: null },
  wasRunning: false,
  searchTimer: null,
  toastTimer: null
};

const elements = {
  form: document.querySelector("#scrape-form"),
  city: document.querySelector("#city"),
  localKeywords: document.querySelector("#local-keywords"),
  areas: document.querySelector("#areas"),
  includeB2b: document.querySelector("#include-b2b"),
  b2bField: document.querySelector("#b2b-field"),
  b2bQueries: document.querySelector("#b2b-queries"),
  maxResults: document.querySelector("#max-results"),
  queryLimit: document.querySelector("#query-limit"),
  headless: document.querySelector("#headless"),
  plannedSearches: document.querySelector("#planned-searches"),
  runButton: document.querySelector("#run-button"),
  runButtonLabel: document.querySelector("#run-button-label"),
  formMessage: document.querySelector("#form-message"),
  statusPill: document.querySelector("#status-pill"),
  statusText: document.querySelector("#status-text"),
  runStatusTitle: document.querySelector("#run-status-title"),
  runStatusDetail: document.querySelector("#run-status-detail"),
  progressLabel: document.querySelector("#progress-label"),
  progressPercent: document.querySelector("#progress-percent"),
  runProgress: document.querySelector("#run-progress"),
  refreshButton: document.querySelector("#refresh-button"),
  previousCard: document.querySelector("#previous-card"),
  latestCard: document.querySelector("#latest-card"),
  previousCount: document.querySelector("#previous-count"),
  latestCount: document.querySelector("#latest-count"),
  previousMeta: document.querySelector("#previous-meta"),
  latestMeta: document.querySelector("#latest-meta"),
  tableTitle: document.querySelector("#table-title"),
  rowBadge: document.querySelector("#row-badge"),
  activeFileName: document.querySelector("#active-file-name"),
  tableSearch: document.querySelector("#table-search"),
  downloadButton: document.querySelector("#download-button"),
  leadHead: document.querySelector("#lead-head"),
  leadBody: document.querySelector("#lead-body"),
  pageSummary: document.querySelector("#page-summary"),
  pageIndicator: document.querySelector("#page-indicator"),
  previousPage: document.querySelector("#previous-page"),
  nextPage: document.querySelector("#next-page"),
  logPanel: document.querySelector("#log-panel"),
  logOutput: document.querySelector("#log-output"),
  logCount: document.querySelector("#log-count"),
  toast: document.querySelector("#toast")
};

initialize();

async function initialize() {
  bindEvents();
  try {
    const [{ config }, { status }, { workbooks }] = await Promise.all([
      requestJson("/api/config"),
      requestJson("/api/status"),
      requestJson("/api/workbooks")
    ]);
    populateForm(config);
    renderStatus(status);
    renderWorkbooks(workbooks);
    await loadTable();
  } catch (error) {
    showToast(error.message);
    renderTableError("Could not load the dashboard data.");
  }

  window.setInterval(pollStatus, 2000);
}

function bindEvents() {
  elements.form.addEventListener("submit", startScrape);
  elements.includeB2b.addEventListener("change", () => {
    updateB2bState();
    updatePlannedSearches();
  });

  for (const input of [
    elements.localKeywords,
    elements.areas,
    elements.b2bQueries,
    elements.queryLimit
  ]) {
    input.addEventListener("input", updatePlannedSearches);
  }

  elements.previousCard.addEventListener("click", () => selectWorkbook("previous"));
  elements.latestCard.addEventListener("click", () => selectWorkbook("latest"));
  elements.refreshButton.addEventListener("click", refreshWorkbooks);
  elements.previousPage.addEventListener("click", () => changePage(state.page - 1));
  elements.nextPage.addEventListener("click", () => changePage(state.page + 1));
  elements.tableSearch.addEventListener("input", () => {
    window.clearTimeout(state.searchTimer);
    state.searchTimer = window.setTimeout(() => {
      state.query = elements.tableSearch.value.trim();
      state.page = 1;
      loadTable();
    }, 280);
  });
}

function populateForm(config) {
  elements.city.value = config.city || "";
  elements.localKeywords.value = (config.localKeywords || []).join("\n");
  elements.areas.value = (config.areas || []).join("\n");
  elements.includeB2b.checked = Boolean(config.includeB2b);
  elements.b2bQueries.value = (config.b2bQueries || []).join("\n");
  elements.maxResults.value = config.maxResults ?? 15;
  elements.queryLimit.value = config.queryLimit ?? 0;
  elements.headless.checked = config.headless !== false;
  updateB2bState();
  updatePlannedSearches();
}

function collectConfig() {
  return {
    city: elements.city.value.trim(),
    localKeywords: linesFrom(elements.localKeywords.value),
    areas: linesFrom(elements.areas.value),
    includeB2b: elements.includeB2b.checked,
    b2bQueries: linesFrom(elements.b2bQueries.value),
    maxResults: Number(elements.maxResults.value),
    queryLimit: Number(elements.queryLimit.value),
    headless: elements.headless.checked
  };
}

function linesFrom(value) {
  return [...new Set(value.split(/\r?\n/).map((item) => item.trim()).filter(Boolean))];
}

function updateB2bState() {
  const enabled = elements.includeB2b.checked;
  elements.b2bQueries.disabled = !enabled;
  elements.b2bField.classList.toggle("disabled", !enabled);
}

function updatePlannedSearches() {
  const keywords = linesFrom(elements.localKeywords.value).length;
  const areas = linesFrom(elements.areas.value).length;
  const b2b = elements.includeB2b.checked ? linesFrom(elements.b2bQueries.value).length * 2 : 0;
  const total = keywords * areas + b2b;
  const limit = Number(elements.queryLimit.value);
  const planned = Number.isInteger(limit) && limit > 0 ? Math.min(total, limit) : total;
  elements.plannedSearches.textContent = planned.toLocaleString();
}

async function startScrape(event) {
  event.preventDefault();
  elements.formMessage.textContent = "";
  const config = collectConfig();

  if (!config.city || config.localKeywords.length === 0 || config.areas.length === 0) {
    elements.formMessage.textContent = "City, keywords, and areas are required.";
    return;
  }

  if (config.includeB2b && config.b2bQueries.length === 0) {
    elements.formMessage.textContent = "Add at least one B2B query or disable B2B sources.";
    return;
  }

  setRunButton(true);
  try {
    const { status } = await requestJson("/api/scrape", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ config })
    });
    state.wasRunning = true;
    renderStatus(status);
    await refreshWorkbooks();
    elements.logPanel.open = true;
    showToast("Scraper started. The previous workbook has been saved.");
  } catch (error) {
    elements.formMessage.textContent = error.message;
    setRunButton(false);
  }
}

async function pollStatus() {
  try {
    const { status } = await requestJson("/api/status");
    renderStatus(status);
    if (state.wasRunning && !status.running) {
      state.wasRunning = false;
      await refreshWorkbooks();
      await loadTable();
      showToast(status.exitCode === 0 ? "New Excel workbook is ready." : "Scraper run stopped with an error.");
    }
    if (status.running) state.wasRunning = true;
  } catch (error) {
    elements.statusText.textContent = "Server unavailable";
    elements.statusPill.dataset.state = "error";
  }
}

function renderStatus(status) {
  const completed = Number(status.progress?.completed || 0);
  const total = Number(status.progress?.total || 0);
  const percent = total > 0 ? Math.min(100, Math.round((completed / total) * 100)) : 0;
  setRunButton(status.running);
  renderLogs(status.logs || []);

  elements.runProgress.value = percent;
  elements.progressPercent.textContent = `${percent}%`;

  if (status.running) {
    elements.statusPill.dataset.state = "running";
    elements.statusText.textContent = "Scraper running";
    elements.runStatusTitle.textContent = "Generating a new workbook";
    elements.runStatusDetail.textContent = status.latestLeadCount
      ? `${status.latestLeadCount.toLocaleString()} unique leads saved at the latest checkpoint.`
      : "The browser is collecting and validating business leads.";
    elements.progressLabel.textContent = total
      ? `${completed.toLocaleString()} of ${total.toLocaleString()} searches started`
      : "Preparing searches";
    return;
  }

  if (status.error || (status.exitCode != null && status.exitCode !== 0)) {
    elements.statusPill.dataset.state = "error";
    elements.statusText.textContent = "Run failed";
    elements.runStatusTitle.textContent = "The last run did not finish";
    elements.runStatusDetail.textContent = status.error || "Open the run log for details.";
    elements.progressLabel.textContent = "Stopped before completion";
    return;
  }

  if (status.exitCode === 0 && status.finishedAt) {
    elements.statusPill.dataset.state = "success";
    elements.statusText.textContent = "Run complete";
    elements.runStatusTitle.textContent = "Latest workbook is ready";
    elements.runStatusDetail.textContent = `Finished ${formatRelativeTime(status.finishedAt)}.`;
    elements.progressLabel.textContent = `${total.toLocaleString()} searches completed`;
    elements.runProgress.value = 100;
    elements.progressPercent.textContent = "100%";
    return;
  }

  elements.statusPill.dataset.state = "idle";
  elements.statusText.textContent = "Ready";
  elements.runStatusTitle.textContent = "No scraper running";
  elements.runStatusDetail.textContent = "Configure your search and generate a new workbook when ready.";
  elements.progressLabel.textContent = "Waiting to start";
}

function setRunButton(running) {
  elements.runButton.disabled = running;
  elements.runButtonLabel.textContent = running ? "Generating leads…" : "Generate new leads";
}

async function refreshWorkbooks() {
  const { workbooks } = await requestJson("/api/workbooks");
  renderWorkbooks(workbooks);
}

function renderWorkbooks(workbooks) {
  state.workbooks = workbooks;
  renderWorkbookCard("previous", workbooks.previous);
  renderWorkbookCard("latest", workbooks.latest);
  updateDownloadButton();
}

function renderWorkbookCard(kind, workbook) {
  const countElement = kind === "previous" ? elements.previousCount : elements.latestCount;
  const metaElement = kind === "previous" ? elements.previousMeta : elements.latestMeta;
  if (!workbook) {
    countElement.textContent = "No file";
    metaElement.textContent = kind === "previous" ? "Created before your next run" : "Generate your first workbook";
    return;
  }

  countElement.textContent = workbook.rowCount == null
    ? "Workbook"
    : `${Number(workbook.rowCount).toLocaleString()} leads`;
  metaElement.textContent = `${workbook.filename} · ${formatBytes(workbook.size)} · ${formatRelativeTime(workbook.updatedAt)}`;
}

async function selectWorkbook(kind) {
  if (state.activeKind === kind) return;
  state.activeKind = kind;
  state.page = 1;
  elements.previousCard.classList.toggle("active", kind === "previous");
  elements.latestCard.classList.toggle("active", kind === "latest");
  elements.tableTitle.textContent = kind === "previous" ? "Previous workbook leads" : "Latest generated leads";
  updateDownloadButton();
  await loadTable();
}

function updateDownloadButton() {
  const workbook = state.workbooks[state.activeKind];
  if (!workbook) {
    elements.downloadButton.href = "#";
    elements.downloadButton.classList.add("disabled");
    elements.downloadButton.setAttribute("aria-disabled", "true");
    return;
  }
  elements.downloadButton.href = workbook.downloadUrl;
  elements.downloadButton.classList.remove("disabled");
  elements.downloadButton.setAttribute("aria-disabled", "false");
}

async function loadTable() {
  renderTableLoading();
  try {
    const params = new URLSearchParams({
      kind: state.activeKind,
      page: String(state.page),
      pageSize: String(state.pageSize),
      q: state.query
    });
    const data = await requestJson(`/api/leads?${params}`);
    state.page = data.page;
    state.pages = data.pages;
    renderTable(data);
  } catch (error) {
    renderTableError(error.message);
  }
}

function renderTable(data) {
  elements.leadHead.replaceChildren();
  elements.leadBody.replaceChildren();
  elements.rowBadge.textContent = `${Number(data.total).toLocaleString()} rows`;
  elements.activeFileName.textContent = data.file?.filename || "No workbook available";

  if (data.columns.length > 0) {
    const headerRow = document.createElement("tr");
    for (const column of data.columns) {
      const cell = document.createElement("th");
      cell.scope = "col";
      cell.textContent = humanizeColumn(column);
      headerRow.append(cell);
    }
    elements.leadHead.append(headerRow);
  }

  if (data.rows.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.className = "empty-cell";
    cell.colSpan = Math.max(1, data.columns.length);
    cell.textContent = data.file
      ? state.query
        ? "No leads match your search."
        : "This workbook does not contain any rows."
      : state.activeKind === "previous"
        ? "A previous snapshot will appear here after you start a new run."
        : "Generate a new workbook to see leads here.";
    row.append(cell);
    elements.leadBody.append(row);
  } else {
    for (const lead of data.rows) {
      const row = document.createElement("tr");
      for (const column of data.columns) {
        const cell = document.createElement("td");
        cell.textContent = String(lead[column] ?? "") || "—";
        cell.title = String(lead[column] ?? "");
        row.append(cell);
      }
      elements.leadBody.append(row);
    }
  }

  const start = data.total > 0 ? (data.page - 1) * data.pageSize + 1 : 0;
  const end = Math.min(data.total, data.page * data.pageSize);
  elements.pageSummary.textContent = data.total > 0
    ? `Showing ${start.toLocaleString()}–${end.toLocaleString()} of ${Number(data.total).toLocaleString()} leads`
    : "No rows to display";
  elements.pageIndicator.textContent = `Page ${data.page} of ${Math.max(1, data.pages)}`;
  elements.previousPage.disabled = data.page <= 1;
  elements.nextPage.disabled = data.pages === 0 || data.page >= data.pages;
}

function renderTableLoading() {
  elements.leadHead.replaceChildren();
  elements.leadBody.replaceChildren();
  const row = document.createElement("tr");
  const cell = document.createElement("td");
  cell.className = "empty-cell";
  cell.textContent = "Loading workbook…";
  row.append(cell);
  elements.leadBody.append(row);
}

function renderTableError(message) {
  elements.leadHead.replaceChildren();
  elements.leadBody.replaceChildren();
  const row = document.createElement("tr");
  const cell = document.createElement("td");
  cell.className = "empty-cell";
  cell.textContent = message;
  row.append(cell);
  elements.leadBody.append(row);
}

function renderLogs(logs) {
  elements.logCount.textContent = `${logs.length.toLocaleString()} ${logs.length === 1 ? "message" : "messages"}`;
  if (logs.length === 0) {
    elements.logOutput.textContent = "No run messages yet.";
    return;
  }

  elements.logOutput.replaceChildren();
  const fragment = document.createDocumentFragment();
  for (const log of logs) {
    const line = document.createElement("div");
    if (log.source === "stderr") line.className = "log-error";
    line.textContent = `[${formatLogTime(log.time)}] ${log.message}`;
    fragment.append(line);
  }
  elements.logOutput.append(fragment);
  if (elements.logPanel.open) elements.logOutput.scrollTop = elements.logOutput.scrollHeight;
}

function changePage(page) {
  if (page < 1 || page > state.pages || page === state.page) return;
  state.page = page;
  loadTable();
}

async function requestJson(url, options) {
  const response = await fetch(url, options);
  let payload = {};
  try {
    payload = await response.json();
  } catch (error) {
    throw new Error(`Server returned ${response.status}`);
  }
  if (!response.ok) throw new Error(payload.error || `Request failed with status ${response.status}`);
  return payload;
}

function humanizeColumn(value) {
  return String(value)
    .replace(/[_-]+/g, " ")
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 KB";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** index;
  return `${value.toFixed(index === 0 || value >= 10 ? 0 : 1)} ${units[index]}`;
}

function formatRelativeTime(value) {
  const time = new Date(value).getTime();
  if (!Number.isFinite(time)) return "unknown time";
  const seconds = Math.round((time - Date.now()) / 1000);
  const formatter = new Intl.RelativeTimeFormat("en", { numeric: "auto" });
  const ranges = [
    [60, "second"],
    [60, "minute"],
    [24, "hour"],
    [7, "day"],
    [4.345, "week"],
    [12, "month"],
    [Infinity, "year"]
  ];
  let valueToFormat = seconds;
  for (const [amount, unit] of ranges) {
    if (Math.abs(valueToFormat) < amount) return formatter.format(Math.round(valueToFormat), unit);
    valueToFormat /= amount;
  }
  return "recently";
}

function formatLogTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--:--:--";
  return date.toLocaleTimeString([], { hour12: false });
}

function showToast(message) {
  window.clearTimeout(state.toastTimer);
  elements.toast.textContent = message;
  elements.toast.classList.add("visible");
  state.toastTimer = window.setTimeout(() => elements.toast.classList.remove("visible"), 3500);
}
