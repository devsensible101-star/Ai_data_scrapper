const fs = require("fs");
const http = require("http");
const path = require("path");
const { spawn } = require("child_process");
const XLSX = require("xlsx");
const { DEFAULT_SCRAPER_CONFIG } = require("./index");

const ROOT_DIR = __dirname;
const PUBLIC_DIR = path.join(ROOT_DIR, "public");
const CONFIG_FILE = path.join(ROOT_DIR, "scraper-config.json");
const SCRAPER_FILE = path.join(ROOT_DIR, "index.js");
const PREVIOUS_FILE = path.join(ROOT_DIR, "previous_leads.xlsx");
const LATEST_FILE = path.join(ROOT_DIR, "latest_leads.xlsx");
const LATEST_CSV_FILE = path.join(ROOT_DIR, "latest_leads.csv");
const PORT = readPort(process.env.PORT, 3000);
const HOST = process.env.HOST || "127.0.0.1";
const MAX_BODY_BYTES = 1024 * 1024;
const MAX_LOG_LINES = 250;
const WORKBOOK_PAGE_SIZE_MAX = 250;

const staticFiles = new Map([
  ["/", { file: "index.html", type: "text/html; charset=utf-8" }],
  ["/index.html", { file: "index.html", type: "text/html; charset=utf-8" }],
  ["/app.js", { file: "app.js", type: "text/javascript; charset=utf-8" }],
  ["/styles.css", { file: "styles.css", type: "text/css; charset=utf-8" }]
]);

const workbookCache = new Map();
const runState = {
  running: false,
  startedAt: null,
  finishedAt: null,
  exitCode: null,
  error: "",
  progress: { completed: 0, total: 0 },
  logs: []
};

const server = http.createServer(async (request, response) => {
  try {
    const url = new URL(request.url, `http://${request.headers.host || `${HOST}:${PORT}`}`);
    setSecurityHeaders(response);

    if (request.method === "GET" && staticFiles.has(url.pathname)) {
      return serveStatic(response, staticFiles.get(url.pathname));
    }

    if (request.method === "GET" && url.pathname === "/favicon.ico") {
      response.writeHead(204, { "Cache-Control": "public, max-age=86400" });
      return response.end();
    }

    if (request.method === "GET" && url.pathname === "/api/config") {
      return sendJson(response, 200, { config: loadConfig() });
    }

    if (request.method === "GET" && url.pathname === "/api/status") {
      return sendJson(response, 200, { status: publicRunState() });
    }

    if (request.method === "GET" && url.pathname === "/api/workbooks") {
      return sendJson(response, 200, { workbooks: getWorkbookDescriptors() });
    }

    if (request.method === "GET" && url.pathname === "/api/leads") {
      return handleLeadPage(response, url.searchParams);
    }

    if (request.method === "GET" && url.pathname === "/api/download") {
      return handleDownload(response, url.searchParams);
    }

    if (request.method === "POST" && url.pathname === "/api/scrape") {
      const body = await readJsonBody(request);
      return startScrape(response, body);
    }

    return sendJson(response, 404, { error: "Route not found" });
  } catch (error) {
    const status = Number(error.statusCode) || 500;
    if (status >= 500) console.error(error);
    return sendJson(response, status, { error: error.message || "Unexpected server error" });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Lead Scraper Dashboard: http://${HOST}:${PORT}`);
});

function serveStatic(response, descriptor) {
  const filePath = path.join(PUBLIC_DIR, descriptor.file);
  if (!fs.existsSync(filePath)) return sendJson(response, 404, { error: "UI file not found" });
  response.writeHead(200, {
    "Content-Type": descriptor.type,
    "Cache-Control": "no-cache"
  });
  fs.createReadStream(filePath).pipe(response);
}

function handleLeadPage(response, searchParams) {
  const kind = searchParams.get("kind") === "previous" ? "previous" : "latest";
  const page = clampInteger(searchParams.get("page"), 1, 1, 1000000);
  const pageSize = clampInteger(searchParams.get("pageSize"), 50, 1, WORKBOOK_PAGE_SIZE_MAX);
  const query = cleanText(searchParams.get("q")).toLocaleLowerCase("en-IN");
  const filePath = resolveWorkbook(kind);

  if (!filePath) {
    return sendJson(response, 200, {
      file: null,
      columns: [],
      rows: [],
      total: 0,
      page: 1,
      pageSize,
      pages: 0
    });
  }

  const workbook = readWorkbook(filePath);
  const filteredRows = query
    ? workbook.rows.filter((row) =>
        workbook.columns.some((column) =>
          cleanText(row[column]).toLocaleLowerCase("en-IN").includes(query)
        )
      )
    : workbook.rows;
  const pages = Math.ceil(filteredRows.length / pageSize);
  const safePage = pages > 0 ? Math.min(page, pages) : 1;
  const start = (safePage - 1) * pageSize;

  return sendJson(response, 200, {
    file: describeWorkbook(kind, filePath, workbook.rows.length),
    columns: workbook.columns,
    rows: filteredRows.slice(start, start + pageSize),
    total: filteredRows.length,
    unfilteredTotal: workbook.rows.length,
    page: safePage,
    pageSize,
    pages
  });
}

function handleDownload(response, searchParams) {
  const kind = searchParams.get("kind") === "previous" ? "previous" : "latest";
  const filePath = resolveWorkbook(kind);
  if (!filePath) return sendJson(response, 404, { error: "Workbook not found" });

  const stat = fs.statSync(filePath);
  response.writeHead(200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition": `attachment; filename="${path.basename(filePath)}"`,
    "Content-Length": stat.size,
    "Cache-Control": "no-store"
  });
  fs.createReadStream(filePath).pipe(response);
}

function startScrape(response, body) {
  if (runState.running) {
    return sendJson(response, 409, { error: "A scraper run is already in progress" });
  }

  const config = validateConfig(body?.config || body || {});
  const currentLatest = resolveWorkbook("latest");
  if (currentLatest) {
    fs.copyFileSync(currentLatest, PREVIOUS_FILE);
  }

  fs.writeFileSync(CONFIG_FILE, `${JSON.stringify(config, null, 2)}\n`, "utf8");
  workbookCache.clear();

  Object.assign(runState, {
    running: true,
    startedAt: new Date().toISOString(),
    finishedAt: null,
    exitCode: null,
    error: "",
    latestLeadCount: null,
    progress: { completed: 0, total: calculateSearchCount(config) },
    logs: []
  });

  appendLog("system", `Starting scrape for ${config.city}`);
  const child = spawn(process.execPath, [SCRAPER_FILE], {
    cwd: ROOT_DIR,
    windowsHide: true,
    env: {
      ...process.env,
      SCRAPER_CITY: config.city,
      SCRAPER_LOCAL_KEYWORDS: JSON.stringify(config.localKeywords),
      SCRAPER_AREAS: JSON.stringify(config.areas),
      SCRAPER_B2B_QUERIES: JSON.stringify(config.b2bQueries),
      ENABLE_B2B: String(config.includeB2b),
      HEADLESS: String(config.headless),
      MAX_RESULTS: String(config.maxResults),
      QUERY_LIMIT: String(config.queryLimit),
      EXISTING_LEADS_FILE: currentLatest || "",
      OUTPUT_FILE: LATEST_FILE,
      CSV_OUTPUT_FILE: LATEST_CSV_FILE
    }
  });

  pipeChildOutput(child.stdout, "stdout");
  pipeChildOutput(child.stderr, "stderr");

  child.on("error", (error) => {
    runState.error = error.message;
    appendLog("stderr", `Could not start scraper: ${error.message}`);
  });

  child.on("close", (code) => {
    runState.running = false;
    runState.finishedAt = new Date().toISOString();
    runState.exitCode = code;
    if (code === 0) {
      runState.progress.completed = runState.progress.total;
      appendLog("system", "Scrape finished successfully");
    } else {
      runState.error ||= `Scraper stopped with exit code ${code}`;
      appendLog("stderr", runState.error);
    }
    workbookCache.clear();
  });

  return sendJson(response, 202, { status: publicRunState(), config });
}

function pipeChildOutput(stream, source) {
  let buffer = "";
  stream.setEncoding("utf8");
  stream.on("data", (chunk) => {
    buffer += chunk;
    const lines = buffer.split(/\r?\n/);
    buffer = lines.pop() || "";
    for (const line of lines) appendLog(source, line);
  });
  stream.on("end", () => {
    if (buffer) appendLog(source, buffer);
  });
}

function appendLog(source, input) {
  const message = cleanText(input);
  if (!message) return;

  const searchProgress = message.match(/^\[(\d+)\/(\d+)\]\s+Searching:/i);
  if (searchProgress) {
    runState.progress.completed = Math.max(0, Number(searchProgress[1]) - 1);
    runState.progress.total = Number(searchProgress[2]);
  }

  const checkpoint = message.match(/^Checkpoint saved\s+(\d+)\s+unique leads/i);
  if (checkpoint) runState.latestLeadCount = Number(checkpoint[1]);

  runState.logs.push({
    time: new Date().toISOString(),
    source,
    message
  });
  if (runState.logs.length > MAX_LOG_LINES) runState.logs.splice(0, runState.logs.length - MAX_LOG_LINES);
}

function publicRunState() {
  return {
    running: runState.running,
    startedAt: runState.startedAt,
    finishedAt: runState.finishedAt,
    exitCode: runState.exitCode,
    error: runState.error,
    progress: runState.progress,
    latestLeadCount: runState.latestLeadCount || null,
    logs: runState.logs
  };
}

function getWorkbookDescriptors() {
  return {
    previous: describeWorkbook("previous", resolveWorkbook("previous")),
    latest: describeWorkbook("latest", resolveWorkbook("latest"))
  };
}

function describeWorkbook(kind, filePath, knownRows) {
  if (!filePath || !fs.existsSync(filePath)) return null;
  const stat = fs.statSync(filePath);
  let rowCount = knownRows;
  if (rowCount == null) {
    try {
      rowCount = readWorkbook(filePath).rows.length;
    } catch (error) {
      rowCount = null;
    }
  }
  return {
    kind,
    filename: path.basename(filePath),
    rowCount,
    size: stat.size,
    updatedAt: stat.mtime.toISOString(),
    downloadUrl: `/api/download?kind=${kind}`
  };
}

function resolveWorkbook(kind) {
  if (kind === "previous") {
    if (fs.existsSync(PREVIOUS_FILE)) return PREVIOUS_FILE;
    const original = path.join(ROOT_DIR, "combined_leads.xlsx");
    return fs.existsSync(original) ? original : null;
  }

  return newestExisting([
    LATEST_FILE,
    path.join(ROOT_DIR, "latest_leads_new.xlsx"),
    path.join(ROOT_DIR, "combined_leads_new.xlsx"),
    path.join(ROOT_DIR, "combined_leads.xlsx")
  ]);
}

function newestExisting(paths) {
  return paths
    .filter((filePath) => fs.existsSync(filePath))
    .map((filePath) => ({ filePath, mtime: fs.statSync(filePath).mtimeMs }))
    .sort((first, second) => second.mtime - first.mtime)[0]?.filePath || null;
}

function readWorkbook(filePath) {
  const stat = fs.statSync(filePath);
  const cacheKey = `${filePath}:${stat.mtimeMs}:${stat.size}`;
  if (workbookCache.has(cacheKey)) return workbookCache.get(cacheKey);

  for (const key of workbookCache.keys()) {
    if (key.startsWith(`${filePath}:`)) workbookCache.delete(key);
  }

  const workbook = XLSX.readFile(filePath, { cellDates: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]] || {};
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const columns = uniqueValues([
    "name",
    "phone",
    "address",
    "category",
    "city",
    "source",
    ...rows.flatMap((row) => Object.keys(row))
  ]).filter((column) => rows.some((row) => Object.hasOwn(row, column)));
  const result = { columns, rows };
  workbookCache.set(cacheKey, result);
  return result;
}

function loadConfig() {
  if (!fs.existsSync(CONFIG_FILE)) return cloneDefaultConfig();
  try {
    return validateConfig(JSON.parse(fs.readFileSync(CONFIG_FILE, "utf8")));
  } catch (error) {
    console.warn(`Could not load scraper-config.json: ${error.message}`);
    return cloneDefaultConfig();
  }
}

function validateConfig(input) {
  const defaults = cloneDefaultConfig();
  const city = cleanText(input.city || defaults.city);
  const localKeywords = validateList(input.localKeywords, "Local keywords", defaults.localKeywords);
  const areas = validateList(input.areas, "Areas", defaults.areas);
  const b2bQueries = validateList(input.b2bQueries, "B2B queries", defaults.b2bQueries);
  const maxResults = clampInteger(input.maxResults, defaults.maxResults, 1, 100);
  const queryLimit = clampInteger(input.queryLimit, defaults.queryLimit, 0, 100000);

  if (!city || city.length > 100) throw httpError(400, "City must be between 1 and 100 characters");
  if (localKeywords.length > 100 || areas.length > 500 || b2bQueries.length > 100) {
    throw httpError(400, "The scraper configuration contains too many entries");
  }

  return {
    city,
    localKeywords,
    areas,
    b2bQueries,
    includeB2b: Boolean(input.includeB2b),
    headless: input.headless !== false,
    maxResults,
    queryLimit
  };
}

function validateList(value, label, fallback) {
  const source = Array.isArray(value) ? value : fallback;
  const values = uniqueValues(source).filter((item) => item.length <= 150);
  if (values.length === 0) throw httpError(400, `${label} must contain at least one entry`);
  return values;
}

function calculateSearchCount(config) {
  const localCount = config.localKeywords.length * config.areas.length;
  const b2bCount = config.includeB2b ? config.b2bQueries.length * 2 : 0;
  const total = localCount + b2bCount;
  return config.queryLimit > 0 ? Math.min(total, config.queryLimit) : total;
}

function cloneDefaultConfig() {
  return {
    ...DEFAULT_SCRAPER_CONFIG,
    localKeywords: [...DEFAULT_SCRAPER_CONFIG.localKeywords],
    areas: [...DEFAULT_SCRAPER_CONFIG.areas],
    b2bQueries: [...DEFAULT_SCRAPER_CONFIG.b2bQueries]
  };
}

function readJsonBody(request) {
  return new Promise((resolve, reject) => {
    let body = "";
    let size = 0;
    request.setEncoding("utf8");
    request.on("data", (chunk) => {
      size += Buffer.byteLength(chunk);
      if (size > MAX_BODY_BYTES) {
        reject(httpError(413, "Request body is too large"));
        request.destroy();
        return;
      }
      body += chunk;
    });
    request.on("end", () => {
      if (!body.trim()) return resolve({});
      try {
        resolve(JSON.parse(body));
      } catch (error) {
        reject(httpError(400, "Request body must be valid JSON"));
      }
    });
    request.on("error", reject);
  });
}

function sendJson(response, status, payload) {
  if (response.writableEnded) return;
  const body = JSON.stringify(payload);
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    "Cache-Control": "no-store"
  });
  response.end(body);
}

function setSecurityHeaders(response) {
  response.setHeader("X-Content-Type-Options", "nosniff");
  response.setHeader("X-Frame-Options", "DENY");
  response.setHeader("Referrer-Policy", "no-referrer");
  response.setHeader(
    "Content-Security-Policy",
    "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'"
  );
}

function uniqueValues(values) {
  const seen = new Set();
  return values
    .map(cleanText)
    .filter(Boolean)
    .filter((value) => {
      const key = value.toLocaleLowerCase("en-IN");
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function cleanText(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(/[\u0000-\u001F\u007F]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function clampInteger(value, fallback, min, max) {
  const number = Number(value);
  if (!Number.isInteger(number)) return fallback;
  return Math.min(max, Math.max(min, number));
}

function readPort(value, fallback) {
  return clampInteger(value, fallback, 1, 65535);
}

function httpError(statusCode, message) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}
