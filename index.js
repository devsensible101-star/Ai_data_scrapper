const axios = require("axios");
const cheerio = require("cheerio");
const fs = require("fs");
const http = require("http");
const https = require("https");
const path = require("path");
const puppeteer = require("puppeteer");
const XLSX = require("xlsx");

const DEFAULT_SCRAPER_CONFIG = {
  city: "Bengaluru",
  localKeywords: ["RFID", "HRMS", "ERP", "Access Control"],
  areas: [
    "South Bengaluru",
    "Jayanagar",
    "JP Nagar",
    "Banashankari",
    "BTM Layout",
    "Koramangala",
    "HSR Layout",
    "Bommanahalli",
    "Bommasandra",
    "Electronic City",
    "Begur",
    "Arekere",
    "Hulimavu",
    "Gottigere",
    "Bannerghatta Road",
    "Anjanapura",
    "Uttarahalli",
    "Kumaraswamy Layout",
    "Padmanabhanagar",
    "Girinagar",
    "Basaveshwar Nagar",
    "Kanakapura Road",
    "Talaghattapura",
    "Konanakunte",
    "Kothanur (South)"
  ],
  b2bQueries: [
    "pos machine supplier",
    "billing machine dealer",
    "restaurant billing software provider"
  ],
  includeB2b: false,
  headless: true,
  maxResults: 15,
  queryLimit: 0
};

const CITY = cleanText(process.env.SCRAPER_CITY || DEFAULT_SCRAPER_CONFIG.city);
const LOCAL_KEYWORDS = readListEnv("SCRAPER_LOCAL_KEYWORDS", DEFAULT_SCRAPER_CONFIG.localKeywords);
const AREAS = readListEnv("SCRAPER_AREAS", DEFAULT_SCRAPER_CONFIG.areas);
const B2B_QUERIES = readListEnv("SCRAPER_B2B_QUERIES", DEFAULT_SCRAPER_CONFIG.b2bQueries);
const ENABLE_B2B = readBooleanEnv("ENABLE_B2B", DEFAULT_SCRAPER_CONFIG.includeB2b);

const HEADLESS = readBooleanEnv("HEADLESS", DEFAULT_SCRAPER_CONFIG.headless);
const MAX_RETRIES = readIntegerEnv("MAX_RETRIES", 1, 0, 5);
const MAX_RESULTS = readIntegerEnv("MAX_RESULTS", DEFAULT_SCRAPER_CONFIG.maxResults, 1, 100);
const MAX_SCROLL_ATTEMPTS = readIntegerEnv("MAX_SCROLL_ATTEMPTS", 8, 1, 50);
const PLACE_CONCURRENCY = Math.max(
  1,
  Math.min(readIntegerEnv("PLACE_CONCURRENCY", 3, 1, 5), 5)
);
const SEARCH_CONCURRENCY = Math.max(
  1,
  Math.min(readIntegerEnv("SEARCH_CONCURRENCY", 2, 1, 3), 3)
);
const QUERY_LIMIT = readIntegerEnv("QUERY_LIMIT", DEFAULT_SCRAPER_CONFIG.queryLimit, 0, 100000);
const DETAIL_WAIT_MS = readIntegerEnv("DETAIL_WAIT_MS", 8000, 1000, 120000);
const CHECKPOINT_EVERY = readIntegerEnv("CHECKPOINT_EVERY", 5, 1, 10000);
const SCRAPER_FILTER = normalizeKeyText(process.env.SCRAPER_FILTER || "");
const BROWSER_EXECUTABLE_PATH = resolveBrowserExecutablePath();
const OUTPUT_FILE = readPathEnv("OUTPUT_FILE", "combined_leads.xlsx");
const CSV_OUTPUT_FILE = readPathEnv("CSV_OUTPUT_FILE", "combined_leads.csv");
const EXISTING_LEADS_FILE = readPathEnv("EXISTING_LEADS_FILE", "");
const OUTPUT_COLUMNS = ["name", "phone", "address", "category", "city", "source"];
const BLOCKED_RESOURCE_TYPES = new Set(["image", "font", "media"]);
const MAX_ACTIVE_BROWSER_PAGES = Math.max(SEARCH_CONCURRENCY, PLACE_CONCURRENCY);
const PHONE_CANDIDATE_REGEX = /(?:\+|00)?\d[\d\s().-]{6,}\d/g;

const seenGooglePlaceKeys = new Set();
const pageSemaphore = createSemaphore(MAX_ACTIVE_BROWSER_PAGES);
const pageSlotReleases = new WeakMap();

let googleConsentHandled = false;
let googleConsentHandlingPromise = null;

const metrics = {
  searchesPlanned: 0,
  searchesCompleted: 0,
  searchesFailed: 0,
  placeUrlsFound: 0,
  duplicatePlacesSkipped: 0,
  detailPagesOpened: 0,
  validLeads: 0,
  invalidLeads: 0,
  existingLeadsLoaded: 0,
  newUniqueLeads: 0,
  totalUniqueLeads: 0,
  retries: 0,
  rateLimits: 0
};

const userAgents = [
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
];

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const httpClient = axios.create({
  timeout: 20000,
  httpAgent: new http.Agent({ keepAlive: true, maxSockets: 8 }),
  httpsAgent: new https.Agent({ keepAlive: true, maxSockets: 8 }),
  validateStatus: (status) => status >= 200 && status < 400
});

async function main() {
  const startedAt = Date.now();
  const { rows: loadedRows, filePath: loadedFrom } = loadExistingLeads();
  const existingRows = dedupeLeads(loadedRows);
  const searches = limitItems(dedupeSearches(buildSearches()), QUERY_LIMIT);
  const state = {
    rows: existingRows,
    processedSearches: 0,
    lastCheckpointAt: 0
  };

  metrics.existingLeadsLoaded = existingRows.length;
  metrics.searchesPlanned = searches.length;

  console.log(
    `Loaded ${existingRows.length} existing leads${loadedFrom ? ` from ${loadedFrom}` : ""}`
  );
  console.log(`Starting lead scrape for ${searches.length} searches`);
  console.log(
    `Concurrency: ${SEARCH_CONCURRENCY} searches, ${PLACE_CONCURRENCY} place workers, ${MAX_ACTIVE_BROWSER_PAGES} total browser pages`
  );
  logSearchValidation(searches);

  let browser = null;
  let fatalError = null;

  try {
    const browserSources = new Set(["Google Maps", "Justdial", "Sulekha"]);
    const needsBrowser = searches.some((search) => browserSources.has(search.source));
    if (needsBrowser) {
      if (BROWSER_EXECUTABLE_PATH) {
        console.log(`Using installed browser: ${path.basename(BROWSER_EXECUTABLE_PATH)}`);
      }
      browser = await puppeteer.launch({
        headless: HEADLESS,
        ...(BROWSER_EXECUTABLE_PATH ? { executablePath: BROWSER_EXECUTABLE_PATH } : {}),
        protocolTimeout: 120000,
        defaultViewport: { width: 1366, height: 768 },
        args: [
          "--lang=en-US,en",
          "--no-sandbox",
          "--disable-setuid-sandbox",
          "--disable-dev-shm-usage",
          "--disable-extensions",
          "--disable-quic",
          "--disable-blink-features=AutomationControlled",
          "--disable-background-networking",
          "--disable-sync"
        ]
      });
    }

    await mapWithConcurrency(searches, SEARCH_CONCURRENCY, async (search, index) => {
      console.log(`[${index + 1}/${searches.length}] Searching: ${search.text} (${search.source})`);
      let leads = [];

      try {
        leads = await scrapeSingleSearch(browser, search);
        metrics.searchesCompleted += 1;
      } catch (error) {
        metrics.searchesFailed += 1;
        if (isRateLimitError(error)) {
          reportRateLimitIfNeeded(error, `${search.source} search`);
          await sleep(calculateBackoffMs(0, true));
        }
        console.warn(
          `[${index + 1}/${searches.length}] Search failed: ${search.source} | ${shortError(error)}`
        );
      }

      const added = mergeLeadsIntoState(state, leads);
      state.processedSearches += 1;
      console.log(`${search.source} | ${search.text} | ${added} new unique leads`);

      if (state.processedSearches - state.lastCheckpointAt >= CHECKPOINT_EVERY) {
        try {
          saveCheckpoint(state, `after ${state.processedSearches}/${searches.length} searches`);
        } catch (error) {
          console.warn(`Checkpoint failed: ${shortError(error)}`);
        }
      }
    });
  } catch (error) {
    fatalError = error;
    const accountedSearches = metrics.searchesCompleted + metrics.searchesFailed;
    metrics.searchesFailed += Math.max(0, metrics.searchesPlanned - accountedSearches);
    console.error(`Fatal scraper error: ${shortError(error)}`);
  } finally {
    try {
      saveCheckpoint(state, "final browser shutdown");
    } catch (error) {
      console.error(`Final save failed: ${shortError(error)}`);
      fatalError ||= error;
    }

    if (browser) {
      try {
        await browser.close();
      } catch (error) {
        console.warn(`Browser close warning: ${shortError(error)}`);
      }
    }
  }

  metrics.newUniqueLeads = Math.max(0, state.rows.length - existingRows.length);
  metrics.totalUniqueLeads = state.rows.length;
  printMetrics(startedAt);

  if (fatalError) throw fatalError;
}

function buildSearches() {
  const localSearches = LOCAL_KEYWORDS.flatMap((category) =>
    AREAS.flatMap((area) => {
      const text = `${category} in ${area} ${CITY}`;
      return [
        { source: "Google Maps", type: "local", category, area, city: CITY, text }
        // { source: "Justdial", type: "local", category, area, city: CITY, text },
        // { source: "Sulekha", type: "local", category, area, city: CITY, text }
      ];
    })
  );

  const b2bSearches = ENABLE_B2B
    ? B2B_QUERIES.flatMap((text) => [
        { source: "IndiaMART", type: "b2b", category: text, city: CITY, text },
        { source: "TradeIndia", type: "b2b", category: text, city: CITY, text }
      ])
    : [];

  const allSearches = [...localSearches, ...b2bSearches];
  if (!SCRAPER_FILTER) return allSearches;
  return allSearches.filter((search) => normalizeKeyText(search.source) === SCRAPER_FILTER);
}

function dedupeSearches(searches) {
  const seen = new Set();
  const unique = [];

  for (const search of searches) {
    const key = [search.source, search.category, search.area, search.city]
      .map(normalizeKeyText)
      .join("|");
    if (!seen.has(key)) {
      seen.add(key);
      unique.push(search);
    }
  }

  return unique;
}

function limitItems(items, limit) {
  return limit > 0 ? items.slice(0, limit) : items;
}

function logSearchValidation(searches) {
  const localSearches = searches.filter((search) => search.type === "local");
  const sources = uniqueValues(searches.map((search) => search.source));
  const categories = uniqueValues(localSearches.map((search) => search.category));
  const areas = uniqueValues(localSearches.map((search) => search.area));

  console.log(`Active sources: ${sources.join(", ") || "none"}`);
  console.log(`Active categories: ${categories.join(", ") || "none"}`);
  console.log(`Active areas: ${areas.join(", ") || "none"}`);
}

function uniqueValues(values) {
  return [...new Set(values.map(cleanText).filter(Boolean))];
}

async function mapWithConcurrency(items, concurrency, handler) {
  const results = new Array(items.length);
  const errors = [];
  let nextIndex = 0;

  async function worker() {
    while (true) {
      const index = nextIndex;
      nextIndex += 1;
      if (index >= items.length) return;
      try {
        results[index] = await handler(items[index], index);
      } catch (error) {
        errors.push(error);
      }
    }
  }

  const workerCount = Math.min(Math.max(1, concurrency), items.length);
  await Promise.all(Array.from({ length: workerCount }, () => worker()));
  if (errors.length > 0) throw errors[0];
  return results;
}

async function scrapeSingleSearch(browser, search) {
  switch (search.source) {
    case "Google Maps":
      return scrapeGoogleMaps(browser, search);
    case "Justdial":
      return scrapeJustdial(browser, search);
    case "Sulekha":
      return scrapeSulekha(browser, search);
    case "IndiaMART":
      return scrapeIndiaMart(search);
    case "TradeIndia":
      return scrapeTradeIndia(search);
    default:
      throw new Error(`Unsupported scraper source: ${search.source}`);
  }
}

async function scrapeGoogleMaps(browser, search) {
  const placeUrls = await runWithFreshPageOnCrash(
    browser,
    {},
    async (page) => {
      const url = `https://www.google.com/maps/search/${encodeURIComponent(search.text)}?hl=en`;
      await navigateWithRetry(
        page,
        url,
        { waitUntil: "domcontentloaded", timeout: 45000 },
        "Google Maps search navigation",
        true
      );
      await acceptGoogleConsentOnce(page);

      try {
        await waitForMapsResults(page);
      } catch (error) {
        const blockReason = await detectGoogleBlock(page);
        if (blockReason) throw createRateLimitError(`Google Maps block detected: ${blockReason}`);
        throw error;
      }

      await autoScrollMaps(page, MAX_RESULTS);
      return page.$$eval("a.hfpxzc", (links) =>
        [...new Set(links.map((link) => link.href).filter(Boolean))]
      );
    },
    "Google Maps search page"
  );

  metrics.placeUrlsFound += placeUrls.length;
  const newPlaceUrls = filterNewGooglePlaceUrls(placeUrls, MAX_RESULTS);
  const leads = await scrapePlacesWithWorkers(browser, newPlaceUrls, search);
  console.log(`Leads found per source: Google Maps ${leads.length}`);
  return leads;
}

function extractGooglePlaceKey(url) {
  const rawUrl = cleanText(url);
  if (!rawUrl) return "";

  try {
    const parsed = new URL(rawUrl);
    const directId =
      parsed.searchParams.get("query_place_id") ||
      parsed.searchParams.get("place_id") ||
      parsed.searchParams.get("ftid") ||
      parsed.searchParams.get("cid");
    if (directId) return `place:${normalizeKeyText(directId)}`;

    const decoded = safeDecodeURIComponent(`${parsed.pathname}${parsed.search}`);
    const placeIdMatch = decoded.match(/(?:!1s|place_id[:=])([^!&/?]+)/i);
    if (placeIdMatch?.[1]) return `place:${normalizeKeyText(placeIdMatch[1])}`;

    const volatileParams = ["hl", "entry", "g_st", "authuser", "ved", "sa", "source"];
    for (const param of volatileParams) parsed.searchParams.delete(param);
    parsed.hash = "";
    parsed.hostname = parsed.hostname.toLowerCase();
    parsed.pathname = parsed.pathname.replace(/\/+$/, "");
    parsed.searchParams.sort();
    return `url:${parsed.toString().toLowerCase()}`;
  } catch (error) {
    return `url:${normalizeKeyText(rawUrl)}`;
  }
}

function filterNewGooglePlaceUrls(urls, limit = Infinity) {
  const fresh = [];
  const localKeys = new Set();

  for (const url of urls) {
    const key = extractGooglePlaceKey(url);
    if (!key) continue;

    if (localKeys.has(key) || seenGooglePlaceKeys.has(key)) {
      metrics.duplicatePlacesSkipped += 1;
      continue;
    }

    localKeys.add(key);
    if (fresh.length >= limit) continue;
    seenGooglePlaceKeys.add(key);
    fresh.push(url);
  }

  return fresh;
}

async function scrapePlacesWithWorkers(browser, urls, search) {
  if (urls.length === 0) return [];

  const leads = [];
  let nextIndex = 0;
  const workerCount = Math.min(PLACE_CONCURRENCY, urls.length);

  async function worker() {
    let page = await newPage(browser);

    try {
      while (true) {
        const index = nextIndex;
        nextIndex += 1;
        if (index >= urls.length) return;

        const url = urls[index];
        metrics.detailPagesOpened += 1;
        console.log(`Opening place ${index + 1}/${urls.length} for ${search.text}`);

        let recoveredFromCrash = false;
        while (true) {
          try {
            const lead = await scrapeGoogleMapsDetailWithPage(page, url, search);
            if (lead) leads.push(lead);
            break;
          } catch (error) {
            if (!recoveredFromCrash && MAX_RETRIES > 0 && isPageCrashError(error)) {
              recoveredFromCrash = true;
              metrics.retries += 1;
              console.warn("Google Maps detail page crashed; reopening one worker page");
              await sleep(calculateBackoffMs(0, false));
              await safeClosePage(page);
              page = await newPage(browser);
              continue;
            }

            if (isRateLimitError(error)) {
              reportRateLimitIfNeeded(error, "Google Maps detail page");
              await sleep(calculateBackoffMs(0, true));
            }
            metrics.invalidLeads += 1;
            console.warn(`Google Maps detail rejected: ${shortError(error)}`);
            break;
          }
        }
      }
    } finally {
      await safeClosePage(page);
    }
  }

  const workerResults = await Promise.allSettled(
    Array.from({ length: workerCount }, () => worker())
  );
  const failedWorker = workerResults.find((result) => result.status === "rejected");
  if (failedWorker) throw failedWorker.reason;
  return leads;
}

async function scrapeGoogleMapsDetailWithPage(page, url, search) {
  await navigateWithRetry(
    page,
    url,
    { waitUntil: "domcontentloaded", timeout: 30000 },
    "Google Maps detail navigation",
    true
  );
  await acceptGoogleConsentOnce(page);

  try {
    await waitForLeadFields(page, DETAIL_WAIT_MS);
  } catch (error) {
    if (!isTimeoutError(error)) throw error;
    console.log("Google Maps detail fields timed out; extracting available data");
  }

  const blockReason = await detectGoogleBlock(page);
  if (blockReason) throw createRateLimitError(`Google Maps block detected: ${blockReason}`);

  let data = await extractGoogleMapsDetail(page);
  if (!cleanPhone(data.phone)) {
    await page
      .waitForSelector('button[data-item-id^="phone"], a[href^="tel:"]', { timeout: 2500 })
      .catch(() => null);
    data = await extractGoogleMapsDetail(page);
  }

  const lead = normalizeLead({
    ...data,
    category: data.category || search.category,
    city: search.city,
    source: "Google Maps",
    googlePlaceKey: extractGooglePlaceKey(url)
  });

  return trackExtractedLead(lead);
}

async function scrapeJustdial(browser, search) {
  const rawLeads = await runWithFreshPageOnCrash(
    browser,
    {},
    async (page) => {
      const url = getJustdialUrl(search);
      await page.setExtraHTTPHeaders({
        "Accept-Language": "en-US,en;q=0.9",
        Accept: "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
      });
      await navigateWithRetry(
        page,
        url,
        { waitUntil: "domcontentloaded", timeout: 45000 },
        "Justdial navigation"
      );
      await waitForListingContent(page, [
        "div.resultbox",
        ".store-details",
        ".cntanr",
        ".jdresult",
        "article[class*='result']"
      ]);
      await autoScrollPage(page);

      return page.evaluate(() => {
        const clean = (value) => String(value || "").normalize("NFKC").replace(/\s+/g, " ").trim();
        const phoneCandidate = /(?:\+|00)?\d[\d\s().-]{6,}\d/;
        const primaryBoxes = [...document.querySelectorAll("div.resultbox")];
        const fallbackBoxes = [
          ...document.querySelectorAll(
            ".store-details, .cntanr, .jdresult, article[class*='result'], li"
          )
        ].filter((box) => {
          const text = clean(box.innerText || "");
          return text.length > 30 &&
            (box.querySelector(".lng_cont_name, a[href^='tel:']") || phoneCandidate.test(text));
        });
        const boxes = primaryBoxes.length ? primaryBoxes : fallbackBoxes;

        return boxes.map((box) => {
          const text = clean(box.innerText || "");
          const tel = clean(box.querySelector("a[href^='tel:']")?.getAttribute("href")).replace(
            /^tel:/i,
            ""
          );
          return {
            name:
              clean(box.querySelector(".lng_cont_name")?.innerText) ||
              clean(box.querySelector("h2, h3, a[title], [class*='name']")?.innerText) ||
              clean(box.querySelector("a[title]")?.getAttribute("title")),
            phone: tel || text.match(phoneCandidate)?.[0] || "",
            address:
              clean(box.querySelector(".cont_fl_addr")?.innerText) ||
              clean(box.querySelector("[class*='addr'], [class*='loc']")?.innerText)
          };
        });
      });
    },
    "Justdial page"
  );

  const leads = normalizeExtractedLeads(rawLeads, {
    category: search.category,
    city: search.city,
    source: "Justdial"
  });
  console.log(`Leads found per source: Justdial ${leads.length}`);
  return leads;
}

async function scrapeIndiaMart(search) {
  const url = `https://dir.indiamart.com/search.mp?ss=${encodeURIComponent(search.text)}`;
  const $ = await fetchHtml(url, "IndiaMART fetch");
  const leads = extractB2BLeads($, search, "IndiaMART", [
    ".card",
    ".lst",
    ".prd-list",
    ".company",
    ".clg",
    "[class*='card']",
    "[class*='list']",
    "[class*='prd']",
    "[class*='seller']"
  ]);

  console.log(`Leads found per source: IndiaMART ${leads.length}`);
  return leads;
}

async function scrapeTradeIndia(search) {
  const url = `https://www.tradeindia.com/search.html?keyword=${encodeURIComponent(search.text)}`;
  const $ = await fetchHtml(url, "TradeIndia fetch");
  const leads = extractB2BLeads($, search, "TradeIndia", [
    ".product-card",
    ".seller-card",
    ".card",
    ".listing",
    "article",
    "[class*='card']",
    "[class*='seller']",
    "[class*='listing']",
    "[class*='product']"
  ]);

  console.log(`Leads found per source: TradeIndia ${leads.length}`);
  return leads;
}

async function scrapeSulekha(browser, search) {
  const rawLeads = await runWithFreshPageOnCrash(
    browser,
    {},
    async (page) => {
      const url = `https://www.sulekha.com/search?keyword=${encodeURIComponent(
        search.category
      )}&location=${encodeURIComponent(`${search.area} ${search.city}`)}`;
      await navigateWithRetry(
        page,
        url,
        { waitUntil: "domcontentloaded", timeout: 50000 },
        "Sulekha navigation"
      );
      await waitForListingContent(page, [
        ".business-card",
        ".listing-card",
        ".list-item",
        "article",
        "[class*='listing']"
      ]);
      await autoScrollPage(page);

      return page.evaluate(() => {
        const clean = (value) => String(value || "").normalize("NFKC").replace(/\s+/g, " ").trim();
        const phoneCandidate = /(?:\+|00)?\d[\d\s().-]{6,}\d/;
        const cards = [
          ...document.querySelectorAll(
            ".business-card, .listing-card, .list-item, .card, article, li, [class*='listing']"
          )
        ];

        return cards.map((card) => {
          const text = clean(card.innerText || "");
          const tel = clean(card.querySelector("a[href^='tel:']")?.getAttribute("href")).replace(
            /^tel:/i,
            ""
          );
          return {
            name: clean(
              card.querySelector(".business-name, .name, .title, h2, h3, a[title]")?.innerText
            ),
            address: clean(
              card.querySelector(".address, .location, .area, [class*='addr']")?.innerText
            ),
            phone: tel || text.match(phoneCandidate)?.[0] || ""
          };
        });
      });
    },
    "Sulekha page"
  );

  const leads = normalizeExtractedLeads(rawLeads, {
    category: search.category,
    city: search.city,
    source: "Sulekha"
  });
  console.log(`Leads found per source: Sulekha ${leads.length}`);
  return leads;
}

function extractB2BLeads($, search, source, selectors) {
  const leads = [];
  const candidateSelector = selectors.join(", ");
  const candidates = $(candidateSelector).toArray();
  const telephoneParents = $("a[href^='tel:']")
    .toArray()
    .map((element) => $(element).closest(candidateSelector).get(0) || $(element).parent().get(0))
    .filter(Boolean);
  const elements = [...new Set([...candidates, ...telephoneParents])];
  const usedPhones = new Set();

  for (const element of elements) {
    if (leads.length >= MAX_RESULTS) break;

    const card = $(element);
    const text = cleanText(card.text());
    const telValue = card.find("a[href^='tel:']").first().attr("href") || "";
    const phone = cleanPhone(telValue || text);
    if (!phone || usedPhones.has(phone)) continue;

    const name =
      firstText(card, [
        ".companyname",
        ".lcname",
        ".company-name",
        ".seller-name",
        ".prd-name",
        ".title",
        "[class*='company']",
        "[class*='seller']",
        "h2",
        "h3",
        "a[title]"
      ]) || inferNameFromText(text, phone);
    const address =
      firstText(card, [
        ".address",
        ".locality",
        ".city",
        ".sellerlocation",
        ".location",
        "[class*='addr']",
        "[class*='loc']"
      ]) || inferAddressFromText(text);

    const lead = normalizeLead({
      name,
      phone,
      address,
      category: search.category,
      city: search.city,
      source
    });
    const trackedLead = trackExtractedLead(lead);
    if (trackedLead) {
      usedPhones.add(phone);
      leads.push(trackedLead);
    }
  }

  return leads;
}

async function fetchHtml(url, label) {
  const response = await retry(
    async () => {
      const result = await httpClient.get(url, {
        headers: {
          "User-Agent": randomUserAgent(),
          "Accept-Language": "en-US,en;q=0.9",
          Accept: "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        }
      });
      const blockReason = detectHtmlBlock(result.data);
      if (blockReason) throw createRateLimitError(`${label} blocked: ${blockReason}`);
      return result;
    },
    { retries: MAX_RETRIES, label }
  );

  return cheerio.load(response.data);
}

async function newPage(browser) {
  const releaseSlot = await pageSemaphore.acquire();
  let page;

  try {
    page = await browser.newPage();
    pageSlotReleases.set(page, releaseSlot);
    page.once("close", () => releaseManagedPageSlot(page));
    await page.setUserAgent(randomUserAgent());
    await page.setExtraHTTPHeaders({ "Accept-Language": "en-US,en;q=0.9" });
    await page.evaluateOnNewDocument(() => {
      Object.defineProperty(navigator, "webdriver", { get: () => false });
    });
    page.setDefaultTimeout(20000);
    page.setDefaultNavigationTimeout(45000);
    await page.setRequestInterception(true);
    page.on("request", handleInterceptedRequest);
    return page;
  } catch (error) {
    if (page) await safeClosePage(page);
    else releaseSlot();
    throw error;
  }
}

function handleInterceptedRequest(request) {
  if (request.isInterceptResolutionHandled()) return;
  const shouldBlock = BLOCKED_RESOURCE_TYPES.has(request.resourceType());
  const action = shouldBlock ? request.abort() : request.continue();

  action.catch((error) => {
    if (!/request is already handled|intercept resolution/i.test(shortError(error))) {
      console.warn(`Request interception warning: ${shortError(error)}`);
    }
  });
}

async function acceptGoogleConsentOnce(page) {
  if (googleConsentHandled) return;
  if (googleConsentHandlingPromise) return googleConsentHandlingPromise;

  googleConsentHandlingPromise = (async () => {
    const clicked = await page.evaluate(() => {
      const candidates = [
        ...document.querySelectorAll("button, input[type='submit'], [role='button']")
      ];
      const button = candidates.find((element) => {
        const value =
          element.innerText ||
          element.value ||
          element.getAttribute("aria-label") ||
          element.textContent ||
          "";
        return /^(accept all|i agree|agree|accept)$/i.test(value.trim());
      });
      if (!button) return false;
      button.click();
      return true;
    });

    if (clicked) {
      await page
        .waitForFunction(
          () =>
            ![...document.querySelectorAll("button, input[type='submit'], [role='button']")].some(
              (element) => {
                const value =
                  element.innerText ||
                  element.value ||
                  element.getAttribute("aria-label") ||
                  element.textContent ||
                  "";
                return /^(accept all|i agree|agree|accept)$/i.test(value.trim());
              }
            ),
          { timeout: 5000 }
        )
        .catch(() => null);
    }

    googleConsentHandled = true;
  })();

  try {
    await googleConsentHandlingPromise;
  } finally {
    googleConsentHandlingPromise = null;
  }
}

async function waitForMapsResults(page) {
  await page.waitForFunction(
    () => document.querySelector("a.hfpxzc, div[role='feed'], h1") !== null,
    { timeout: 30000 }
  );
}

async function waitForLeadFields(page, timeout = DETAIL_WAIT_MS) {
  await page.waitForFunction(
    () => {
      const hasName = [...document.querySelectorAll("h1, [role='heading'][aria-level='1']")].some(
        (element) => (element.innerText || element.textContent || "").trim()
      );
      const hasPhone = Boolean(
        document.querySelector('button[data-item-id^="phone"], a[href^="tel:"]')
      );
      const hasAddress = Boolean(
        document.querySelector('button[data-item-id="address"], [data-item-id="address"]')
      );
      return hasName && (hasPhone || hasAddress);
    },
    { timeout }
  );
}

async function waitForListingContent(page, selectors) {
  try {
    await page.waitForFunction(
      (candidateSelectors) =>
        candidateSelectors.some((selector) => document.querySelector(selector)) ||
        (document.body?.innerText || "").trim().length > 250,
      { timeout: 15000 },
      selectors
    );
  } catch (error) {
    const bodyText = await page.evaluate(() => (document.body?.innerText || "").trim());
    if (!bodyText) throw error;
  }
}

async function extractGoogleMapsDetail(page) {
  return page.evaluate(() => {
    const clean = (value) => String(value || "").normalize("NFKC").replace(/\s+/g, " ").trim();
    const text = (element) => clean(element?.innerText || element?.textContent || "");
    const firstText = (selectors) => {
      for (const selector of selectors) {
        const value = text(document.querySelector(selector));
        if (value) return value;
      }
      return "";
    };
    const firstAttr = (selectors, attribute) => {
      for (const selector of selectors) {
        const value = clean(document.querySelector(selector)?.getAttribute(attribute));
        if (value) return value;
      }
      return "";
    };
    const valueByAriaLabel = (prefix) => {
      const element = [...document.querySelectorAll("button, a")].find((item) =>
        prefix.test(item.getAttribute("aria-label") || "")
      );
      const label = clean(element?.getAttribute("aria-label"));
      return text(element) || label.replace(prefix, "");
    };
    const categoryButton = [...document.querySelectorAll("button")].find((button) => {
      const action = button.getAttribute("jsaction") || "";
      const label = button.getAttribute("aria-label") || "";
      return action.includes("category") || /^category:/i.test(label);
    });
    const bodyPhone =
      (document.body?.innerText || "").match(/(?:\+|00)?\d[\d\s().-]{6,}\d/)?.[0] || "";

    return {
      name:
        firstText(["h1.DUwDvf", "h1", "[role='heading'][aria-level='1']", "[data-attrid='title']"]) ||
        clean(document.title.replace(/\s+-\s+Google Maps$/i, "")),
      phone:
        firstText(['button[data-item-id^="phone"]', 'a[href^="tel:"]']) ||
        firstAttr(['a[href^="tel:"]'], "href").replace(/^tel:/i, "") ||
        valueByAriaLabel(/^phone:\s*/i) ||
        bodyPhone,
      address:
        firstText(['button[data-item-id="address"]', '[data-item-id="address"]']) ||
        valueByAriaLabel(/^address:\s*/i),
      category: text(categoryButton).replace(/^category:\s*/i, "")
    };
  });
}

async function autoScrollMaps(page, targetCount) {
  let previousCount = 0;
  let stableAttempts = 0;

  for (let attempt = 0; attempt < MAX_SCROLL_ATTEMPTS; attempt += 1) {
    const count = await page.evaluate(() => {
      const feed = document.querySelector('div[role="feed"]');
      if (feed) feed.scrollTo(0, feed.scrollHeight);
      return document.querySelectorAll("a.hfpxzc").length;
    });

    if (count >= targetCount) break;

    const changed = await page
      .waitForFunction(
        (oldCount) => {
          const currentCount = document.querySelectorAll("a.hfpxzc").length;
          const text = document.querySelector('div[role="feed"]')?.innerText || "";
          return currentCount > oldCount || /you've reached the end|no more results/i.test(text);
        },
        { timeout: 2500 },
        count
      )
      .then(() => true)
      .catch(() => false);

    stableAttempts = !changed || count === previousCount ? stableAttempts + 1 : 0;
    previousCount = count;
    if (stableAttempts >= 2) break;
  }
}

async function autoScrollPage(page) {
  let previousHeight = 0;
  let stableAttempts = 0;

  for (let attempt = 0; attempt < MAX_SCROLL_ATTEMPTS; attempt += 1) {
    const height = await page.evaluate(() => {
      window.scrollTo(0, document.body?.scrollHeight || 0);
      return document.body?.scrollHeight || 0;
    });

    const changed = await page
      .waitForFunction(
        (oldHeight) => (document.body?.scrollHeight || 0) > oldHeight,
        { timeout: 2000 },
        height
      )
      .then(() => true)
      .catch(() => false);

    stableAttempts = !changed || height === previousHeight ? stableAttempts + 1 : 0;
    previousHeight = height;
    if (stableAttempts >= 2) break;
  }
}

async function detectGoogleBlock(page) {
  const currentUrl = page.url().toLowerCase();
  if (/google\.[^/]+\/sorry|\/sorry\//i.test(currentUrl)) return "Google sorry page";

  const bodyText = await page.evaluate(() =>
    (document.body?.innerText || "").slice(0, 10000).toLowerCase()
  );
  const patterns = [
    ["unusual traffic", "unusual traffic"],
    ["captcha", "captcha"],
    ["too many requests", "too many requests"],
    ["our systems have detected", "automated traffic warning"],
    ["sorry, but your computer", "Google sorry page"]
  ];

  return patterns.find(([needle]) => bodyText.includes(needle))?.[1] || "";
}

function detectHtmlBlock(html) {
  const text = cleanText(html).toLowerCase().slice(0, 20000);
  if (text.includes("too many requests")) return "too many requests";
  if (text.includes("unusual traffic")) return "unusual traffic";
  if (text.includes("captcha")) return "captcha";
  return "";
}

async function navigateWithRetry(page, url, options, label, checkGoogleBlock = false) {
  return retry(
    async () => {
      const response = await page.goto(url, options);
      const status = response?.status();
      if (status && status >= 400) throw createHttpError(status, `${label} returned HTTP ${status}`);

      if (checkGoogleBlock) {
        const blockReason = await detectGoogleBlock(page);
        if (blockReason) throw createRateLimitError(`Google block detected: ${blockReason}`);
      }

      return response;
    },
    {
      retries: MAX_RETRIES,
      label,
      shouldRetry: (error) => isRetryableError(error) && !isPageCrashError(error)
    }
  );
}

async function runWithFreshPageOnCrash(browser, pageOptions, handler, label) {
  const crashRetries = Math.min(MAX_RETRIES, 1);

  for (let attempt = 0; ; attempt += 1) {
    const page = await newPage(browser, pageOptions);
    try {
      return await handler(page);
    } catch (error) {
      if (attempt >= crashRetries || !isPageCrashError(error)) throw error;
      metrics.retries += 1;
      console.warn(`${label} crashed; reopening the page`);
      await sleep(calculateBackoffMs(attempt, false));
    } finally {
      await safeClosePage(page);
    }
  }
}

function normalizeExtractedLeads(rawLeads, context) {
  const leads = [];

  for (const rawLead of rawLeads) {
    if (leads.length >= MAX_RESULTS) break;
    if (!cleanText(rawLead.name) && !cleanText(rawLead.phone)) continue;
    const lead = trackExtractedLead(normalizeLead({ ...rawLead, ...context }));
    if (lead) leads.push(lead);
  }

  return leads;
}

function trackExtractedLead(lead) {
  if (isValidLead(lead)) {
    metrics.validLeads += 1;
    return lead;
  }

  metrics.invalidLeads += 1;
  return null;
}

function mergeLeadsIntoState(state, leads) {
  if (!leads.length) return 0;
  const previousCount = state.rows.length;
  const merged = dedupeLeads([...state.rows, ...leads]);
  state.rows = merged;
  metrics.newUniqueLeads = Math.max(0, state.rows.length - metrics.existingLeadsLoaded);
  return Math.max(0, state.rows.length - previousCount);
}

function loadExistingLeads() {
  const candidates = uniqueValues([EXISTING_LEADS_FILE, CSV_OUTPUT_FILE, OUTPUT_FILE]);

  for (const filePath of candidates) {
    if (!fs.existsSync(filePath)) continue;
    try {
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName] || {}, { defval: "" });
      return {
        rows: rows.map(normalizeLead).filter(isValidLead),
        filePath
      };
    } catch (error) {
      console.warn(`Could not load ${filePath}: ${shortError(error)}`);
    }
  }

  return { rows: [], filePath: "" };
}

function dedupeLeads(leads) {
  const unique = [];
  let indexes = createLeadIndexes();

  for (const input of leads) {
    const lead = normalizeLead(input);
    if (!isValidLead(lead)) continue;

    const matchingIndexes = new Set();
    for (const key of getDedupeKeys(lead)) {
      const index = indexes.get(key);
      if (index !== undefined && unique[index]) matchingIndexes.add(index);
    }

    if (matchingIndexes.size === 0) {
      const index = unique.length;
      unique.push(lead);
      addLeadToIndexes(indexes, lead, index);
      continue;
    }

    const [primaryIndex, ...secondaryIndexes] = [...matchingIndexes].sort((a, b) => a - b);
    let merged = mergeLeadRecords(unique[primaryIndex], lead);
    for (const index of secondaryIndexes) {
      merged = mergeLeadRecords(merged, unique[index]);
      unique[index] = null;
    }
    unique[primaryIndex] = merged;

    if (secondaryIndexes.length > 0) {
      indexes = createLeadIndexes();
      unique.forEach((item, index) => {
        if (item) addLeadToIndexes(indexes, item, index);
      });
    } else {
      addLeadToIndexes(indexes, merged, primaryIndex);
    }
  }

  return unique.filter(Boolean);
}

function createLeadIndexes() {
  return new Map();
}

function addLeadToIndexes(indexes, lead, index) {
  for (const key of getDedupeKeys(lead)) indexes.set(key, index);
}

function getDedupeKeys(lead) {
  const normalized = normalizeLead(lead);
  const keys = normalized.googlePlaceKeys.map((key) => `google:${key}`);
  if (normalized.phone) {
    keys.push(`phone:${normalized.phone}`);
    return keys;
  }

  const nameKey = normalizeKeyText(normalized.name);
  const addressKey = normalizeKeyText(normalized.address);
  const cityKey = normalizeKeyText(normalized.city);
  if (nameKey && addressKey) keys.push(`name-address:${nameKey}|${addressKey}`);
  if (nameKey && cityKey) keys.push(`name-city:${nameKey}|${cityKey}`);
  return keys;
}

function mergeLeadRecords(existing, incoming) {
  const first = normalizeLead(existing);
  const second = normalizeLead(incoming);
  const firstCompleteness = leadCompleteness(first);
  const secondCompleteness = leadCompleteness(second);
  const preferred = secondCompleteness > firstCompleteness ? second : first;
  const other = preferred === first ? second : first;

  return normalizeLead({
    name: chooseBusinessName(preferred.name, other.name),
    phone: preferred.phone || other.phone,
    address: chooseLongerText(first.address, second.address),
    category: combineListValues(first.category, second.category),
    city: preferred.city || other.city,
    source: preferred.source || other.source,
    googlePlaceKeys: [...first.googlePlaceKeys, ...second.googlePlaceKeys]
  });
}

function leadCompleteness(lead) {
  return [lead.name, lead.phone, lead.address, lead.category, lead.city, lead.source].filter(Boolean)
    .length;
}

function chooseBusinessName(preferred, other) {
  if (!preferred) return other;
  if (!other) return preferred;
  return preferred.length >= 3 ? preferred : other;
}

function chooseLongerText(first, second) {
  if (!first) return second;
  if (!second) return first;
  return second.length > first.length ? second : first;
}

function combineListValues(first, second) {
  const values = `${first || ""};${second || ""}`
    .split(/\s*;\s*/)
    .map(cleanText)
    .filter(Boolean);
  const seen = new Set();
  return values
    .filter((value) => {
      const key = normalizeKeyText(value);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .join("; ");
}

function saveCheckpoint(state, reason) {
  state.rows = dedupeLeads(state.rows);
  writeLeadFiles(state.rows, { writeCsv: true, writeExcel: true });
  state.lastCheckpointAt = state.processedSearches;
  metrics.totalUniqueLeads = state.rows.length;
  console.log(
    `Checkpoint saved ${state.rows.length} unique leads (${reason}) to ${CSV_OUTPUT_FILE} and ${OUTPUT_FILE}`
  );
}

function writeLeadFiles(rows, options = {}) {
  const { writeCsv = true, writeExcel = true } = options;
  const outputRows = rows.map(toOutputRow);
  const worksheet = XLSX.utils.json_to_sheet(outputRows, { header: OUTPUT_COLUMNS });

  if (writeCsv) safeWriteCsv(worksheet, CSV_OUTPUT_FILE);
  if (writeExcel) {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Leads");
    safeWriteWorkbook(workbook, OUTPUT_FILE);
  }
}

function toOutputRow(lead) {
  const normalized = normalizeLead(lead);
  return Object.fromEntries(OUTPUT_COLUMNS.map((column) => [column, normalized[column] || ""]));
}

function safeWriteCsv(worksheet, filePath) {
  const csv = `\uFEFF${XLSX.utils.sheet_to_csv(worksheet)}`;
  const temporaryPath = createTemporaryOutputPath(filePath);
  try {
    fs.writeFileSync(temporaryPath, csv, "utf8");
    replaceOutputFile(temporaryPath, filePath);
  } catch (error) {
    tryRemoveFile(temporaryPath);
    throw error;
  }
}

function safeWriteWorkbook(workbook, filePath) {
  const temporaryPath = createTemporaryOutputPath(filePath);
  try {
    XLSX.writeFile(workbook, temporaryPath, { bookType: "xlsx" });
    replaceOutputFile(temporaryPath, filePath);
  } catch (error) {
    tryRemoveFile(temporaryPath);
    throw error;
  }
}

function createTemporaryOutputPath(filePath) {
  const extension = path.extname(filePath);
  const base = filePath.slice(0, -extension.length);
  return `${base}.tmp-${process.pid}-${Date.now()}${extension}`;
}

function replaceOutputFile(temporaryPath, filePath) {
  try {
    fs.renameSync(temporaryPath, filePath);
  } catch (error) {
    if (!["EBUSY", "EPERM", "EACCES", "EEXIST"].includes(error.code)) {
      tryRemoveFile(temporaryPath);
      throw error;
    }

    const extension = path.extname(filePath);
    const fallbackPath = `${filePath.slice(0, -extension.length)}_new${extension}`;
    tryRemoveFile(fallbackPath);
    fs.renameSync(temporaryPath, fallbackPath);
    console.warn(`${filePath} is locked; wrote ${fallbackPath} instead`);
  }
}

function tryRemoveFile(filePath) {
  try {
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
  } catch (error) {
    // The caller will surface a later write/rename error when the path remains unavailable.
  }
}

function normalizeLead(lead = {}) {
  const googlePlaceKeys = uniqueValues([
    ...(Array.isArray(lead.googlePlaceKeys) ? lead.googlePlaceKeys : []),
    lead.googlePlaceKey
  ]);

  return {
    name: cleanText(lead.name),
    phone: cleanPhone(lead.phone),
    address: cleanGoogleField(lead.address, "address"),
    category: cleanGoogleField(lead.category, "category"),
    city: cleanText(lead.city || CITY),
    source: cleanText(lead.source),
    googlePlaceKeys
  };
}

function isValidLead(lead) {
  return Boolean(cleanText(lead.name).length >= 2 && isValidPhone(lead.phone));
}

function cleanText(value) {
  if (value == null) return "";

  return String(value)
    .normalize("NFKC")
    .replace(/[\u0000-\u001F\u007F]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanPhone(value) {
  if (value == null) return "";
  const text = String(value).replace(/(?:ext\.?|extension|x)\s*\d+.*$/i, " ");
  const candidates = text.match(PHONE_CANDIDATE_REGEX) || [];
  if (/^\s*(?:\+|00)?\d[\d\s().-]*\s*$/.test(text)) candidates.unshift(text);

  for (const candidate of candidates) {
    const phone = normalizePhoneCandidate(candidate);
    if (isValidPhone(phone)) return phone;
  }

  return "";
}

function normalizePhoneCandidate(value) {
  const rawValue = String(value || "").trim();
  let digits = rawValue.replace(/\D/g, "");
  if (digits.startsWith("0091")) digits = digits.slice(2);

  if (digits.startsWith("91") && (digits.length === 11 || digits.length === 12)) {
    const nationalNumber = digits.slice(2);
    if (/^[6-9]\d{9}$/.test(nationalNumber)) {
      const formattedCountryStdCode = /(?:\+|00)?91[\s.-]*\(?(?:20|22|33|40|44|79|80)\)?[\s.-]/.test(
        rawValue
      );
      if (formattedCountryStdCode) return `0${nationalNumber}`;
      return nationalNumber;
    }
    if (/^[1-8]\d{8,9}$/.test(nationalNumber)) return `0${nationalNumber}`;
  }

  if (/^0[6-9]\d{9}$/.test(digits)) {
    const hasFormattedStdCode = /^\(?0\d{2,4}\)?[\s.-]/.test(rawValue);
    const hasKnownMetroStdCode = /^0(?:20|22|33|40|44|79|80)/.test(digits);
    if (hasFormattedStdCode || hasKnownMetroStdCode) return digits;
    return digits.slice(1);
  }
  if (/^[6-9]\d{9}$/.test(digits)) return digits;
  if (/^0[1-8]\d{8,9}$/.test(digits)) return digits;
  if (/^[1-5]\d{8,9}$/.test(digits)) return `0${digits}`;
  if (/^[2-5]\d{7}$/.test(digits)) return digits;
  return "";
}

function isValidPhone(phone) {
  const digits = String(phone || "").replace(/\D/g, "");
  return (
    /^[6-9]\d{9}$/.test(digits) ||
    /^0[1-8]\d{8,9}$/.test(digits) ||
    /^[2-5]\d{7}$/.test(digits)
  );
}

function cleanGoogleField(value, label) {
  return cleanText(value).replace(new RegExp(`^${escapeRegExp(label)}\\s*:\\s*`, "i"), "");
}

function normalizeKeyText(value) {
  return cleanText(value)
    .toLocaleLowerCase("en-IN")
    .replace(/[^\p{L}\p{N}]+/gu, "")
    .trim();
}

function firstText(card, selectors) {
  for (const selector of selectors) {
    const element = card.find(selector).first();
    const value = element.text() || element.attr("title");
    if (cleanText(value)) return value;
  }
  return "";
}

function inferNameFromText(text, phone) {
  const rawText = String(text || "");
  const phoneMatch = rawText.match(PHONE_CANDIDATE_REGEX)?.[0] || "";
  const phoneIndex = phoneMatch ? rawText.indexOf(phoneMatch) : -1;
  const source = phoneIndex > 0 ? rawText.slice(0, phoneIndex) : rawText;
  const parts = cleanText(source)
    .split(/ Get Best Price | Contact Supplier | View Mobile Number /i)
    .map(cleanText)
    .filter((part) => part.length >= 3 && part.length <= 100);
  return parts[0] || "";
}

function inferAddressFromText(text) {
  const match = String(text || "").match(
    /(?:Pune|Mumbai|Maharashtra|Delhi|Bengaluru|Bangalore|Chennai|Hyderabad|India)[^|]{0,120}/i
  );
  return cleanText(match?.[0] || "");
}

function getJustdialUrl(search) {
  const rawSearch =
    typeof search === "string"
      ? search
      : search.text || `${search.category} in ${search.area} ${search.city || CITY}`;
  const city = typeof search === "string" ? CITY : search.city || CITY;
  const parsed = parseLocalSearch(rawSearch, city);
  const category = toJustdialCategory(parsed.category);
  const area = toTitleSlug(parsed.area);
  return `https://www.justdial.com/${encodeURIComponent(city)}/${category}-in-${area}`;
}

function parseLocalSearch(searchText, city = CITY) {
  const cleaned = cleanText(searchText).replace(new RegExp(`\\b${escapeRegExp(city)}\\b`, "i"), "");
  const match = cleaned.match(/^(.+?)\s+in\s+(.+)$/i);
  if (!match) return { category: cleaned, area: "" };
  return { category: cleanText(match[1]), area: cleanText(match[2]) };
}

function toJustdialCategory(category) {
  const normalized = normalizeKeyText(category);
  const map = {
    attendance: "Attendance",
    biometric: "Biometric",
    payroll: "Payroll",
    crm: "CRM",
    billing: "Billing",
    fingerprint: "Fingerprint"
  };
  return map[normalized] || toTitleSlug(category);
}

function toTitleSlug(value) {
  return cleanText(value)
    .split(/\s+/)
    .filter(Boolean)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join("-");
}

async function retry(fn, options = {}) {
  const {
    retries = MAX_RETRIES,
    label = "operation",
    shouldRetry = isRetryableError
  } = options;
  let lastError;

  for (let attempt = 0; attempt <= retries; attempt += 1) {
    try {
      return await fn(attempt);
    } catch (error) {
      lastError = error;
      const rateLimited = isRateLimitError(error);
      if (rateLimited) {
        reportRateLimitIfNeeded(error, label);
      }

      const retryLimit = rateLimited ? Math.min(retries, 1) : retries;
      if (attempt >= retryLimit || !shouldRetry(error)) break;

      metrics.retries += 1;
      const backoffMs = calculateBackoffMs(attempt, rateLimited);
      console.warn(
        `Retrying ${label} (${attempt + 1}/${retryLimit}) in ${backoffMs}ms: ${shortError(error)}`
      );
      await sleep(backoffMs);
    }
  }

  throw lastError;
}

function isRetryableError(error) {
  const status = Number(error?.response?.status || error?.status || error?.statusCode || 0);
  const code = String(error?.code || "").toUpperCase();
  const message = shortError(error).toLowerCase();

  if ([403, 408, 425, 429, 500, 502, 503, 504].includes(status)) return true;
  if (
    [
      "ECONNRESET",
      "ECONNREFUSED",
      "EAI_AGAIN",
      "ENETUNREACH",
      "ETIMEDOUT",
      "ECONNABORTED",
      "GOOGLE_RATE_LIMIT",
      "WEBSITE_BLOCK"
    ].includes(code)
  ) {
    return true;
  }

  return /navigation timeout|net::err_|target closed|session closed|connection reset|socket hang up|page crashed|protocol error.*target/.test(
    message
  );
}

function isRateLimitError(error) {
  const status = Number(error?.response?.status || error?.status || error?.statusCode || 0);
  const code = String(error?.code || "").toUpperCase();
  const message = shortError(error).toLowerCase();
  return (
    status === 403 ||
    status === 429 ||
    code === "GOOGLE_RATE_LIMIT" ||
    code === "WEBSITE_BLOCK" ||
    /unusual traffic|captcha|too many requests|rate limit|google sorry|temporarily blocked/.test(message)
  );
}

function reportRateLimitIfNeeded(error, label) {
  if (error.rateLimitReported) return;
  error.rateLimitReported = true;
  metrics.rateLimits += 1;
  console.warn(`Rate limit/captcha detected during ${label}: ${shortError(error)}`);
}

function isPageCrashError(error) {
  const message = shortError(error).toLowerCase();
  return /target closed|session closed|page crashed|protocol error.*target|most likely the page has been closed/.test(
    message
  );
}

function isTimeoutError(error) {
  return error?.name === "TimeoutError" || /timeout/i.test(shortError(error));
}

function calculateBackoffMs(attempt, rateLimited) {
  const normalBackoff = Math.min(1000 * 2 ** attempt, 10000) + Math.floor(Math.random() * 500);
  return rateLimited ? Math.max(5000, normalBackoff * 3) : normalBackoff;
}

function createHttpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

function createRateLimitError(message) {
  const error = new Error(message);
  error.code = "GOOGLE_RATE_LIMIT";
  return error;
}

function createSemaphore(limit) {
  let active = 0;
  const waiters = [];

  function createRelease() {
    let released = false;
    return () => {
      if (released) return;
      released = true;
      const next = waiters.shift();
      if (next) next(createRelease());
      else active -= 1;
    };
  }

  return {
    acquire() {
      if (active < limit) {
        active += 1;
        return Promise.resolve(createRelease());
      }
      return new Promise((resolve) => waiters.push(resolve));
    }
  };
}

function releaseManagedPageSlot(page) {
  const release = pageSlotReleases.get(page);
  if (!release) return;
  pageSlotReleases.delete(page);
  release();
}

async function safeClosePage(page) {
  if (!page) return;
  try {
    if (!page.isClosed()) await page.close();
  } catch (error) {
    if (!isPageCrashError(error)) console.warn(`Page close warning: ${shortError(error)}`);
  } finally {
    releaseManagedPageSlot(page);
  }
}

function readIntegerEnv(name, fallback, min, max) {
  const rawValue = process.env[name];
  if (rawValue == null || rawValue === "") return fallback;
  const value = Number(rawValue);
  if (!Number.isInteger(value) || value < min || value > max) {
    console.warn(`${name}=${rawValue} is invalid; using ${fallback}`);
    return fallback;
  }
  return value;
}

function readBooleanEnv(name, fallback) {
  const rawValue = process.env[name];
  if (rawValue == null || rawValue === "") return fallback;
  if (/^(true|1|yes)$/i.test(rawValue)) return true;
  if (/^(false|0|no)$/i.test(rawValue)) return false;
  console.warn(`${name}=${rawValue} is invalid; using ${fallback}`);
  return fallback;
}

function readListEnv(name, fallback) {
  const rawValue = process.env[name];
  if (rawValue == null || rawValue === "") return [...fallback];

  let values = [];
  try {
    const parsed = JSON.parse(rawValue);
    if (Array.isArray(parsed)) values = parsed;
    else if (typeof parsed === "string") values = parsed.split(/\r?\n|,/);
  } catch (error) {
    values = String(rawValue).split(/\r?\n|,/);
  }

  const cleanedValues = uniqueValues(values);
  if (cleanedValues.length > 0) return cleanedValues;

  console.warn(`${name} is invalid or empty; using default values`);
  return [...fallback];
}

function readPathEnv(name, fallback) {
  return cleanText(process.env[name] || "") || fallback;
}

function resolveBrowserExecutablePath() {
  const configuredPath = cleanText(process.env.PUPPETEER_EXECUTABLE_PATH || "");
  if (configuredPath) {
    if (fs.existsSync(configuredPath)) return configuredPath;
    console.warn("PUPPETEER_EXECUTABLE_PATH does not exist; checking installed browsers");
  }

  const candidates = [
    path.join(process.env.PROGRAMFILES || "", "Google", "Chrome", "Application", "chrome.exe"),
    path.join(
      process.env["PROGRAMFILES(X86)"] || "",
      "Google",
      "Chrome",
      "Application",
      "chrome.exe"
    ),
    path.join(
      process.env.LOCALAPPDATA || "",
      "Google",
      "Chrome",
      "Application",
      "chrome.exe"
    ),
    path.join(process.env.PROGRAMFILES || "", "Microsoft", "Edge", "Application", "msedge.exe"),
    path.join(
      process.env["PROGRAMFILES(X86)"] || "",
      "Microsoft",
      "Edge",
      "Application",
      "msedge.exe"
    ),
    path.join(
      process.env.LOCALAPPDATA || "",
      "Microsoft",
      "Edge",
      "Application",
      "msedge.exe"
    )
  ];

  return candidates.find((candidate) => candidate && fs.existsSync(candidate)) || "";
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function safeDecodeURIComponent(value) {
  try {
    return decodeURIComponent(value);
  } catch (error) {
    return value;
  }
}

function shortError(error) {
  return String(error?.message || error || "Unknown error").split("\n")[0];
}

function randomUserAgent() {
  return userAgents[Math.floor(Math.random() * userAgents.length)];
}

function printMetrics(startedAt) {
  const elapsedMs = Date.now() - startedAt;
  const finishedSearches = metrics.searchesCompleted + metrics.searchesFailed;
  const averageMs = finishedSearches > 0 ? elapsedMs / finishedSearches : 0;

  console.log("Scraping metrics:");
  console.log(`- Total searches planned: ${metrics.searchesPlanned}`);
  console.log(`- Searches completed: ${metrics.searchesCompleted}`);
  console.log(`- Searches failed: ${metrics.searchesFailed}`);
  console.log(`- Total place URLs found: ${metrics.placeUrlsFound}`);
  console.log(`- Duplicate place URLs skipped: ${metrics.duplicatePlacesSkipped}`);
  console.log(`- Detail pages opened: ${metrics.detailPagesOpened}`);
  console.log(`- Valid leads extracted: ${metrics.validLeads}`);
  console.log(`- Invalid leads rejected: ${metrics.invalidLeads}`);
  console.log(`- Existing leads loaded: ${metrics.existingLeadsLoaded}`);
  console.log(`- New unique leads: ${metrics.newUniqueLeads}`);
  console.log(`- Total unique leads: ${metrics.totalUniqueLeads}`);
  console.log(`- Retry count: ${metrics.retries}`);
  console.log(`- Rate-limit count: ${metrics.rateLimits}`);
  console.log(`- Total execution time: ${(elapsedMs / 1000).toFixed(1)}s`);
  console.log(`- Average time per search: ${(averageMs / 1000).toFixed(2)}s`);
  console.log(`Excel generated: ${OUTPUT_FILE}`);
  console.log(`CSV generated: ${CSV_OUTPUT_FILE}`);
}

if (require.main === module) {
  main().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}

module.exports = {
  DEFAULT_SCRAPER_CONFIG,
  cleanPhone,
  cleanText,
  dedupeLeads,
  dedupeSearches,
  extractGooglePlaceKey,
  filterNewGooglePlaceUrls,
  isRetryableError,
  isValidPhone,
  mapWithConcurrency,
  normalizeLead,
  resolveBrowserExecutablePath,
  writeLeadFiles
};
