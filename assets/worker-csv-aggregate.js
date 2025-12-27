/* eslint-disable no-restricted-globals */

let activeParser = null;
let cancelRequested = false;

self.onmessage = async (event) => {
  const message = event.data;
  if (!message || typeof message !== "object") return;

  if (message.type === "cancel") {
    cancelRequested = true;
    try {
      if (activeParser && typeof activeParser.abort === "function") activeParser.abort();
    } catch {}
    return;
  }

  if (message.type !== "parse") return;

  cancelRequested = false;
  activeParser = null;

  try {
    await ensurePapaParse();
    const result = await parseAndAggregate(message.file, message.requiredDimensions || [], message.filter || null);
    if (cancelRequested) return;
    self.postMessage({ type: "result", result });
  } catch (error) {
    if (cancelRequested) return;
    self.postMessage({ type: "error", message: String(error?.message || error) });
  } finally {
    activeParser = null;
  }
};

async function ensurePapaParse() {
  if (typeof self.Papa !== "undefined" && typeof self.Papa.parse === "function") return;
  importScripts("https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js");
  if (typeof self.Papa === "undefined" || typeof self.Papa.parse !== "function") {
    throw new Error("PapaParse failed to load in worker.");
  }
}

function parseAndAggregate(file, requiredDimensions, filter) {
  return new Promise((resolve, reject) => {
    if (!file) {
      reject(new Error("No file provided."));
      return;
    }

    const normalizedFilter = normalizeFilter(filter);
    const totalsByMetric = Object.create(null);
    const dailyTotalsByMetric = Object.create(null);
    const dimensionTotalsByMetric = {
      organization: Object.create(null),
      functionType: Object.create(null)
    };
    const dimensionDailyTotalsByMetric = {
      organization: Object.create(null),
      functionType: Object.create(null)
    };

    const uniquePersons = new Set();
    let rowCount = 0;
    let filteredOutRows = 0;
    let invalidDateRows = 0;
    let blankOrganizationRows = 0;
    let blankFunctionTypeRows = 0;
    let minDate = "";
    let maxDate = "";

    let headerRow = null;
    let normalizedHeaders = null;
    let index = null;
    let metricColumns = [];
    let metricCatalog = [];
    let missingRequired = [];
    let abortedWithError = false;
    let coreMetricKeys = [];

    const PROGRESS_EVERY = 5000;
    let lastProgressAt = 0;

    self.Papa.parse(file, {
      header: false,
      skipEmptyLines: "greedy",
      chunkSize: 1024 * 1024,
      error: (error) => reject(error),
      chunk: (results, parser) => {
        activeParser = parser;
        if (cancelRequested) {
          try {
            parser.abort();
          } catch {}
          return;
        }

        if (!results || !Array.isArray(results.data)) return;
        const rows = results.data;

        for (let i = 0; i < rows.length; i++) {
          if (cancelRequested) {
            try {
              parser.abort();
            } catch {}
            return;
          }

          const row = rows[i];
          if (!row || !Array.isArray(row) || row.length === 0) continue;

          if (!headerRow) {
            headerRow = row.map((v) => String(v ?? ""));
            normalizedHeaders = headerRow.map(normalizeHeader);
            const schema = buildSchema(normalizedHeaders, requiredDimensions);
            index = schema.index;
            missingRequired = schema.missingRequired;
            if (missingRequired.length) {
              abortedWithError = true;
              try {
                parser.abort();
              } catch {}
              reject(new Error("Missing required columns: " + missingRequired.join(", ")));
              return;
            }
            metricColumns = buildMetricColumns(headerRow, normalizedHeaders, schema.dimensionIndexes);
            metricCatalog = metricColumns.catalog;
            metricColumns = metricColumns.columns;
            coreMetricKeys = detectCoreMetricKeys(metricCatalog);
            continue;
          }

          rowCount += 1;

          const metricDateRaw = row[index.metricDateIndex];
          const dateKey = normalizeDateKey(metricDateRaw);
          if (!dateKey) {
            invalidDateRows += 1;
            continue;
          }

          if (!isDateWithinRange(dateKey, normalizedFilter?.dateRange)) {
            filteredOutRows += 1;
            continue;
          }

          if (!minDate || dateKey < minDate) minDate = dateKey;
          if (!maxDate || dateKey > maxDate) maxDate = dateKey;

          const organization = bucketDimension(row[index.organizationIndex], 96);
          const functionType = bucketDimension(row[index.functionTypeIndex], 96);
          if (organization === "(blank)") blankOrganizationRows += 1;
          if (functionType === "(blank)") blankFunctionTypeRows += 1;

          if (!isDimensionAllowed(organization, functionType, normalizedFilter?.dimensionFilter)) {
            filteredOutRows += 1;
            continue;
          }

          const personId = safeText(row[index.personIdIndex], 128);
          if (personId) uniquePersons.add(personId);

          for (let m = 0; m < metricColumns.length; m++) {
            const { metricKey, colIndex } = metricColumns[m];
            const value = parseNumeric(row[colIndex]);
            if (!value) continue;

            totalsByMetric[metricKey] = (totalsByMetric[metricKey] || 0) + value;

            addNestedMetric(dailyTotalsByMetric, dateKey, metricKey, value);
            addNestedMetric(dimensionTotalsByMetric.organization, organization, metricKey, value);
            addNestedMetric(dimensionTotalsByMetric.functionType, functionType, metricKey, value);

            if (coreMetricKeys.includes(metricKey)) {
              addDimensionDailyMetric(dimensionDailyTotalsByMetric.organization, metricKey, dateKey, organization, value);
              addDimensionDailyMetric(dimensionDailyTotalsByMetric.functionType, metricKey, dateKey, functionType, value);
            }
          }

          if (rowCount - lastProgressAt >= PROGRESS_EVERY) {
            lastProgressAt = rowCount;
            self.postMessage({ type: "progress", phase: "Aggregating", rowsProcessed: rowCount });
          }
        }
      },
      complete: () => {
        if (cancelRequested || abortedWithError) return;
        resolve({
          meta: {
            rowCount,
            uniquePersons: uniquePersons.size,
            minDate,
            maxDate,
            metricCount: metricCatalog.length,
            invalidDateRows,
            blankOrganizationRows,
            blankFunctionTypeRows,
            filteredOutRows,
            missingRequired
          },
          metricCatalog,
          coreMetricKeys,
          totalsByMetric,
          dailyTotalsByMetric,
          dimensionTotalsByMetric,
          dimensionDailyTotalsByMetric,
          appliedFilter: normalizedFilter
        });
      }
    });
  });
}

function buildSchema(normalizedHeaders, requiredDimensions) {
  const dimensionIndexes = Object.create(null);
  for (const dim of requiredDimensions) {
    const aliases = Array.isArray(dim.aliases) ? dim.aliases : [];
    const idx = findFirstIndex(normalizedHeaders, aliases);
    dimensionIndexes[dim.key] = idx;
  }

  const missingRequired = [];
  if (dimensionIndexes.personId === -1) missingRequired.push("PersonId");
  if (dimensionIndexes.metricDate === -1) missingRequired.push("MetricDate");
  if (dimensionIndexes.organization === -1) missingRequired.push("Organization");
  if (dimensionIndexes.functionType === -1) missingRequired.push("FunctionType");

  return {
    missingRequired,
    dimensionIndexes,
    index: {
      personIdIndex: dimensionIndexes.personId,
      metricDateIndex: dimensionIndexes.metricDate,
      organizationIndex: dimensionIndexes.organization,
      functionTypeIndex: dimensionIndexes.functionType
    }
  };
}

function findFirstIndex(headers, aliases) {
  const aliasSet = new Set((aliases || []).map(normalizeHeader));
  for (let i = 0; i < headers.length; i++) {
    if (aliasSet.has(headers[i])) return i;
  }
  return -1;
}

function buildMetricColumns(headerRow, normalizedHeaders, dimensionIndexes) {
  const dimensionIndexSet = new Set(
    Object.values(dimensionIndexes)
      .filter((v) => typeof v === "number" && v >= 0)
      .map((v) => v)
  );

  const columns = [];
  const catalog = [];

  for (let i = 0; i < headerRow.length; i++) {
    if (dimensionIndexSet.has(i)) continue;
    const original = String(headerRow[i] ?? "").trim();
    const key = normalizedHeaders[i];
    if (!key) continue;
    catalog.push({
      key,
      label: original || key,
      group: guessGroup(key)
    });
    columns.push({ metricKey: key, colIndex: i });
  }

  return { columns, catalog };
}

function detectCoreMetricKeys(metricCatalog) {
  const wanted = [
    "total copilot actions taken",
    "total copilot active days",
    "total copilot enabled days"
  ];
  const wantedSet = new Set(wanted.map(normalizeHeader));
  const keys = [];
  for (const metric of metricCatalog || []) {
    if (wantedSet.has(metric.key)) keys.push(metric.key);
  }
  return keys;
}

function guessGroup(metricKey) {
  const v = String(metricKey || "");
  if (v.includes("teams")) return "Teams";
  if (v.includes("word")) return "Word";
  if (v.includes("excel")) return "Excel";
  if (v.includes("outlook")) return "Outlook";
  if (v.includes("powerpoint")) return "PowerPoint";
  if (v.includes("chat")) return "Chat";
  if (v.includes("meeting") || v.includes("recap") || v.includes("summar")) return "Meetings";
  return "Other";
}

function addNestedMetric(container, key, metricKey, delta) {
  let bucket = container[key];
  if (!bucket) {
    bucket = Object.create(null);
    container[key] = bucket;
  }
  bucket[metricKey] = (bucket[metricKey] || 0) + delta;
}

function addDimensionDailyMetric(container, metricKey, dateKey, dimValue, delta) {
  let byMetric = container[metricKey];
  if (!byMetric) {
    byMetric = Object.create(null);
    container[metricKey] = byMetric;
  }
  let byDate = byMetric[dateKey];
  if (!byDate) {
    byDate = Object.create(null);
    byMetric[dateKey] = byDate;
  }
  byDate[dimValue] = (byDate[dimValue] || 0) + delta;
}

function parseNumeric(value) {
  if (value === null || typeof value === "undefined") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const s = String(value).trim();
  if (!s) return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function normalizeDateKey(value) {
  const s = String(value ?? "").trim();
  if (!s) return "";
  const candidate = s.length >= 10 ? s.slice(0, 10) : s;
  if (/^\d{4}-\d{2}-\d{2}$/.test(candidate)) return candidate;
  const parsed = new Date(s);
  if (Number.isNaN(parsed.getTime())) return "";
  return parsed.toISOString().slice(0, 10);
}

function bucketDimension(value, maxLen) {
  const text = safeText(value, maxLen);
  return text ? text : "(blank)";
}

function safeText(value, maxLen) {
  const raw = String(value ?? "").replace(/[\r\n\t]+/g, " ").trim();
  if (!raw) return "";
  const cleaned = raw.replace(/\s+/g, " ");
  if (cleaned.length <= maxLen) return cleaned;
  return cleaned.slice(0, maxLen - 1) + "…";
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ");
}

function normalizeFilter(filter) {
  if (!filter || typeof filter !== "object") return null;

  const start = normalizeDateKey(filter.startDate || "");
  const end = normalizeDateKey(filter.endDate || "");
  const dimension = filter.dimension === "functionType" ? "functionType" : filter.dimension === "organization" ? "organization" : null;
  const values = Array.isArray(filter.allowedValues) ? filter.allowedValues.map((v) => safeText(v, 96)).filter(Boolean) : [];
  const allowedSet = values.length ? new Set(values) : null;

  return {
    dateRange: start || end ? { start, end } : null,
    dimensionFilter: dimension && allowedSet ? { dimension, allowed: allowedSet } : null
  };
}

function isDateWithinRange(dateKey, range) {
  if (!range) return true;
  if (range.start && dateKey < range.start) return false;
  if (range.end && dateKey > range.end) return false;
  return true;
}

function isDimensionAllowed(organization, functionType, filter) {
  if (!filter) return true;
  if (!filter.allowed || filter.allowed.size === 0) return true;
  if (filter.dimension === "organization") return filter.allowed.has(organization);
  if (filter.dimension === "functionType") return filter.allowed.has(functionType);
  return true;
}
