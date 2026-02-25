const DATA_PATH = "./data/t2_master.json";
const URL_MAX = 1800;

const dom = {
  banner: document.getElementById("banner"),
  summary: document.getElementById("summary"),
  sheetSelect: document.getElementById("sheetSelect"),
  countrySelect: document.getElementById("countrySelect"),
  sheetFilter: document.getElementById("sheetFilter"),
  countryFilter: document.getElementById("countryFilter"),
  commandInput: document.getElementById("commandInput"),
  commandApplyBtn: document.getElementById("commandApplyBtn"),
  suggestions: document.getElementById("suggestions"),
  selectFilteredSheetsBtn: document.getElementById("selectFilteredSheetsBtn"),
  clearSheetsBtn: document.getElementById("clearSheetsBtn"),
  selectFilteredCountriesBtn: document.getElementById("selectFilteredCountriesBtn"),
  clearCountriesBtn: document.getElementById("clearCountriesBtn"),
  undoBtn: document.getElementById("undoBtn"),
  clearAllBtn: document.getElementById("clearAllBtn"),
  seriesManager: document.getElementById("seriesManager"),
  emptyState: document.getElementById("emptyState"),
  chartCanvas: document.getElementById("seriesChart"),
  rangeButtons: Array.from(document.querySelectorAll(".rangeBtn")),
  axisButtons: Array.from(document.querySelectorAll(".axisBtn")),
};

const runtimeTabId = (window.crypto && crypto.randomUUID && crypto.randomUUID()) || String(Date.now());

let workbook;
let chart;
let allSheets = [];
let allCountries = [];
let sheetToCountries = new Map();
let seriesCache = new Map();
let dateDomain = { min: null, max: null };

const state = {
  selectedSheets: new Set(),
  selectedCountries: new Set(),
  hiddenSeries: new Set(),
  hiddenSeriesOrder: [],
  activeRange: "all",
  axisMode: "raw",
  undoStack: [],
  isSharePartial: false,
};

let sheetFilterTerm = "";
let countryFilterTerm = "";
let commandDebounce;
let urlDebounce;

function normalizeToken(text) {
  return String(text)
    .toLowerCase()
    .replace(/[^\w\s/]+/g, " ")
    .replace(/[\/]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function canonicalizeQuery(text) {
  return normalizeToken(text)
    .replace(/\btraining\b/g, "trailing")
    .replace(/\bp e\b/g, "pe")
    .replace(/\bp\/e\b/g, "pe");
}

function strictIntList(value, maxItems) {
  if (!value) return [];
  const out = [];
  const parts = value.split(",");
  for (const p of parts) {
    if (!/^\d+$/.test(p)) continue;
    const n = Number(p);
    if (Number.isInteger(n)) {
      out.push(n);
      if (out.length >= maxItems) break;
    }
  }
  return out;
}

function setBanner(message = "", warn = false) {
  dom.banner.textContent = message;
  dom.banner.classList.toggle("show", Boolean(message));
  if (warn) dom.banner.classList.add("warn");
  else dom.banner.classList.remove("warn");
}

function addUndoSnapshot() {
  const snap = {
    selectedSheets: Array.from(state.selectedSheets),
    selectedCountries: Array.from(state.selectedCountries),
    hiddenSeries: Array.from(state.hiddenSeries),
    hiddenSeriesOrder: [...state.hiddenSeriesOrder],
    activeRange: state.activeRange,
    axisMode: state.axisMode,
  };
  state.undoStack.push(snap);
  if (state.undoStack.length > 3) state.undoStack.shift();
}

function restoreSnapshot(snap) {
  if (!snap) return;
  state.selectedSheets = new Set(snap.selectedSheets);
  state.selectedCountries = new Set(snap.selectedCountries);
  state.hiddenSeries = new Set(snap.hiddenSeries);
  state.hiddenSeriesOrder = [...snap.hiddenSeriesOrder];
  state.activeRange = snap.activeRange;
  state.axisMode = snap.axisMode;
  cascadeAndPrune(false);
  syncUIAndRender();
}

function getSeriesKey(sheet, country) {
  return `${sheet}|||${country}`;
}

function hashCode(text) {
  let hash = 0;
  for (let i = 0; i < text.length; i += 1) {
    hash = (hash << 5) - hash + text.charCodeAt(i);
    hash |= 0;
  }
  return Math.abs(hash);
}

function hslColor(text) {
  const h = hashCode(text) % 360;
  const s = 55 + (hashCode(text + "s") % 15);
  const l = 35 + (hashCode(text + "l") % 20);
  return `hsl(${h} ${s}% ${l}%)`;
}

function buildIndexes() {
  allSheets = Object.keys(workbook.sheets).sort((a, b) => a.localeCompare(b));
  sheetToCountries = new Map();
  seriesCache = new Map();
  const countrySet = new Set();
  let minMs = Infinity;
  let maxMs = -Infinity;

  for (const sheet of allSheets) {
    const ws = workbook.sheets[sheet];
    const countries = Array.from(new Set(ws.countries)).sort((a, b) => a.localeCompare(b));
    sheetToCountries.set(sheet, new Set(countries));
    countries.forEach((c) => countrySet.add(c));

    for (const country of countries) {
      const points = [];
      for (const row of ws.rows) {
        const raw = row.values[country];
        if (raw === null || raw === undefined || Number.isNaN(raw)) continue;
        const ms = Date.parse(row.date);
        if (Number.isNaN(ms)) continue;
        points.push({ date: row.date, ms, value: Number(raw) });
        if (ms < minMs) minMs = ms;
        if (ms > maxMs) maxMs = ms;
      }
      seriesCache.set(getSeriesKey(sheet, country), points);
    }
  }

  allCountries = Array.from(countrySet).sort((a, b) => a.localeCompare(b));
  dateDomain = {
    min: Number.isFinite(minMs) ? minMs : null,
    max: Number.isFinite(maxMs) ? maxMs : null,
  };
}

function getValidCountryUnion(selectedSheets) {
  const union = new Set();
  selectedSheets.forEach((sheet) => {
    const set = sheetToCountries.get(sheet);
    if (!set) return;
    set.forEach((c) => union.add(c));
  });
  return union;
}

function cascadeAndPrune(notify) {
  const validCountries = getValidCountryUnion(state.selectedSheets);
  let prunedCountries = 0;
  for (const c of Array.from(state.selectedCountries)) {
    if (!validCountries.has(c)) {
      state.selectedCountries.delete(c);
      prunedCountries += 1;
    }
  }

  const validHidden = new Set();
  const newOrder = [];
  let prunedHidden = 0;
  for (const key of state.hiddenSeriesOrder) {
    const [sheet, country] = key.split("|||");
    if (state.selectedSheets.has(sheet) && state.selectedCountries.has(country)) {
      validHidden.add(key);
      newOrder.push(key);
    } else {
      prunedHidden += 1;
    }
  }
  state.hiddenSeries = validHidden;
  state.hiddenSeriesOrder = newOrder;

  if (state.selectedCountries.size === 0 && validCountries.size > 0) {
    state.selectedCountries.add(Array.from(validCountries).sort((a, b) => a.localeCompare(b))[0]);
  }

  if (notify && (prunedCountries > 0 || prunedHidden > 0)) {
    setBanner(`Removed ${prunedCountries} invalid countries and ${prunedHidden} stale hidden series.`);
  }
}

function applyDefaultState() {
  const defaultSheet = allSheets.includes("Trailing PE") ? "Trailing PE" : allSheets[0];
  if (defaultSheet) state.selectedSheets.add(defaultSheet);

  const defaultCountryPool = Array.from(getValidCountryUnion(state.selectedSheets));
  const defaultCountry = defaultCountryPool.includes("India")
    ? "India"
    : defaultCountryPool.sort((a, b) => a.localeCompare(b))[0];
  if (defaultCountry) state.selectedCountries.add(defaultCountry);
}

function applyHydratedState(payload) {
  if (!payload) return false;
  const used = { sheets: 0, countries: 0 };

  if (payload.sheets) {
    payload.sheets.forEach((idx) => {
      if (idx >= 0 && idx < allSheets.length) {
        state.selectedSheets.add(allSheets[idx]);
        used.sheets += 1;
      }
    });
  }

  if (payload.countries) {
    payload.countries.forEach((idx) => {
      if (idx >= 0 && idx < allCountries.length) {
        state.selectedCountries.add(allCountries[idx]);
        used.countries += 1;
      }
    });
  }

  if (payload.range && ["all", "10y", "5y", "3y", "1y"].includes(payload.range)) {
    state.activeRange = payload.range;
  }
  if (payload.axis && ["raw", "indexed", "zscore"].includes(payload.axis)) {
    state.axisMode = payload.axis;
  }

  if (payload.hidden) {
    payload.hidden.forEach((pair) => {
      const [sIdx, cIdx] = pair;
      if (sIdx < 0 || cIdx < 0 || sIdx >= allSheets.length || cIdx >= allCountries.length) return;
      const key = getSeriesKey(allSheets[sIdx], allCountries[cIdx]);
      state.hiddenSeries.add(key);
      state.hiddenSeriesOrder.push(key);
    });
  }

  state.isSharePartial = Boolean(payload.partial);
  cascadeAndPrune(false);
  return used.sheets > 0;
}

function hydrateFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const sheetIdx = strictIntList(params.get("s"), 80);
  const countryIdx = strictIntList(params.get("c"), 80);
  const hiddenRaw = params.get("h") || "";
  const hiddenPairs = hiddenRaw
    .split(".")
    .slice(0, 500)
    .map((token) => {
      const [a, b] = token.split("-");
      if (!/^\d+$/.test(a || "") || !/^\d+$/.test(b || "")) return null;
      return [Number(a), Number(b)];
    })
    .filter(Boolean);

  return {
    sheets: sheetIdx,
    countries: countryIdx,
    range: params.get("r"),
    axis: params.get("a"),
    hidden: hiddenPairs,
    partial: params.get("partial") === "1",
  };
}

function hydrateFromStorage() {
  try {
    const raw = localStorage.getItem("t2viz:last");
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.sheets) || !Array.isArray(parsed.countries)) return null;
    return parsed;
  } catch {
    return null;
  }
}

function persistToStorage(payload) {
  try {
    const now = Date.now();
    localStorage.setItem("t2viz:last", JSON.stringify(payload));
    localStorage.setItem(`t2viz:tab:${runtimeTabId}`, JSON.stringify({ ...payload, ts: now }));
    const tabKeys = Object.keys(localStorage).filter((k) => k.startsWith("t2viz:tab:"));
    const items = tabKeys
      .map((k) => {
        try {
          return { key: k, ts: JSON.parse(localStorage.getItem(k) || "{}").ts || 0 };
        } catch {
          return { key: k, ts: 0 };
        }
      })
      .sort((a, b) => b.ts - a.ts);
    items.slice(10).forEach((x) => localStorage.removeItem(x.key));
  } catch {
    // non-fatal
  }
}

function updateUrlFromState() {
  const params = new URLSearchParams();
  const sheetIndices = Array.from(state.selectedSheets)
    .map((sheet) => allSheets.indexOf(sheet))
    .filter((x) => x >= 0)
    .sort((a, b) => a - b)
    .slice(0, 80);
  const countryIndices = Array.from(state.selectedCountries)
    .map((country) => allCountries.indexOf(country))
    .filter((x) => x >= 0)
    .sort((a, b) => a - b)
    .slice(0, 80);

  params.set("s", sheetIndices.join(","));
  params.set("c", countryIndices.join(","));
  params.set("r", state.activeRange);
  params.set("a", state.axisMode);

  let hiddenTokens = state.hiddenSeriesOrder
    .filter((key) => state.hiddenSeries.has(key))
    .map((key) => {
      const [sheet, country] = key.split("|||");
      const s = allSheets.indexOf(sheet);
      const c = allCountries.indexOf(country);
      if (s < 0 || c < 0) return null;
      return `${s}-${c}`;
    })
    .filter(Boolean)
    .slice(0, 500);

  params.set("h", hiddenTokens.join("."));
  params.delete("partial");

  let query = params.toString();
  let partial = false;
  while (query.length > URL_MAX && hiddenTokens.length > 0) {
    hiddenTokens.pop();
    params.set("h", hiddenTokens.join("."));
    partial = true;
    query = params.toString();
  }

  if (partial) {
    params.set("partial", "1");
    query = params.toString();
    state.isSharePartial = true;
    setBanner("Share link trimmed: hidden-series state is partial.", true);
  } else {
    state.isSharePartial = false;
  }

  const url = `${window.location.pathname}?${query}`;
  window.history.replaceState({}, "", url);

  persistToStorage({
    sheets: sheetIndices,
    countries: countryIndices,
    range: state.activeRange,
    axis: state.axisMode,
    hidden: hiddenTokens.map((token) => token.split("-").map(Number)),
    partial,
  });
}

function scheduleUrlSync() {
  clearTimeout(urlDebounce);
  urlDebounce = setTimeout(updateUrlFromState, 300);
}

function buildSelectOptions(selectEl, values, selectedSet) {
  selectEl.innerHTML = "";
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    option.selected = selectedSet.has(value);
    selectEl.appendChild(option);
  });
}

function getFilteredSheets() {
  if (!sheetFilterTerm) return allSheets;
  return allSheets.filter((s) => normalizeToken(s).includes(sheetFilterTerm));
}

function getFilteredCountries() {
  const valid = getValidCountryUnion(state.selectedSheets);
  const base = Array.from(valid).sort((a, b) => a.localeCompare(b));
  if (!countryFilterTerm) return base;
  return base.filter((c) => normalizeToken(c).includes(countryFilterTerm));
}

function levenshtein(a, b) {
  const m = a.length;
  const n = b.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i += 1) dp[i][0] = i;
  for (let j = 0; j <= n; j += 1) dp[0][j] = j;
  for (let i = 1; i <= m; i += 1) {
    for (let j = 1; j <= n; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
    }
  }
  return dp[m][n];
}

function fuzzyScore(query, candidate) {
  const q = normalizeToken(query);
  const c = normalizeToken(candidate);
  if (!q || !c) return 0;
  if (c.includes(q) || q.includes(c)) return 1;
  const d = levenshtein(q, c);
  const maxLen = Math.max(q.length, c.length);
  return maxLen === 0 ? 0 : 1 - d / maxLen;
}

function parseCommand(text) {
  const query = canonicalizeQuery(text);
  if (!query) return { match: null, suggestions: [] };

  const countryCandidates = allCountries.map((country) => {
    const c = normalizeToken(country);
    const exact = query.includes(c);
    return { country, token: c, score: exact ? c.length * 10 : 0 };
  });

  const sheetCandidates = allSheets.map((sheet) => {
    const s = normalizeToken(sheet);
    const exact = query.includes(s);
    return { sheet, token: s, score: exact ? s.length * 10 : 0 };
  });

  let best = null;
  for (const c of countryCandidates) {
    for (const s of sheetCandidates) {
      const cWords = c.token.split(" ");
      const sWords = s.token.split(" ");
      const overlap = cWords.filter((w) => sWords.includes(w)).length;
      const score = c.score + s.score - overlap;
      if (!best || score > best.score || (score === best.score && s.token.length > best.sheetTokenLen)) {
        best = { country: c.country, sheet: s.sheet, score, sheetTokenLen: s.token.length };
      }
    }
  }

  if (best && best.score > 0) {
    return { match: { country: best.country, sheet: best.sheet, fuzzy: false }, suggestions: [] };
  }

  const fuzzyPairs = [];
  for (const country of allCountries) {
    const cScore = fuzzyScore(query, country);
    if (cScore < 0.82) continue;
    for (const sheet of allSheets) {
      const sScore = fuzzyScore(query, sheet);
      const combined = (cScore + sScore) / 2;
      if (combined >= 0.82) {
        fuzzyPairs.push({ country, sheet, score: combined });
      }
    }
  }
  fuzzyPairs.sort((a, b) => b.score - a.score);

  return { match: null, suggestions: fuzzyPairs.slice(0, 3) };
}

function applySelection(sheet, country) {
  addUndoSnapshot();
  state.selectedSheets = new Set([sheet]);
  state.selectedCountries = new Set([country]);
  cascadeAndPrune(false);
  syncUIAndRender();
}

function getRangeStartMs(maxMs, range) {
  if (range === "all" || !maxMs) return null;
  const years = { "10y": 10, "5y": 5, "3y": 3, "1y": 1 }[range];
  if (!years) return null;
  const d = new Date(maxMs);
  d.setFullYear(d.getFullYear() - years);
  return d.getTime();
}

function computeSeriesForRender() {
  const snapshotSheets = Array.from(state.selectedSheets);
  const snapshotCountries = Array.from(state.selectedCountries);
  const datasets = [];
  const managerRows = [];
  let totalVisiblePoints = 0;
  const startMs = getRangeStartMs(dateDomain.max, state.activeRange);

  snapshotSheets.forEach((sheet) => {
    snapshotCountries.forEach((country) => {
      const key = getSeriesKey(sheet, country);
      const raw = seriesCache.get(key) || [];
      const filtered = raw.filter((p) => (startMs ? p.ms >= startMs : true));

      if (filtered.length === 0) {
        managerRows.push({ key, label: `${country} - ${sheet}`, color: hslColor(key), status: "No data", visible: false });
        return;
      }

      let values = filtered.map((p) => p.value);

      if (state.axisMode === "indexed") {
        const first = filtered.find((p) => p.value !== null && p.value !== 0);
        if (!first) {
          managerRows.push({ key, label: `${country} - ${sheet}`, color: hslColor(key), status: "Not indexable", visible: false });
          return;
        }
        values = filtered.map((p) => (p.value / first.value) * 100);
      }

      if (state.axisMode === "zscore") {
        const nonNull = values.filter((v) => v !== null && v !== undefined && Number.isFinite(v));
        if (nonNull.length < 2) {
          managerRows.push({ key, label: `${country} - ${sheet}`, color: hslColor(key), status: "Not normalizable", visible: false });
          return;
        }
        const mean = nonNull.reduce((a, b) => a + b, 0) / nonNull.length;
        const variance = nonNull.reduce((a, b) => a + (b - mean) ** 2, 0) / nonNull.length;
        const std = Math.sqrt(variance);
        if (std === 0) {
          managerRows.push({ key, label: `${country} - ${sheet}`, color: hslColor(key), status: "Not normalizable", visible: false });
          return;
        }
        values = values.map((v) => (v - mean) / std);
      }

      const data = filtered.map((p, i) => ({ x: p.ms, y: values[i] }));
      const hidden = state.hiddenSeries.has(key);
      managerRows.push({ key, label: `${country} - ${sheet}`, color: hslColor(key), status: "", visible: !hidden });

      if (!hidden) {
        datasets.push({
          key,
          label: `${country} - ${sheet}`,
          data,
          borderColor: hslColor(key),
          pointRadius: 0,
          pointHoverRadius: 3,
          borderWidth: 2,
          tension: 0.1,
          fill: false,
        });
        totalVisiblePoints += data.length;
      }
    });
  });

  return { datasets, managerRows, totalVisiblePoints, visibleSeriesCount: datasets.length };
}

function enforceGuardrails(visibleSeriesCount, totalVisiblePoints) {
  const warn =
    (visibleSeriesCount > 50 && visibleSeriesCount <= 80) ||
    (totalVisiblePoints > 100000 && totalVisiblePoints <= 200000);
  const block = visibleSeriesCount > 80 || totalVisiblePoints > 200000;

  if (block) {
    setBanner("Selection too large to render. Narrow sheets/countries or date range.", true);
    return false;
  }
  if (warn) {
    const ok = window.confirm(
      `Large render: ${visibleSeriesCount} series / ${totalVisiblePoints.toLocaleString()} points. Render all?`
    );
    if (!ok) return false;
  }
  return true;
}

function updateSummary(info) {
  const text = [
    `${state.selectedSheets.size} sheet(s)`,
    `${state.selectedCountries.size} country(s)`,
    `${info.visibleSeriesCount} visible series`,
    `${info.totalVisiblePoints.toLocaleString()} points`,
    `Range: ${state.activeRange.toUpperCase()}`,
    `Axis: ${state.axisMode}`,
  ].join(" | ");
  dom.summary.textContent = text;
}

function renderSeriesManager(rows) {
  dom.seriesManager.innerHTML = "";
  rows.forEach((row) => {
    const wrap = document.createElement("div");
    wrap.className = "seriesItem";

    const swatch = document.createElement("span");
    swatch.className = "swatch";
    swatch.style.background = row.color;

    const label = document.createElement("span");
    label.textContent = row.label;

    const status = document.createElement("span");
    status.className = `statusTag ${row.status ? "warn" : ""}`;
    status.textContent = row.status || "";

    const toggle = document.createElement("input");
    toggle.type = "checkbox";
    toggle.checked = row.visible;
    toggle.disabled = Boolean(row.status);
    toggle.addEventListener("change", () => {
      if (toggle.checked) {
        state.hiddenSeries.delete(row.key);
      } else if (!state.hiddenSeries.has(row.key)) {
        state.hiddenSeries.add(row.key);
        state.hiddenSeriesOrder.push(row.key);
      }
      syncUIAndRender();
    });

    wrap.appendChild(swatch);
    wrap.appendChild(label);
    wrap.appendChild(status);
    wrap.appendChild(toggle);
    dom.seriesManager.appendChild(wrap);
  });
}

function getScaleRatio(datasets) {
  const ranges = datasets
    .map((ds) => {
      const ys = ds.data.map((p) => p.y).filter(Number.isFinite);
      if (!ys.length) return null;
      return Math.max(...ys) - Math.min(...ys);
    })
    .filter((v) => v && v > 0);
  if (ranges.length < 2) return 1;
  const minPos = Math.min(...ranges);
  const max = Math.max(...ranges);
  if (!minPos) return 1;
  return max / minPos;
}

function ensureChart() {
  if (!window.Chart) {
    dom.emptyState.classList.add("show");
    dom.emptyState.textContent = "Chart library failed to load.";
    return null;
  }
  if (!chart) {
    chart = new Chart(dom.chartCanvas, {
      type: "line",
      data: { datasets: [] },
      options: {
        parsing: false,
        responsive: true,
        maintainAspectRatio: false,
        animation: false,
        normalized: true,
        interaction: { mode: "nearest", intersect: false },
        plugins: {
          legend: { display: false },
          tooltip: {
            mode: "nearest",
            intersect: false,
            callbacks: {
              title: (items) => {
                if (!items || items.length === 0) return "";
                const ms = Number(items[0].parsed.x);
                if (!Number.isFinite(ms)) return "";
                return new Date(ms).toISOString().slice(0, 10);
              },
              label: (ctx) => {
                const val = Number(ctx.parsed.y);
                return `${ctx.dataset.label}: ${val.toLocaleString(undefined, { maximumFractionDigits: 4 })}`;
              },
            },
          },
        },
        scales: {
          x: {
            type: "linear",
            ticks: {
              maxTicksLimit: 10,
              callback: (v) => {
                const ms = Number(v);
                if (!Number.isFinite(ms)) return "";
                return String(new Date(ms).getFullYear());
              },
            },
          },
          y: { ticks: { callback: (v) => Number(v).toLocaleString() } },
        },
      },
    });
  }
  return chart;
}

function renderChart() {
  const computed = computeSeriesForRender();
  updateSummary(computed);
  renderSeriesManager(computed.managerRows);

  if (state.selectedSheets.size === 0 || state.selectedCountries.size === 0) {
    dom.emptyState.textContent = "Select at least one sheet and one country.";
    dom.emptyState.classList.add("show");
    if (chart) {
      chart.data.datasets = [];
      chart.update();
    }
    scheduleUrlSync();
    return;
  }

  if (!enforceGuardrails(computed.visibleSeriesCount, computed.totalVisiblePoints)) {
    dom.emptyState.textContent = "Selection too large. Narrow filters or date range.";
    dom.emptyState.classList.add("show");
    if (chart) {
      chart.data.datasets = [];
      chart.update();
    }
    return;
  }

  const c = ensureChart();
  if (!c) return;

  dom.emptyState.classList.toggle("show", computed.datasets.length === 0);
  if (computed.datasets.length === 0) {
    dom.emptyState.textContent = "No visible series after filtering.";
  }

  c.data.datasets = computed.datasets;
  c.update();

  if (state.axisMode === "raw") {
    const ratio = getScaleRatio(computed.datasets);
    if (ratio >= 50) {
      setBanner("Large scale differences detected. Consider Indexed or Z-Score mode.", true);
    }
  }

  scheduleUrlSync();
}

function updateRangeAxisButtons() {
  dom.rangeButtons.forEach((btn) => btn.classList.toggle("active", btn.dataset.range === state.activeRange));
  dom.axisButtons.forEach((btn) => btn.classList.toggle("active", btn.dataset.axis === state.axisMode));
}

function syncSelects() {
  const filteredSheets = getFilteredSheets();
  buildSelectOptions(dom.sheetSelect, filteredSheets, state.selectedSheets);

  const filteredCountries = getFilteredCountries();
  buildSelectOptions(dom.countrySelect, filteredCountries, state.selectedCountries);
}

function syncUIAndRender() {
  updateRangeAxisButtons();
  syncSelects();
  renderChart();
}

function onSheetSelectionChange() {
  addUndoSnapshot();
  state.selectedSheets = new Set(Array.from(dom.sheetSelect.selectedOptions).map((o) => o.value));
  cascadeAndPrune(true);
  syncUIAndRender();
}

function onCountrySelectionChange() {
  addUndoSnapshot();
  state.selectedCountries = new Set(Array.from(dom.countrySelect.selectedOptions).map((o) => o.value));
  cascadeAndPrune(false);
  syncUIAndRender();
}

function attachEvents() {
  dom.sheetSelect.addEventListener("change", onSheetSelectionChange);
  dom.countrySelect.addEventListener("change", onCountrySelectionChange);

  dom.sheetFilter.addEventListener("input", () => {
    sheetFilterTerm = normalizeToken(dom.sheetFilter.value);
    syncSelects();
  });

  dom.countryFilter.addEventListener("input", () => {
    countryFilterTerm = normalizeToken(dom.countryFilter.value);
    syncSelects();
  });

  dom.selectFilteredSheetsBtn.addEventListener("click", () => {
    addUndoSnapshot();
    getFilteredSheets().forEach((s) => state.selectedSheets.add(s));
    cascadeAndPrune(false);
    syncUIAndRender();
  });

  dom.selectFilteredCountriesBtn.addEventListener("click", () => {
    addUndoSnapshot();
    getFilteredCountries().forEach((c) => state.selectedCountries.add(c));
    cascadeAndPrune(false);
    syncUIAndRender();
  });

  dom.clearSheetsBtn.addEventListener("click", () => {
    addUndoSnapshot();
    state.selectedSheets.clear();
    cascadeAndPrune(false);
    syncUIAndRender();
  });

  dom.clearCountriesBtn.addEventListener("click", () => {
    addUndoSnapshot();
    state.selectedCountries.clear();
    cascadeAndPrune(false);
    syncUIAndRender();
  });

  dom.clearAllBtn.addEventListener("click", () => {
    addUndoSnapshot();
    state.selectedSheets.clear();
    state.selectedCountries.clear();
    state.hiddenSeries.clear();
    state.hiddenSeriesOrder = [];
    syncUIAndRender();
  });

  dom.undoBtn.addEventListener("click", () => {
    restoreSnapshot(state.undoStack.pop());
  });

  dom.commandApplyBtn.addEventListener("click", () => {
    const parsed = parseCommand(dom.commandInput.value);
    if (parsed.match) {
      dom.suggestions.textContent = "Matched exact/substring command.";
      applySelection(parsed.match.sheet, parsed.match.country);
      return;
    }
    if (parsed.suggestions.length > 0) {
      dom.suggestions.innerHTML = "";
      parsed.suggestions.forEach((item) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "ghost";
        btn.textContent = `${item.country} + ${item.sheet} (${item.score.toFixed(2)})`;
        btn.addEventListener("click", () => applySelection(item.sheet, item.country));
        dom.suggestions.appendChild(btn);
      });
    } else {
      dom.suggestions.textContent = "No match found.";
    }
  });

  dom.commandInput.addEventListener("input", () => {
    clearTimeout(commandDebounce);
    commandDebounce = setTimeout(() => {
      const parsed = parseCommand(dom.commandInput.value);
      if (parsed.match) dom.suggestions.textContent = `Will apply: ${parsed.match.country} + ${parsed.match.sheet}`;
      else if (parsed.suggestions.length) dom.suggestions.textContent = "Fuzzy matches available. Click Apply.";
      else dom.suggestions.textContent = "";
    }, 180);
  });

  dom.commandInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      dom.commandApplyBtn.click();
    }
  });

  dom.rangeButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      state.activeRange = btn.dataset.range;
      syncUIAndRender();
    });
  });

  dom.axisButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      state.axisMode = btn.dataset.axis;
      syncUIAndRender();
    });
  });
}

async function loadWorkbook() {
  setBanner("Loading dataset...");
  const resp = await fetch(DATA_PATH);
  if (!resp.ok) throw new Error(`Failed to load data (${resp.status})`);
  const parsed = await resp.json();

  if (!parsed || typeof parsed.sheets !== "object") {
    throw new Error("Malformed dataset: missing sheets");
  }

  const cleanedSheets = {};
  const skipped = [];
  Object.entries(parsed.sheets).forEach(([name, sheet]) => {
    if (!sheet || !Array.isArray(sheet.countries) || !Array.isArray(sheet.rows)) {
      skipped.push(name);
      return;
    }
    cleanedSheets[name] = sheet;
  });

  workbook = { ...parsed, sheets: cleanedSheets };
  if (skipped.length > 0) {
    setBanner(`Skipped malformed sheets: ${skipped.join(", ")}`, true);
  } else {
    setBanner("");
  }
}

async function init() {
  try {
    await loadWorkbook();
    buildIndexes();

    const fromUrl = hydrateFromUrl();
    const hydratedByUrl = applyHydratedState(fromUrl);
    if (!hydratedByUrl) {
      const fromStorage = hydrateFromStorage();
      const hydratedByStorage = applyHydratedState(fromStorage);
      if (!hydratedByStorage) applyDefaultState();
    }

    cascadeAndPrune(false);
    attachEvents();
    syncUIAndRender();

    if (fromUrl.partial) {
      setBanner("Shared link is partial.", true);
    }
  } catch (err) {
    console.error(err);
    dom.emptyState.classList.add("show");
    dom.emptyState.textContent = `Error loading app: ${err.message}`;
    setBanner("Failed to load dataset.", true);
  }
}

window.addEventListener("beforeunload", () => {
  if (chart) chart.destroy();
});

init();
