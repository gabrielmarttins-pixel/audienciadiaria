const baseFilesInput = document.querySelector("#baseFilesInput");
const baseStatus = document.querySelector("#baseStatus");
const emptyState = document.querySelector("#emptyState");
const statusText = document.querySelector("#status");
const chartCanvas = document.querySelector("#audienceChart");
const metricStrip = document.querySelector("#metricStrip");
const reportDate = document.querySelector("#reportDate");
const reportRange = document.querySelector("#reportRange");
const marketBadge = document.querySelector("#marketBadge");
const intervalList = document.querySelector("#intervalList");
const addIntervalButton = document.querySelector("#addIntervalButton");
const localParticipationList = document.querySelector("#localParticipationList");
const addLocalParticipationButton = document.querySelector("#addLocalParticipationButton");
const downloadImageButton = document.querySelector("#downloadImageButton");
const downloadWarning = document.querySelector("#downloadWarning");
const historicalAudience = document.querySelector("#historicalAudience");
const historicalShare = document.querySelector("#historicalShare");
const productionFallback = document.querySelector("#productionFallback");
const manualProductionTotal = document.querySelector("#manualProductionTotal");
const presenterName = document.querySelector("#presenterName");
const presenterOtherName = document.querySelector("#presenterOtherName");
const presenterName2 = document.querySelector("#presenterName2");
const presenterOtherName2 = document.querySelector("#presenterOtherName2");
const presenterPreview = document.querySelector("#presenterPreview");
const inewsJournalDialog = document.querySelector("#inewsJournalDialog");
const programSelectionDialog = document.querySelector("#programSelectionDialog");
const programSelectionForm = document.querySelector("#programSelectionForm");
const programSelection = document.querySelector("#programSelection");
const programSelectionHint = document.querySelector("#programSelectionHint");
const otherProgramFields = document.querySelector("#otherProgramFields");
const otherProgramName = document.querySelector("#otherProgramName");
const otherProgramStart = document.querySelector("#otherProgramStart");
const otherProgramEnd = document.querySelector("#otherProgramEnd");
const cancelProgramSelection = document.querySelector("#cancelProgramSelection");

let rows = [];
let headers = [];
let metadata = {};
let summary = {};
let inewsItems = [];
let inewsSourceItems = [];
let exhibitionRows = [];
let exhibitionFileLoaded = false;
let inewsFileLoaded = false;
let localParticipationFileLoaded = false;
let fullAudienceData = null;
let selectedProgram = null;

const inewsDotColor = "#6fbfff";

const numberFormatter = new Intl.NumberFormat("pt-BR", {
  maximumFractionDigits: 2,
  minimumFractionDigits: 1,
});

const seriesConfig = [
  {
    key: "GLOBO",
    label: "GLOBO",
    mark: "globo",
    color: "#0b36b8",
    imageKey: "GLOBO",
  },
  {
    key: "CONTEUDO TV",
    label: "CONTEÚDO TV",
    mark: "NIC",
    color: "#8a3d0f",
  },
  {
    key: "RECORD",
    label: "RECORD",
    mark: "rec",
    color: "#e95145",
    imageKey: "RECORD",
  },
  {
    key: "SBT",
    label: "SBT",
    mark: "sbt",
    color: "#57bf3a",
    imageKey: "SBT",
  },
  {
    key: "TV BAND",
    label: "TV BAND",
    mark: "band",
    color: "#d93bd0",
    imageKey: "BAND",
  },
  {
    key: "TV BRASILIA",
    label: "TV BRASILIA",
    mark: "!",
    color: "#f07a2a",
    imageKey: "REDETV",
  },
  {
    key: "TOTAL LIGADOS",
    label: "TOTAL LIGADOS",
    mark: "TLE",
    color: "#a8a8a8",
  },
];

const normalize = (text) =>
  String(text ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();

function canonicalProgramName(value) {
  const name = String(value ?? "").trim();
  return normalize(name) === "gco" ? "GLOBO COMUNIDADE" : name;
}

function isSundayDate(value) {
  const match = String(value ?? "").match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (!match) return false;
  const [, day, month, year] = match.map(Number);
  const date = new Date(year, month - 1, day);
  return date.getFullYear() === year &&
    date.getMonth() === month - 1 &&
    date.getDate() === day &&
    date.getDay() === 0;
}

function detectDelimiter(text) {
  const sample = text.split(/\r?\n/).slice(0, 20).join("\n");
  const options = [",", ";", "\t"];
  return options
    .map((delimiter) => ({
      delimiter,
      count: sample.split(delimiter).length,
    }))
    .sort((a, b) => b.count - a.count)[0].delimiter;
}

function normalizeDelimiter(delimiter) {
  return delimiter === "\\t" ? "\t" : delimiter;
}

function parseCsvLine(line, delimiter) {
  const result = [];
  let value = "";
  let quoted = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const next = line[index + 1];

    if (char === '"' && quoted && next === '"') {
      value += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === delimiter && !quoted) {
      result.push(value.trim());
      value = "";
    } else {
      value += char;
    }
  }

  result.push(value.trim());
  return result;
}

function parseCsv(text, forcedDelimiter) {
  const delimiter =
    forcedDelimiter === "auto" || !forcedDelimiter
      ? detectDelimiter(text)
      : normalizeDelimiter(forcedDelimiter);
  const lines = text
    .replace(/^\uFEFF/, "")
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0);

  if (lines.length < 2) {
    throw new Error("O CSV precisa ter cabeçalho e pelo menos uma linha de dados.");
  }

  const parsedLines = lines.map((line) => parseCsvLine(line, delimiter));
  const headerIndex = findHeaderIndex(parsedLines);

  if (headerIndex < 0) {
    throw new Error("Não encontrei uma linha de cabeçalho válida no CSV.");
  }

  const parsedHeaders = parsedLines[headerIndex];
  const parsedRows = parsedLines.slice(headerIndex + 1).map((values) =>
    parsedHeaders.reduce((row, header, index) => {
      row[header] = values[index] ?? "";
      return row;
    }, {}),
  );

  return {
    parsedHeaders,
    parsedRows,
    parsedMetadata: parseMetadata(parsedLines.slice(0, headerIndex)),
    parsedSummary: parseSummary(parsedLines.slice(0, headerIndex), parsedHeaders),
  };
}

const dailyCsvColumns = [
  { audience: "GLOBO", share: "GLOBO*", target: "GLOBO" },
  { audience: "NIC", share: "NIC*", target: "CONTEUDO TV" },
  { audience: "REC", share: "REC*", target: "RECORD" },
  { audience: "SBT", share: "SBT*", target: "SBT" },
  { audience: "BAND", share: "BAND*", target: "TV BAND" },
  { audience: "RTV", share: "RTV*", target: "TV BRASILIA" },
  { audience: "TLE", share: "TLE*", target: "TOTAL LIGADOS" },
];

function findExactHeader(parsedHeaders, expected) {
  const normalizedExpected = normalize(expected);
  const wantsShare = expected.endsWith("*");
  return parsedHeaders.find((header) => {
    const value = String(header ?? "").trim();
    return value.endsWith("*") === wantsShare && normalize(value.replace(/\*$/, "")) === normalizedExpected.replace(/\*$/, "");
  });
}

function isDailyAudienceCsv(parsedHeaders) {
  return Boolean(
    findExactHeader(parsedHeaders, "Horario") &&
      findExactHeader(parsedHeaders, "GLOBO") &&
      findExactHeader(parsedHeaders, "GLOBO*"),
  );
}

function extractDateFromFileName(fileName) {
  const match = String(fileName ?? "").match(/(?:^|\D)(\d{2})(\d{2})(\d{4})(?:\D|$)/);
  return match ? `${match[1]}-${match[2]}-${match[3]}` : "";
}

function adaptDailyAudienceCsv(parsed, fileName) {
  if (!isDailyAudienceCsv(parsed.parsedHeaders)) return parsed;

  const timeHeader = findExactHeader(parsed.parsedHeaders, "Horario");
  const columnMap = dailyCsvColumns
    .map((column) => ({
      ...column,
      audienceHeader: findExactHeader(parsed.parsedHeaders, column.audience),
      shareHeader: findExactHeader(parsed.parsedHeaders, column.share),
    }))
    .filter((column) => column.audienceHeader);
  const timePattern = /^\d{1,2}:\d{2}(?::\d{2})?$/;
  const minuteRows = parsed.parsedRows.filter((row) =>
    timePattern.test(String(row[timeHeader] ?? "").trim()),
  );
  const providedSummary = parsed.parsedRows.find((row) => {
    const label = String(row[timeHeader] ?? "").trim();
    return label && !timePattern.test(label) && columnMap.some((column) =>
      Number.isFinite(parseNumber(row[column.audienceHeader])),
    );
  });

  const canonicalRows = minuteRows.map((row) => {
    const canonical = { MIN: String(row[timeHeader] ?? "").trim() };
    columnMap.forEach((column) => {
      canonical[column.target] = row[column.audienceHeader] ?? "";
    });
    return canonical;
  });
  const canonicalSummary = {};
  columnMap.forEach((column) => {
    canonicalSummary[column.target] = {
      average: providedSummary ? parseNumber(providedSummary[column.audienceHeader]) : NaN,
      share: providedSummary && column.shareHeader
        ? parsePercent(providedSummary[column.shareHeader])
        : NaN,
    };
  });

  return {
    parsedHeaders: ["MIN", ...columnMap.map((column) => column.target)],
    parsedRows: canonicalRows,
    parsedMetadata: {
      ...parsed.parsedMetadata,
      data: parsed.parsedMetadata.data || extractDateFromFileName(fileName),
      praca: parsed.parsedMetadata.praca || "Distrito Federal",
    },
    parsedSummary: canonicalSummary,
  };
}

function parseMetadata(lines) {
  const meta = {};

  lines.forEach((line) => {
    if (line.length === 2) {
      meta[normalize(line[0])] = line[1];
    } else if (line.length === 1 && /\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}/.test(line[0])) {
      meta.intervalo = line[0];
    }
  });

  return meta;
}

function parseSummary(lines, parsedHeaders) {
  const summaryData = {};
  const ratingLineIndex = lines.findIndex((line) =>
    normalize(line[0]).includes("rat"),
  );

  if (ratingLineIndex < 0) return summaryData;

  const ratingLine = lines[ratingLineIndex];
  const shareLine = lines[ratingLineIndex + 1] ?? [];

  parsedHeaders.forEach((header, index) => {
    summaryData[header] = {
      average: parseNumber(ratingLine[index]),
      share: parsePercent(shareLine[index]),
    };
  });

  return summaryData;
}

function findHeaderIndex(parsedLines) {
  const explicitHeaderIndex = parsedLines.findIndex((line) => {
    const normalized = line.map(normalize);
    return normalized.includes("min") && normalized.length > 2;
  });

  if (explicitHeaderIndex >= 0) return explicitHeaderIndex;

  return parsedLines.findIndex((line, index) => {
    if (index === parsedLines.length - 1 || line.length < 2) return false;

    const nextLine = parsedLines[index + 1] ?? [];
    const numericCells = nextLine.filter((cell) => Number.isFinite(parseNumber(cell)));
    return numericCells.length >= Math.max(1, Math.floor(line.length / 2));
  });
}

function parseNumber(value) {
  const text = String(value ?? "").trim();
  if (!text || text === "-") return NaN;

  const cleaned = text
    .replace(/[^\d,.-]/g, "")
    .replace(/\.(?=\d{3}(?:\D|$))/g, "")
    .replace(",", ".");

  return Number(cleaned);
}

function parseSeriesNumber(value) {
  const text = String(value ?? "").trim();
  if (text === "-") return 0;
  return parseNumber(value);
}

function parsePercent(value) {
  const number = parseNumber(value);
  return Number.isFinite(number) ? number : NaN;
}

function getLabelColumn() {
  return headers.find((header) => normalize(header) === "min") ?? "MIN";
}

function getSeriesHeader(config) {
  const normalizedKey = normalize(config.key);
  return headers.find((header) => normalize(header).includes(normalizedKey));
}

function availableSeries() {
  return seriesConfig
    .map((config) => ({
      ...config,
      header: getSeriesHeader(config),
    }))
    .filter((series) => series.header);
}

function updateChart() {
  const labelKey = getLabelColumn();
  const activeHeader = getActiveHeader();
  const series = buildSeries(labelKey);

  if (!series.length) {
    statusText.textContent = "Não encontrei colunas de audiência no CSV.";
    emptyState.hidden = false;
    clearChart();
    metricStrip.innerHTML = "";
    return;
  }

  emptyState.hidden = true;
  updateHeaderInfo();
  refreshProductionInputState();
  refreshInewsItems();
  updateMetrics(series);
  drawMultiLineChart(series, activeHeader);

  statusText.textContent = "";
}

function buildSeries(labelKey) {
  return availableSeries()
    .map((series) => {
      const points = rows
        .map((row) => ({
          label: row[labelKey],
          value: parseSeriesNumber(row[series.header]),
        }))
        .filter((point) => point.label && Number.isFinite(point.value))
        .filter((point) => !isAggregateTimeLabel(point.label));

      return {
        ...series,
        points: sortTimePoints(points),
      };
    })
    .filter((series) => series.points.length);
}

function updateHeaderInfo() {
  reportDate.textContent = formatDate(metadata.data) ?? "--/--/----";
  reportRange.textContent = formatRange(metadata.intervalo) ?? inferRange();
  if (!marketBadge.dataset.userEdited) {
    marketBadge.value = getMarketCode(metadata.praca);
    resizeMarketBadge();
  }
}

function formatDate(value) {
  if (!value) return null;

  const [day, month, year] = String(value).split("-");
  if (year) return `${day}/${month}/${year}`;

  return value;
}

function inferRange() {
  const labelKey = getLabelColumn();
  const points = rows
    .map((row) => row[labelKey])
    .filter((label) => /^\d{1,2}:\d{2}$/.test(String(label).trim()))
    .map((label) => ({ label }));

  if (!points.length) return "--:-- - --:--";
  const ordered = sortTimePoints(points);
  return `${ordered[0].label} - ${ordered[ordered.length - 1].label}`;
}

function formatRange(value) {
  const match = String(value ?? "").match(/(\d{1,2}:\d{2})(?::\d{2})?\s*-\s*(\d{1,2}:\d{2})(?::\d{2})?/);
  return match ? `${match[1]} - ${match[2]}` : null;
}

function getActiveHeader() {
  return availableSeries().find((series) => series.key === "GLOBO")?.header ?? "";
}

function resizeMarketBadge() {
  const length = Math.max(3, marketBadge.value.length);
  marketBadge.style.width = `${Math.max(11, length + 4)}ch`;
}

function getMarketCode(place) {
  const text = normalize(place);
  if (text.includes("distrito federal")) return "PROGRAMA";
  return "PROGRAMA";
}

function updateMetrics(series) {
  metricStrip.innerHTML = "";

  series.forEach((item) => {
    const calculatedAverage =
      item.points.reduce((sum, point) => sum + point.value, 0) / item.points.length;
    const average = summary[item.header]?.average;
    const share = summary[item.header]?.share;
    const card = document.createElement("article");
    card.className = "metric-card";
    if (item.key === "GLOBO") card.classList.add("globo-metric-card");
    card.style.setProperty("--series-color", item.color);
    const logoSource = getLogoSource(item);
    const logoContent = logoSource
      ? `<img src="${logoSource}" alt="${item.label}" />`
      : item.mark;
    const historicalBadge = item.key === "GLOBO" ? getHistoricalBadgeData() : null;
    const historicalBadgeHtml = historicalBadge
      ? `<span class="historical-audience-badge ${historicalBadge.tone}" title="Média histórica e variação">
          <span>Média histórica</span>
          <b>${escapeHtml(historicalBadge.audience)} pts | ${escapeHtml(historicalBadge.variation)}</b>
        </span>`
      : "";
    card.innerHTML = `
      <span class="logo-dot" style="background:${item.color}">${logoContent}</span>
      <span class="metric-copy">
        <span class="metric-audience-row">
          <strong>${formatMetricValue(average, calculatedAverage)}</strong>
          ${historicalBadgeHtml}
        </span>
        <span>Audiência</span>
        <b>${formatShareValue(share)}</b>
        <span>Share</span>
      </span>
    `;
    metricStrip.appendChild(card);
  });
}

function getHistoricalBadgeData() {
  const audienceNumber = parseNumber(historicalAudience.value);
  const variationNumber = parseNumber(historicalShare.value);

  if (!Number.isFinite(audienceNumber) && !Number.isFinite(variationNumber)) return null;

  return {
    audience: Number.isFinite(audienceNumber)
      ? formatSignedMetricValue(historicalAudience.value)
      : "-",
    variation: Number.isFinite(variationNumber)
      ? formatVariationInput(historicalShare.value)
      : "-",
    tone: Number.isFinite(variationNumber) && variationNumber < 0 ? "negative" : "positive",
  };
}

function formatMetricValue(value, fallback) {
  const number = Number.isFinite(value) ? value : fallback;
  return Number.isFinite(number) ? numberFormatter.format(number) : "-";
}

function formatShareValue(value) {
  return Number.isFinite(value) ? `${numberFormatter.format(value)}%` : "-";
}

function isAggregateTimeLabel(label) {
  return /^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}$/.test(String(label).trim());
}

function sortTimePoints(points) {
  const everyPointHasTime = points.every((point) =>
    /^\d{1,2}:\d{2}(?::\d{2})?$/.test(String(point.label).trim()),
  );

  if (!everyPointHasTime) return points;

  const start = timeToSeconds(points[0]?.label);
  return [...points].sort(
    (a, b) => timeToTimelineSeconds(a.label, start) - timeToTimelineSeconds(b.label, start),
  );
}

function timeToMinutes(label) {
  const [hours, minutes] = String(label).split(":").map(Number);
  return hours * 60 + minutes;
}

function clearChart() {
  const context = chartCanvas.getContext("2d");
  context.clearRect(0, 0, chartCanvas.width, chartCanvas.height);
}

function resizeCanvas() {
  const ratio = window.devicePixelRatio || 1;
  const rect = chartCanvas.getBoundingClientRect();
  chartCanvas.width = Math.round(rect.width * ratio);
  chartCanvas.height = Math.round(rect.height * ratio);
  const context = chartCanvas.getContext("2d");
  context.setTransform(ratio, 0, 0, ratio, 0, 0);
  return { width: rect.width, height: rect.height, context };
}

function drawMultiLineChart(series, activeHeader) {
  const { width, height, context } = resizeCanvas();
  const values = series.flatMap((item) => item.points.map((point) => point.value));
  const maxValue = Math.max(...values);
  const minValue = Math.min(0, ...values);
  const topValue = Math.ceil(maxValue + Math.max(1, maxValue * 0.08));
  const range = topValue - minValue || 1;
  const padding = { top: 36, right: 16, bottom: 48, left: 22 };
  const plotWidth = width - padding.left - padding.right;
  const plotHeight = height - padding.top - padding.bottom;
  const longest = series.reduce((winner, item) =>
    item.points.length > winner.points.length ? item : winner,
  );

  context.clearRect(0, 0, width, height);
  drawVerticalBand(context, longest.points, padding, plotWidth, plotHeight);
  drawAxisLabels(context, longest.points, padding, plotWidth, height);

  series
    .slice()
    .sort((a, b) => (a.header === activeHeader ? 1 : 0) - (b.header === activeHeader ? 1 : 0))
    .forEach((item) => {
      drawSeriesLine({
        context,
        item,
        active: item.header === activeHeader,
        padding,
        plotWidth,
        plotHeight,
        minValue,
        range,
      });
    });
  drawInewsMarkers(context, series, longest.points, padding, plotWidth, plotHeight, minValue, range);
  drawInewsLegend(context, padding);
}

function drawMultiLineChartOnContext(context, series, activeHeader, bounds, options = {}) {
  const values = series.flatMap((item) => item.points.map((point) => point.value));
  const maxValue = Math.max(...values);
  const minValue = Math.min(0, ...values);
  const topValue = Math.ceil(maxValue + Math.max(1, maxValue * 0.08));
  const range = topValue - minValue || 1;
  const padding = { top: 36, right: 16, bottom: 48, left: 22 };
  const plotWidth = bounds.width - padding.left - padding.right;
  const plotHeight = bounds.height - padding.top - padding.bottom;
  const longest = series.reduce((winner, item) =>
    item.points.length > winner.points.length ? item : winner,
  );

  context.save();
  context.translate(bounds.x, bounds.y);
  context.beginPath();
  context.rect(0, 0, bounds.width, bounds.height);
  context.clip();
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, bounds.width, bounds.height);

  drawVerticalBand(context, longest.points, padding, plotWidth, plotHeight);
  drawAxisLabels(context, longest.points, padding, plotWidth, bounds.height);

  series
    .slice()
    .sort((a, b) => (a.header === activeHeader ? 1 : 0) - (b.header === activeHeader ? 1 : 0))
    .forEach((item) => {
      drawSeriesLine({
        context,
        item,
        active: item.header === activeHeader,
        padding,
        plotWidth,
        plotHeight,
        minValue,
        range,
        compactHighlights: options.compactHighlights === true,
        highlightCollector: options.highlightCollector,
        chartOffset: bounds,
      });
    });
  if (options.includeInews !== false) {
    drawInewsMarkers(context, series, longest.points, padding, plotWidth, plotHeight, minValue, range);
  }

  context.restore();
}

function timeToSeconds(value) {
  const parts = String(value ?? "").trim().split(":").map(Number);
  if (parts.length < 2 || parts.some((part) => !Number.isFinite(part))) return NaN;
  const [hours, minutes, seconds = 0] = parts;
  return hours * 3600 + minutes * 60 + seconds;
}

function timeToTimelineSeconds(value, start) {
  const seconds = timeToSeconds(value);
  if (!Number.isFinite(seconds) || !Number.isFinite(start)) return NaN;
  return start >= 12 * 3600 && seconds < 12 * 3600 ? seconds + 86400 : seconds;
}

function xForTime(time, points, left, plotWidth) {
  const first = timeToSeconds(points[0]?.label);
  const last = timeToTimelineSeconds(points[points.length - 1]?.label, first);
  const target = timeToTimelineSeconds(time, first);
  if (![first, last, target].every(Number.isFinite) || last === first) return NaN;
  return left + ((target - first) / (last - first)) * plotWidth;
}

function durationToSeconds(value) {
  const parts = String(value ?? "").trim().split(":").map(Number);
  if (parts.length === 2 && parts.every(Number.isFinite)) return parts[0] * 60 + parts[1];
  if (parts.length === 3 && parts.every(Number.isFinite)) {
    return parts[0] * 3600 + parts[1] * 60 + parts[2];
  }
  return 0;
}

function formatProductionDuration(totalSeconds) {
  const seconds = Math.max(0, Math.round(totalSeconds));
  const minutes = Math.floor(seconds / 60);
  const minuteSeconds = `${String(minutes % 60).padStart(2, "0")}:${String(seconds % 60).padStart(2, "0")}`;
  return minutes < 60
    ? `${String(minutes).padStart(2, "0")}:${String(seconds % 60).padStart(2, "0")}`
    : `${String(Math.floor(minutes / 60)).padStart(2, "0")}:${minuteSeconds}`;
}

function getAutomaticProductionTotalSeconds() {
  const program = getExhibitionProgramName(marketBadge.value);
  let total = 0;
  let hasProduction = false;
  exhibitionRows.forEach((row) => {
    if (getExhibitionProgramName(getExhibitionField(row, "Programa")) !== program) return;
    const segment = String(getExhibitionField(row, "Segmento")).trim().toUpperCase();
    if (!/^PD\d+$/.test(segment)) return;
    const duration = durationToSeconds(getExhibitionField(row, "Duracao"));
    if (duration <= 0) return;
    total += duration;
    hasProduction = true;
  });
  return hasProduction ? total : null;
}

function parseManualProductionDuration(value) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  if (!/^\d{1,3}:[0-5]\d(?::[0-5]\d)?$/.test(text)) return NaN;
  return durationToSeconds(text);
}

function getProductionTotalSeconds() {
  const automatic = getAutomaticProductionTotalSeconds();
  return automatic !== null ? automatic : parseManualProductionDuration(manualProductionTotal.value);
}

function refreshProductionInputState() {
  productionFallback.hidden = !rows.length || getAutomaticProductionTotalSeconds() !== null;
}

function secondsToClock(totalSeconds) {
  const normalizedSeconds = ((Math.round(totalSeconds) % 86400) + 86400) % 86400;
  const hours = Math.floor(normalizedSeconds / 3600);
  const minutes = Math.floor((normalizedSeconds % 3600) / 60);
  const seconds = normalizedSeconds % 60;
  return [hours, minutes, seconds].map((part) => String(part).padStart(2, "0")).join(":");
}

function parseExhibitionCsv(text) {
  const delimiter = detectDelimiter(text);
  const lines = String(text ?? "")
    .replace(/^\uFEFF/, "")
    .split(/\r?\n/)
    .filter((line) => line.trim());
  if (lines.length < 2) throw new Error("O arquivo do Maestro está vazio.");

  const parsedLines = lines.map((line) => parseCsvLine(line, delimiter));
  const headerIndex = parsedLines.findIndex((line) => {
    const values = line.map(normalize);
    return values.includes("programa") && values.includes("segmento") && values.includes("horainicio");
  });
  if (headerIndex < 0) {
    throw new Error("Não encontrei as colunas Programa, Segmento e HoraInicio no arquivo do Maestro.");
  }

  const sourceHeaders = parsedLines[headerIndex];
  return parsedLines.slice(headerIndex + 1).map((values) =>
    sourceHeaders.reduce((row, header, index) => {
      row[header] = values[index] ?? "";
      return row;
    }, {}),
  );
}

function getExhibitionField(row, fieldName) {
  const key = Object.keys(row).find((header) => normalize(header) === normalize(fieldName));
  return key ? row[key] : "";
}

function getExhibitionProgramName(programName) {
  const value = normalize(canonicalProgramName(programName)).replace(/\s+/g, " ");
  const aliases = {
    "bom dia df": "bom dia praca",
    bddf: "bom dia praca",
    df1: "praca tv i",
    df2: "praca tv ii",
    "globo esporte": "globo esporte",
    "globo esporte df": "globo esporte",
    "bom dia brasil": "bom dia brasil",
    "boletim df2": "boletim df2",
  };
  return aliases[value] || value;
}

function identifyProgramIntervals(sourceRows, programName) {
  const exhibitionProgram = getExhibitionProgramName(programName);
  const realtimeBounds = getRealtimeBounds();
  const matchingRows = sourceRows.filter(
    (row) => getExhibitionProgramName(getExhibitionField(row, "Programa")) === exhibitionProgram,
  );
  const groups = new Map();

  matchingRows.forEach((row) => {
    const segment = String(getExhibitionField(row, "Segmento")).trim().toUpperCase();
    if (!/^PT\d+$/.test(segment)) return;
    const start = timeToTimelineSeconds(
      getExhibitionField(row, "HoraInicio"),
      realtimeBounds?.start ?? 0,
    );
    const duration = durationToSeconds(getExhibitionField(row, "Duracao"));
    if (!Number.isFinite(start) || duration <= 0) return;
    const current = groups.get(segment) || { start, end: start + duration };
    current.start = Math.min(current.start, start);
    current.end = Math.max(current.end, start + duration);
    groups.set(segment, current);
  });

  return [...groups.values()]
    .filter(({ start, end }) => !realtimeBounds || (end >= realtimeBounds.start && start <= realtimeBounds.end))
    .sort((a, b) => a.start - b.start)
    .map(({ start, end }) => ({ start: secondsToClock(start), end: secondsToClock(end) }));
}

function getMaestroProductionBlocks() {
  const program = getExhibitionProgramName(marketBadge.value);
  const realtimeStart = getRealtimeBounds()?.start ?? 0;
  const groups = new Map();
  exhibitionRows.forEach((row) => {
    if (getExhibitionProgramName(getExhibitionField(row, "Programa")) !== program) return;
    const segment = String(getExhibitionField(row, "Segmento")).trim().toUpperCase();
    if (!/^PD\d+$/.test(segment)) return;
    const start = timeToTimelineSeconds(getExhibitionField(row, "HoraInicio"), realtimeStart);
    const duration = durationToSeconds(getExhibitionField(row, "Duracao"));
    if (!Number.isFinite(start) || duration <= 0) return;
    const block = groups.get(segment) || { start, end: start + duration };
    block.start = Math.min(block.start, start);
    block.end = Math.max(block.end, start + duration);
    groups.set(segment, block);
  });
  return [...groups.values()].sort((a, b) => a.start - b.start);
}

function applyExhibitionIntervals() {
  const intervals = identifyProgramIntervals(exhibitionRows, marketBadge.value);
  intervalList.innerHTML = "";
  intervals.forEach(({ start, end }) => addIntervalRow(start, end));
  if (!intervals.length) addIntervalRow();
  return intervals;
}

function refreshExhibitionIntervals() {
  const intervals = applyExhibitionIntervals();
  return intervals;
}

function applyEveningProgramNameFromMaestro() {
  if (marketBadge.dataset.userEdited === "true" || !getRealtimeBounds()) return;
  if (identifyProgramIntervals(exhibitionRows, marketBadge.value).length) return;

  const bounds = getRealtimeBounds();
  const realtimeEnd = bounds.end + 60;
  const realtimeDuration = realtimeEnd - bounds.start;
  if (realtimeDuration <= 0) return;

  const candidates = [
    { name: "DF2", sourceName: "praca tv ii" },
    { name: "BOLETIM DF2", sourceName: "boletim df2" },
  ].map((candidate) => {
    const production = exhibitionRows
      .filter((row) =>
        normalize(getExhibitionField(row, "Programa")) === candidate.sourceName &&
        /^PD\d+$/.test(String(getExhibitionField(row, "Segmento")).trim().toUpperCase()),
      )
      .map((row) => ({
        start: timeToSeconds(getExhibitionField(row, "HoraInicio")),
        duration: durationToSeconds(getExhibitionField(row, "Duracao")),
      }))
      .filter(({ start, duration }) => Number.isFinite(start) && duration > 0);
    if (!production.length) return { ...candidate, overlap: 0 };
    const start = Math.min(...production.map((item) => item.start));
    const end = Math.max(...production.map((item) => item.start + item.duration));
    return {
      ...candidate,
      overlap: Math.max(0, Math.min(realtimeEnd, end) - Math.max(bounds.start, start)),
    };
  }).sort((a, b) => b.overlap - a.overlap);

  if (candidates[0].overlap < realtimeDuration * 0.5) return;
  marketBadge.value = candidates[0].name;
  marketBadge.dataset.userEdited = "auto";
  resizeMarketBadge();
}

function updateImportWorkflow() {
  baseStatus.textContent = rows.length
    ? `${selectedProgram?.name ?? "Programa"}: ${rows.length} minutos${selectedProgram?.source === "approximate" ? " (horários aproximados)" : ""} · Maestro: ${exhibitionFileLoaded ? "conectado" : "não enviado"} · iNews: ${inewsFileLoaded ? "conectado" : "não enviado"}${isBomDiaBrasil() ? ` · Participação local: ${localParticipationFileLoaded ? "conectada" : "manual/opcional"}` : ""}`
    : "Realtime obrigatório · Maestro, iNews e participação local (BDBR) opcionais";
}

function drawInewsMarkers(
  context,
  series,
  points,
  padding,
  plotWidth,
  plotHeight,
  minValue,
  range,
) {
  const timelineY = getInewsTimelineY(series, padding, plotHeight, minValue, range);
  const dots = buildInewsDotData(
    series,
    points,
    padding,
    plotWidth,
    plotHeight,
    minValue,
    range,
  );
  if (!dots.length || !Number.isFinite(timelineY)) return;
  context.save();
  context.strokeStyle = "#b9ddf8";
  context.lineWidth = 2;
  context.beginPath();
  context.moveTo(padding.left, timelineY);
  context.lineTo(padding.left + plotWidth, timelineY);
  context.stroke();
  dots.forEach((dot) => {
    context.fillStyle = dot.color;
    context.beginPath();
    context.arc(dot.x, timelineY, dot.radius, 0, Math.PI * 2);
    context.fill();
    context.strokeStyle = "#ffffff";
    context.lineWidth = 1.5;
    context.stroke();
  });
  context.restore();
}

function buildInewsDotData(series, points, padding, plotWidth, plotHeight, minValue, range) {
  const globo = series.find((item) => item.key === "GLOBO");
  if (!globo?.points.length) return [];
  const timelineStart = timeToSeconds(points[0]?.label);
  const intervals = getIntervals();
  const localParticipations = getLocalParticipations();
  const localOnly = isBomDiaBrasil();
  const intervalStartXs = intervals
    .map(({ start }) => xForTime(start, points, padding.left, plotWidth))
    .filter(Number.isFinite)
    .sort((a, b) => a - b);
  const rawDots = inewsItems
    .map((item) => {
      const itemSeconds = timeToTimelineSeconds(item.time, timelineStart);
      if (intervals.some(({ start, end }) =>
        itemSeconds >= timeToTimelineSeconds(start, timelineStart) &&
        itemSeconds < timeToTimelineSeconds(end, timelineStart),
      )) return null;
      const localWindow = localOnly
        ? localParticipations.find(({ start, end }) =>
            itemSeconds >= timeToTimelineSeconds(start, timelineStart) &&
            itemSeconds < timeToTimelineSeconds(end, timelineStart),
          )
        : null;
      if (localOnly && !localWindow) return null;
      const x = xForTime(item.time, points, padding.left, plotWidth);
      const endTime = secondsToClock(itemSeconds + durationToSeconds(item.duration));
      const endX = xForTime(endTime, points, padding.left, plotWidth);
      const globoValue = valueAtTime(globo.points, item.time);
      if (
        !Number.isFinite(x) ||
        x < padding.left ||
        x > padding.left + plotWidth ||
        !Number.isFinite(globoValue)
      ) {
        return null;
      }
      const boundedEndX = Number.isFinite(endX)
        ? Math.min(padding.left + plotWidth, Math.max(x, endX))
        : x;
      const localEndX = localWindow
        ? xForTime(localWindow.end, points, padding.left, plotWidth)
        : NaN;
      const nextIntervalX = intervalStartXs.find(
        (intervalX) => intervalX > x && intervalX < boundedEndX,
      );
      return {
        ...item,
        x,
        endX: Math.min(
          boundedEndX,
          Number.isFinite(nextIntervalX) ? nextIntervalX : boundedEndX,
          Number.isFinite(localEndX) ? localEndX : boundedEndX,
        ),
        globoY: yFor(globoValue, padding.top, plotHeight, minValue, range),
        color: inewsDotColor,
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.x - b.x);

  const clusters = [];
  rawDots.forEach((dot) => {
    const current = clusters[clusters.length - 1];
    const lastItemX = current?.items[current.items.length - 1]?.x;
    const intervalBetweenDots = intervalStartXs.some(
      (intervalX) => intervalX > lastItemX && intervalX <= dot.x,
    );
    if (current && dot.x - current.x <= 15 && !intervalBetweenDots) {
      current.items.push(dot);
      current.x = current.items.reduce((sum, item) => sum + item.x, 0) / current.items.length;
    } else {
      clusters.push({ items: [dot], x: dot.x });
    }
  });

  return clusters.map((cluster) => ({
    x: cluster.x,
    color: inewsDotColor,
    radius: Math.min(11, 5 + (cluster.items.length - 1) * 1.5),
    count: cluster.items.length,
    globoY:
      cluster.items.reduce((sum, item) => sum + item.globoY, 0) / cluster.items.length,
    bandX: Math.min(...cluster.items.map((item) => item.x)),
    bandEndX: Math.max(...cluster.items.map((item) => Math.max(item.x, item.endX))),
    title: cluster.items.map((item) => item.title).join(" / "),
    time: cluster.items.map((item) => item.time).join(" / "),
    duration: cluster.items.map((item) => item.duration).filter(Boolean).join(" / "),
  }));
}

function getInewsTimelineY(series, padding, plotHeight, minValue, range) {
  const tle = series.find((item) => item.key === "TOTAL LIGADOS");
  if (!tle?.points.length) return NaN;
  const highestTleValue = Math.max(...tle.points.map((point) => point.value));
  const highestTleY = yFor(highestTleValue, padding.top, plotHeight, minValue, range);
  return Math.max(padding.top + 18, highestTleY - 28);
}

function drawInewsLegend(context, padding) {
  if (!inewsItems.length) return;
  const x = padding.left + (getLocalParticipations().length ? 292 : 126);
  const legendY = 0;
  context.save();
  context.fillStyle = "#eef7ff";
  context.fillRect(x, legendY, 118, 28);
  context.fillStyle = inewsDotColor;
  context.beginPath();
  context.arc(x + 17, legendY + 14, 5, 0, Math.PI * 2);
  context.fill();
  context.fillStyle = "#27648f";
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.font = "11px Globotipo, Arial, sans-serif";
  context.fillText("Retrancas", x + 70, legendY + 14);
  context.restore();
}

function getInewsCorridor(time, series, padding, plotHeight, minValue, range) {
  const tle = series.find((item) => item.key === "TOTAL LIGADOS");
  const lowerSeries = series.filter((item) => item.key === "GLOBO" || item.key === "CONTEUDO TV");
  if (!tle || !lowerSeries.length) return null;

  const tleValue = valueAtTime(tle.points, time);
  const lowerValues = lowerSeries
    .map((item) => valueAtTime(item.points, time))
    .filter(Number.isFinite);
  if (!Number.isFinite(tleValue) || !lowerValues.length) return null;

  const tleY = yFor(tleValue, padding.top, plotHeight, minValue, range);
  const closestLowerY = Math.min(
    ...lowerValues.map((value) => yFor(value, padding.top, plotHeight, minValue, range)),
  );
  const top = tleY + 10;
  const bottom = closestLowerY - 10;
  return bottom - top >= 24 ? { top, bottom } : null;
}

function valueAtTime(points, time) {
  if (!points.length) return NaN;
  const start = timeToSeconds(points[0].label);
  const target = timeToTimelineSeconds(time, start);
  let closest = points[0];
  let distance = Math.abs(timeToTimelineSeconds(closest.label, start) - target);
  points.forEach((point) => {
    const nextDistance = Math.abs(timeToTimelineSeconds(point.label, start) - target);
    if (nextDistance < distance) {
      closest = point;
      distance = nextDistance;
    }
  });
  return closest.value;
}

function drawVerticalBand(context, points, padding, plotWidth, plotHeight) {
  drawTimeBands(
    context,
    points,
    padding,
    plotWidth,
    plotHeight,
    getIntervals(),
    "rgba(0, 0, 0, 0.035)",
  );
  drawTimeBands(
    context,
    points,
    padding,
    plotWidth,
    plotHeight,
    getLocalParticipations(),
    "rgba(246, 172, 142, 0.34)",
  );
}

function drawTimeBands(context, points, padding, plotWidth, plotHeight, ranges, color) {
  ranges.forEach(({ start, end }) => {
    const startX = xForTime(start, points, padding.left, plotWidth);
    const endX = xForTime(end, points, padding.left, plotWidth);
    if (!Number.isFinite(startX) || !Number.isFinite(endX)) return;

    const bandX = Math.min(startX, endX);
    const bandEndX = Math.max(startX, endX);
    const bandWidth = Math.max(2, bandEndX - bandX);

    context.fillStyle = color;
    context.fillRect(bandX, padding.top, bandWidth, plotHeight);
  });
}

function getIntervals() {
  return [...intervalList.querySelectorAll(".interval-row")]
    .map((row) => ({
      start: row.querySelector(".interval-start").value,
      end: row.querySelector(".interval-end").value,
    }))
    .filter((interval) => interval.start && interval.end);
}

function getLocalParticipations() {
  return [...localParticipationList.querySelectorAll(".local-participation-row")]
    .map((row) => ({
      start: row.querySelector(".local-participation-start").value,
      end: row.querySelector(".local-participation-end").value,
    }))
    .filter((participation) => participation.start && participation.end);
}

function isBomDiaBrasil() {
  const name = normalize(marketBadge.value).replace(/\s+/g, "");
  return name === "bomdiabrasil" || name === "bdbr";
}

function addIntervalRow(start = "", end = "") {
  const row = document.createElement("div");
  row.className = "interval-row";
  row.innerHTML = `
    <label>
      Início
      <input class="interval-start" type="time" step="1" value="${start}" />
    </label>
    <label>
      Fim
      <input class="interval-end" type="time" step="1" value="${end}" />
    </label>
    <button class="remove-interval-button" type="button" aria-label="Remover intervalo">×</button>
  `;

  row.querySelectorAll("input").forEach((input) => {
    input.addEventListener("change", updateChart);
  });

  row.querySelector(".remove-interval-button").addEventListener("click", () => {
    row.remove();
    if (!intervalList.children.length) addIntervalRow();
    updateChart();
  });

  intervalList.appendChild(row);
}

function addLocalParticipationRow(start = "", end = "") {
  const row = document.createElement("div");
  row.className = "interval-row local-participation-row";
  row.innerHTML = `
    <label>
      Inicio
      <input class="local-participation-start" type="time" step="1" value="${start}" />
    </label>
    <label>
      Fim
      <input class="local-participation-end" type="time" step="1" value="${end}" />
    </label>
    <button class="remove-interval-button" type="button" aria-label="Remover participacao local">x</button>
  `;

  row.querySelectorAll("input").forEach((input) => {
    input.addEventListener("change", updateChart);
  });

  row.querySelector(".remove-interval-button").addEventListener("click", () => {
    row.remove();
    if (!localParticipationList.children.length) addLocalParticipationRow();
    updateChart();
  });

  localParticipationList.appendChild(row);
}

function findClosestTimeIndex(points, time) {
  const start = timeToSeconds(points[0]?.label);
  const target = timeToTimelineSeconds(time, start);
  if (!Number.isFinite(target)) return -1;

  let closestIndex = -1;
  let closestDistance = Infinity;
  points.forEach((point, index) => {
    const distance = Math.abs(timeToTimelineSeconds(point.label, start) - target);
    if (distance < closestDistance) {
      closestDistance = distance;
      closestIndex = index;
    }
  });

  return closestIndex;
}

function drawAxisLabels(context, points, padding, plotWidth, height) {
  const indexes = pickTimeIndexes(points);
  context.font = "12px Globotipo, Arial, sans-serif";
  context.textAlign = "center";
  context.fillStyle = "#9a9a9a";

  indexes.forEach((index) => {
    const x = xFor(index, points.length, padding.left, plotWidth);
    context.fillText(points[index].label, x, height - 20);
  });

  const legendY = 0;
  context.fillStyle = "#f1f1f1";
  context.fillRect(padding.left, legendY, 118, 28);
  context.fillStyle = "#8f8f8f";
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.font = "11px Globotipo, Arial, sans-serif";
  context.fillText("Intervalo", padding.left + 59, legendY + 14);

  if (getLocalParticipations().length) {
    context.fillStyle = "#f6ac8e";
    context.fillRect(padding.left + 126, legendY, 158, 28);
    context.fillStyle = "#7a321c";
    context.fillText("Participação local", padding.left + 205, legendY + 14);
  }
  const productionSeconds = getProductionTotalSeconds();
  const productionText = `Produção total: ${Number.isFinite(productionSeconds) ? formatProductionDuration(productionSeconds) : "--:--"}`;
  context.font = "700 11px Globotipo, Arial, sans-serif";
  const productionWidth = Math.ceil(context.measureText(productionText).width + 28);
  const productionX = padding.left + plotWidth + padding.right - 4 - productionWidth;
  context.fillStyle = "#eaf3ff";
  roundRect(context, productionX, legendY, productionWidth, 28, 8);
  context.fill();
  context.fillStyle = "#0b36a8";
  context.fillText(productionText, productionX + productionWidth / 2, legendY + 14);
  context.textBaseline = "alphabetic";
}

function drawSeriesLine({
  context,
  item,
  active,
  padding,
  plotWidth,
  plotHeight,
  minValue,
  range,
  compactHighlights = false,
  highlightCollector,
  chartOffset,
}) {
  const line = new Path2D();

  item.points.forEach((point, index) => {
    const x = xFor(index, item.points.length, padding.left, plotWidth);
    const y = yFor(point.value, padding.top, plotHeight, minValue, range);

    if (index === 0) line.moveTo(x, y);
    else line.lineTo(x, y);
  });

  context.save();
  context.globalAlpha = active ? 1 : 0.92;
  context.strokeStyle = item.color;
  context.lineWidth = active ? 4 : item.key === "TOTAL LIGADOS" ? 3.5 : 2.7;
  context.lineJoin = "round";
  context.lineCap = "round";
  context.stroke(line);
  context.restore();

  if (item.key === "GLOBO") {
    const highlightBoxes = drawGloboHighlights(
      context,
      item,
      padding,
      plotWidth,
      plotHeight,
      minValue,
      range,
      compactHighlights,
    );
    if (highlightCollector) {
      highlightBoxes.forEach((box) => highlightCollector.push({
        ...box,
        x: box.x + chartOffset.x,
        y: box.y + chartOffset.y,
      }));
    }
    drawImportantMovements(
      context,
      item,
      padding,
      plotWidth,
      plotHeight,
      minValue,
      range,
      highlightBoxes,
    );
  } else {
    drawPointLabel(context, item, item.points.length - 1, padding, plotWidth, plotHeight, minValue, range);
  }

  if (item.key === "TOTAL LIGADOS") {
    drawPointLabel(context, item, Math.floor(item.points.length * 0.8), padding, plotWidth, plotHeight, minValue, range);
  } else if (item.key === "RECORD" || item.key === "SBT") {
    drawPointLabel(context, item, Math.floor(item.points.length * 0.5), padding, plotWidth, plotHeight, minValue, range);
  }
}

function drawImportantMovements(
  context,
  item,
  padding,
  plotWidth,
  plotHeight,
  minValue,
  range,
  occupiedBoxes = [],
) {
  if (item.points.length < 2) return;

  const movements = buildAudienceTrendMovements(item.points);
  const rises = pickImportantMovements(movements.filter((movement) => movement.delta > 0));
  const falls = pickImportantMovements(movements.filter((movement) => movement.delta < 0));

  rises.forEach((movement) =>
    drawMovementLabel(
      context,
      item,
      movement,
      "Subida importante",
      "▲",
      "#2497df",
      padding,
      plotWidth,
      plotHeight,
      minValue,
      range,
      occupiedBoxes,
    ),
  );
  falls.forEach((movement) =>
    drawMovementLabel(
      context,
      item,
      movement,
      "Queda importante",
      "▼",
      "#e95145",
      padding,
      plotWidth,
      plotHeight,
      minValue,
      range,
      occupiedBoxes,
    ),
  );
}

function buildAudienceTrendMovements(points) {
  const movements = [];
  const maxWindow = 6;

  for (let endIndex = 1; endIndex < points.length; endIndex += 1) {
    let strongest = null;
    const firstStartIndex = Math.max(0, endIndex - maxWindow);

    for (let startIndex = firstStartIndex; startIndex < endIndex; startIndex += 1) {
      const delta = points[endIndex].value - points[startIndex].value;
      if (Math.abs(delta) < 0.5) continue;

      const segment = points.slice(startIndex, endIndex + 1).map((point) => point.value);
      const endsAtExtreme =
        delta > 0
          ? points[endIndex].value === Math.max(...segment)
          : points[endIndex].value === Math.min(...segment);
      if (!endsAtExtreme) continue;

      const confirmation = points
        .slice(endIndex + 1, Math.min(points.length, endIndex + 3))
        .map((point) => point.value);
      const retainedLevel = points[startIndex].value + delta * 0.55;
      const isPersistent = confirmation.every((value) =>
        delta > 0 ? value >= retainedLevel : value <= retainedLevel,
      );
      if (!isPersistent) continue;

      if (!strongest || Math.abs(delta) > Math.abs(strongest.delta)) {
        strongest = { startIndex, index: endIndex, delta };
      }
    }

    if (strongest) movements.push(strongest);
  }

  return movements;
}

function pickImportantMovements(movements) {
  if (!movements.length) return [];
  const sorted = [...movements].sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
  const threshold = Math.max(0.5, Math.abs(sorted[0].delta) * 0.55);
  const selected = [];
  sorted.forEach((movement) => {
    if (selected.length >= 2 || Math.abs(movement.delta) < threshold) return;
    const overlapsSelectedTrend = selected.some((candidate) => {
      const candidateStart = candidate.startIndex ?? candidate.index - 1;
      const movementStart = movement.startIndex ?? movement.index - 1;
      return (
        movementStart <= candidate.index + 2 &&
        movement.index >= candidateStart - 2
      );
    });
    if (overlapsSelectedTrend) return;
    selected.push(movement);
  });
  return selected;
}

function drawMovementLabel(
  context,
  item,
  movement,
  caption,
  arrow,
  color,
  padding,
  plotWidth,
  plotHeight,
  minValue,
  range,
  occupiedBoxes,
) {
  const previous = item.points[movement.startIndex ?? movement.index - 1];
  const current = item.points[movement.index];
  if (!previous || !current) return;

  const currentX = xFor(movement.index, item.points.length, padding.left, plotWidth);
  const currentY = yFor(current.value, padding.top, plotHeight, minValue, range);
  const boxWidth = 94;
  const boxHeight = 26;
  const initialBoxX = Math.min(
    Math.max(4, currentX - boxWidth / 2),
    padding.left + plotWidth - boxWidth,
  );
  const preferredY = movement.delta > 0 ? currentY - 36 : currentY + 10;
  const alternateY = movement.delta > 0 ? currentY + 10 : currentY - 36;
  const candidateOffsets = [
    { x: 0, y: preferredY },
    { x: 0, y: alternateY },
    { x: -36, y: preferredY },
    { x: 36, y: preferredY },
    { x: 0, y: preferredY - 32 },
    { x: 0, y: alternateY + 32 },
  ];
  const candidates = candidateOffsets.map((candidate) => ({
    x: Math.min(
      Math.max(4, initialBoxX + candidate.x),
      padding.left + plotWidth - boxWidth,
    ),
    y: Math.min(
      Math.max(padding.top + 4, candidate.y),
      padding.top + plotHeight - boxHeight - 4,
    ),
    width: boxWidth,
    height: boxHeight,
  }));
  const selectedBox = candidates.find(
    (candidate) => !occupiedBoxes.some((occupied) => rectanglesOverlap(candidate, occupied, 6)),
  );
  if (!selectedBox) return;
  const boxX = selectedBox.x;
  const boxY = selectedBox.y;
  occupiedBoxes.push(selectedBox);
  const formattedDelta = `${movement.delta > 0 ? "+" : "-"}${numberFormatter.format(
    Math.abs(movement.delta),
  )}`;

  context.save();
  context.fillStyle = "rgba(255, 255, 255, 0.96)";
  context.strokeStyle = color;
  context.lineWidth = 1.5;
  roundRect(context, boxX, boxY, boxWidth, boxHeight, 7);
  context.fill();
  context.stroke();
  context.fillStyle = color;
  context.textAlign = "center";
  context.font = "900 13px Globotipo, Arial, sans-serif";
  context.fillText(`${arrow} ${formattedDelta}`, boxX + boxWidth / 2, boxY + 18);
  context.beginPath();
  context.arc(currentX, currentY, 4.5, 0, Math.PI * 2);
  context.fill();
  context.strokeStyle = "#ffffff";
  context.lineWidth = 1.5;
  context.stroke();
  context.restore();
}

function rectanglesOverlap(first, second, gap = 0) {
  return !(
    first.x + first.width + gap <= second.x ||
    second.x + second.width + gap <= first.x ||
    first.y + first.height + gap <= second.y ||
    second.y + second.height + gap <= first.y
  );
}

function drawGloboHighlights(context, item, padding, plotWidth, plotHeight, minValue, range, compact = false) {
  const firstIndex = 0;
  const lastIndex = item.points.length - 1;
  const maxIndex = item.points.reduce(
    (bestIndex, point, index) => (point.value > item.points[bestIndex].value ? index : bestIndex),
    0,
  );
  const minIndex = item.points.reduce(
    (bestIndex, point, index) => (point.value < item.points[bestIndex].value ? index : bestIndex),
    0,
  );

  const boxes = [];
  [
    { index: firstIndex, caption: "Audiência recebida", offset: -70 },
    { index: lastIndex, caption: "Audiência entregue", offset: -70 },
    { index: maxIndex, caption: "Maior audiência", offset: 34 },
    { index: minIndex, caption: "Menor audiência", offset: 34 },
  ].filter((highlight, position, all) =>
    all.findIndex((candidate) => candidate.index === highlight.index && candidate.caption === highlight.caption) === position,
  ).forEach((highlight) => {
    const box = drawGloboHighlightLabel(
      context,
      item,
      highlight.index,
      highlight.caption,
      highlight.offset,
      padding,
      plotWidth,
      plotHeight,
      minValue,
      range,
      compact,
    );
    if (box) boxes.push(box);
  });
  return boxes;
}

function drawGloboHighlightLabel(
  context,
  item,
  index,
  caption,
  offset,
  padding,
  plotWidth,
  plotHeight,
  minValue,
  range,
  compact = false,
) {
  const point = item.points[index];
  if (!point) return;

  const x = xFor(index, item.points.length, padding.left, plotWidth);
  const y = yFor(point.value, padding.top, plotHeight, minValue, range);
  const valueText = numberFormatter.format(point.value);
  context.font = "800 12px Globotipo, Arial, sans-serif";
  const captionWidth = context.measureText(caption).width;
  context.font = "900 17px Globotipo, Arial, sans-serif";
  const valueWidth = context.measureText(valueText).width;
  const boxWidth = compact ? Math.max(58, valueWidth + 24) : Math.max(138, Math.max(captionWidth, valueWidth) + 24);
  const boxHeight = compact ? 32 : 48;
  const boxX = Math.min(Math.max(4, x - boxWidth / 2), padding.left + plotWidth - boxWidth);
  const boxY = Math.min(Math.max(8, y + offset), padding.top + plotHeight - boxHeight - 4);

  context.save();
  context.shadowColor = "rgba(0, 0, 0, 0.18)";
  context.shadowBlur = 8;
  context.shadowOffsetY = 2;
  context.fillStyle = "rgba(11, 54, 184, 0.96)";
  roundRect(context, boxX, boxY, boxWidth, boxHeight, 8);
  context.fill();
  context.restore();

  context.strokeStyle = "rgba(255, 255, 255, 0.9)";
  context.lineWidth = 1.5;
  roundRect(context, boxX, boxY, boxWidth, boxHeight, 8);
  context.stroke();

  context.textAlign = "center";
  context.fillStyle = "#ffffff";
  context.font = "900 17px Globotipo, Arial, sans-serif";
  context.fillText(valueText, boxX + boxWidth / 2, boxY + (compact ? 22 : 20));
  if (!compact) {
    context.fillStyle = "#dfe8ff";
    context.font = "800 12px Globotipo, Arial, sans-serif";
    context.fillText(caption, boxX + boxWidth / 2, boxY + 38);
  }

  context.fillStyle = "#0b36b8";
  context.beginPath();
  context.arc(x, y, 5.2, 0, Math.PI * 2);
  context.fill();
  context.strokeStyle = "#ffffff";
  context.lineWidth = 2;
  context.stroke();
  return { x: boxX, y: boxY, width: boxWidth, height: boxHeight, caption, value: valueText, time: point.label };
}

function drawPointLabel(context, item, index, padding, plotWidth, plotHeight, minValue, range) {
  const point = item.points[index];
  if (!point) return;

  const x = xFor(index, item.points.length, padding.left, plotWidth);
  const y = yFor(point.value, padding.top, plotHeight, minValue, range);
  const text = numberFormatter.format(point.value);
  const width = context.measureText(text).width + 10;
  const labelY = Math.max(14, y - 16);

  if (item.key === "GLOBO") {
    context.fillStyle = "#9ba0ff";
    roundRect(context, x - width / 2, labelY - 13, width, 20, 5);
    context.fill();
    context.fillStyle = "#0b36b8";
  } else {
    context.fillStyle = item.color;
  }

  context.font = "700 12px Globotipo, Arial, sans-serif";
  context.textAlign = "center";
  context.fillText(text, x, labelY + 1);
}

function xFor(index, total, left, plotWidth) {
  return left + (total === 1 ? plotWidth / 2 : (plotWidth / (total - 1)) * index);
}

function yFor(value, top, plotHeight, minValue, range) {
  return top + plotHeight - ((value - minValue) / range) * plotHeight;
}

function roundRect(context, x, y, width, height, radius) {
  context.beginPath();
  context.moveTo(x + radius, y);
  context.arcTo(x + width, y, x + width, y + height, radius);
  context.arcTo(x + width, y + height, x, y + height, radius);
  context.arcTo(x, y + height, x, y, radius);
  context.arcTo(x, y, x + width, y, radius);
  context.closePath();
}

function pickTimeIndexes(points) {
  if (points.length <= 8) return Array.from({ length: points.length }, (_, index) => index);

  const indexes = new Set();
  points.forEach((point, index) => {
    const [, minute] = String(point.label).split(":").map(Number);
    if (Number.isFinite(minute) && minute % 10 === 0) indexes.add(index);
  });

  indexes.add(0);
  indexes.add(points.length - 1);
  return [...indexes].sort((a, b) => a - b);
}

async function downloadCleanImage() {
  if (!rows.length) {
    downloadWarning.textContent = "Conecte o Realtime antes de gerar os arquivos.";
    statusText.textContent = "";
    return;
  }

  if (!productionFallback.hidden && manualProductionTotal.value.trim() &&
      !Number.isFinite(parseManualProductionDuration(manualProductionTotal.value))) {
    downloadWarning.textContent = "Informe a produção total em MM:SS ou HH:MM:SS.";
    manualProductionTotal.focus();
    return;
  }

  const missingFields = getMissingDownloadFields();
  if (missingFields.length) {
    downloadWarning.textContent = `Preencha ${missingFields.join(" e ")} antes de baixar a imagem.`;
    statusText.textContent = "";
    if (missingFields[0] === "o nome do PROGRAMA") {
      marketBadge.focus();
    } else {
      (presenterName.value === "OUTRO" ? presenterOtherName : presenterName).focus();
    }
    return;
  }

  downloadWarning.textContent = "";

  const width = 1600;
  const height = 900;
  const scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const context = canvas.getContext("2d");
  context.scale(scale, scale);
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  const series = buildSeries(getLabelColumn());

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  await drawExportHeader(context, width);
  await drawTvGloboLogo(context, width);
  await drawExportMetrics(context, series);
  const chartBounds = {
    x: 24,
    y: 240,
    width: width - 48,
    height: 600,
  };
  drawMultiLineChartOnContext(context, series, getActiveHeader(), chartBounds, {
    includeInews: false,
  });
  drawExportFooter(context, width, height);

  const dateStamp = new Date().toISOString().slice(0, 10);
  const programFileSuffix = getProgramFileSuffix();
  const imageDataUrl = canvas.toDataURL("image/png");
  downloadDataUrl(imageDataUrl, `audiencia-previa-${dateStamp}-${programFileSuffix}.png`);

  const htmlCanvas = document.createElement("canvas");
  htmlCanvas.width = canvas.width;
  htmlCanvas.height = canvas.height;
  const htmlContext = htmlCanvas.getContext("2d");
  htmlContext.scale(scale, scale);
  htmlContext.drawImage(canvas, 0, 0, width, height);
  const interactiveHighlights = [];
  drawMultiLineChartOnContext(htmlContext, series, getActiveHeader(), chartBounds, {
    includeInews: false,
    compactHighlights: true,
    highlightCollector: interactiveHighlights,
  });

  const interactiveDots = buildExportInewsDots(series, chartBounds);
  const globotipoFontSource = await getGlobotipoFontSource();
  const interactiveHtml = createInteractiveHtml(
    htmlCanvas.toDataURL("image/png"),
    interactiveDots,
    interactiveHighlights,
    globotipoFontSource,
    series,
  );
  downloadTextFile(
    interactiveHtml,
    `audiencia-previa-interativa-${dateStamp}-${programFileSuffix}.html`,
    "text/html",
  );
}

function getProgramFileSuffix() {
  const programName = canonicalProgramName(marketBadge.value) || "programa";
  return programName
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase() || "programa";
}

async function getGlobotipoFontSource() {
  const fontUrl = new URL("./fonts/GlobotipoCorporativa-Regular.ttf", window.location.href).href;
  try {
    const response = await fetch(fontUrl);
    if (!response.ok) throw new Error("Fonte indisponível");
    const blob = await response.blob();
    return await blobToDataUrl(blob);
  } catch {
    return fontUrl;
  }
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function buildExportInewsDots(series, bounds) {
  if (!inewsItems.length || !series.length) return { dots: [], timeline: null };
  const values = series.flatMap((item) => item.points.map((point) => point.value));
  const maxValue = Math.max(...values);
  const minValue = Math.min(0, ...values);
  const topValue = Math.ceil(maxValue + Math.max(1, maxValue * 0.08));
  const range = topValue - minValue || 1;
  const padding = { top: 36, right: 16, bottom: 48, left: 22 };
  const plotWidth = bounds.width - padding.left - padding.right;
  const plotHeight = bounds.height - padding.top - padding.bottom;
  const longest = series.reduce((winner, item) =>
    item.points.length > winner.points.length ? item : winner,
  );

  const dots = buildInewsDotData(
    series,
    longest.points,
    padding,
    plotWidth,
    plotHeight,
    minValue,
    range,
  ).map((dot) => ({
    ...dot,
    x: bounds.x + dot.x,
    globoY: bounds.y + dot.globoY,
    bandX: bounds.x + dot.bandX,
    bandEndX: bounds.x + dot.bandEndX,
  }));
  const timelineY = getInewsTimelineY(series, padding, plotHeight, minValue, range);
  const globo = series.find((item) => item.key === "GLOBO");
  const globoPoints = (globo?.points ?? []).map((point, index) => ({
    x: bounds.x + xFor(index, globo.points.length, padding.left, plotWidth),
    y: bounds.y + yFor(point.value, padding.top, plotHeight, minValue, range),
  }));
  return {
    dots: dots.map((dot) => ({ ...dot, y: bounds.y + timelineY })),
    timeline: {
      x1: bounds.x + padding.left,
      x2: bounds.x + padding.left + plotWidth,
      y: bounds.y + timelineY,
    },
    globoPoints,
    plot: { y: bounds.y + padding.top, height: plotHeight },
  };
}

function downloadDataUrl(url, filename) {
  const link = document.createElement("a");
  link.download = filename;
  link.href = url;
  link.click();
}

function downloadTextFile(text, filename, type) {
  const url = URL.createObjectURL(new Blob([text], { type: `${type};charset=utf-8` }));
  downloadDataUrl(url, filename);
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function getStrongestAnalysisMovement(points, direction) {
  const trends = buildAudienceTrendMovements(points).filter((movement) =>
    direction === "rise" ? movement.delta > 0 : movement.delta < 0,
  );
  if (trends.length) {
    return trends.reduce((strongest, movement) =>
      Math.abs(movement.delta) > Math.abs(strongest.delta) ? movement : strongest,
    );
  }

  let strongest = null;
  for (let endIndex = 1; endIndex < points.length; endIndex += 1) {
    for (let startIndex = Math.max(0, endIndex - 6); startIndex < endIndex; startIndex += 1) {
      const delta = points[endIndex].value - points[startIndex].value;
      const matchesDirection = direction === "rise" ? delta > 0 : delta < 0;
      if (matchesDirection && (!strongest || Math.abs(delta) > Math.abs(strongest.delta))) {
        strongest = { startIndex, index: endIndex, delta };
      }
    }
  }
  return strongest;
}

function analysisArrow(delta, positiveForGlobo = null) {
  if (positiveForGlobo === null || !Number.isFinite(delta) || Math.abs(delta) < 0.1) return "";
  const favorable = (delta > 0) === positiveForGlobo;
  return `<span class="impact-arrow ${favorable ? "positive" : "negative"}" aria-label="${favorable ? "Favorável" : "Desfavorável"} à Globo">${delta > 0 ? "▲" : "▼"}</span> `;
}

function describeAnalysisMovement(item, movement, direction, positiveForGlobo = null) {
  if (!movement) return direction === "rise"
    ? "não teve uma subida registrada nessa escala"
    : "não teve uma queda registrada nessa escala";
  const start = item.points[movement.startIndex ?? movement.index - 1]?.label ?? "-";
  const end = item.points[movement.index]?.label ?? "-";
  const verb = direction === "rise" ? "subiu" : "caiu";
  return `${analysisArrow(movement.delta, positiveForGlobo)}${verb} ${numberFormatter.format(Math.abs(movement.delta))} pts entre ${escapeHtml(start)} e ${escapeHtml(end)}`;
}

function describeAnalysisBalance(analysis, positiveForGlobo = null) {
  const first = analysis.item.points[0].value;
  const last = analysis.item.points[analysis.item.points.length - 1].value;
  const startEnd = `Foi de ${numberFormatter.format(first)} para ${numberFormatter.format(last)} pts`;
  if (analysis.overall > 0) {
    return `${analysisArrow(analysis.overall, positiveForGlobo)}${startEnd}, ganhando ${numberFormatter.format(analysis.overall)} pts no período`;
  }
  if (analysis.overall < 0) {
    return `${analysisArrow(analysis.overall, positiveForGlobo)}${startEnd}, perdendo ${numberFormatter.format(Math.abs(analysis.overall))} pts no período`;
  }
  return `${startEnd}, sem mudança entre o início e o fim`;
}

function describeAnalysisSeries(analysis, positiveForGlobo = null) {
  if (!analysis) return "Não há dados suficientes para esta série.";
  const movements = [
    analysis.rise && `A maior subida veio quando ${describeAnalysisMovement(analysis.item, analysis.rise, "rise", positiveForGlobo)}.`,
    analysis.fall && `A maior queda ocorreu quando ${describeAnalysisMovement(analysis.item, analysis.fall, "fall", positiveForGlobo)}.`,
  ].filter(Boolean).join(" ");
  return `${describeAnalysisBalance(analysis, positiveForGlobo)}. ${movements || "Não houve subida nem queda no período."}`;
}

function describeSeriesDuringGloboFall(globo, other, label, positiveForGlobo = null) {
  if (!globo?.fall || !other) return "";
  const start = globo.item.points[globo.fall.startIndex ?? globo.fall.index - 1]?.label;
  const end = globo.item.points[globo.fall.index]?.label;
  if (!start || !end) return "";
  const difference = valueAtTime(other.item.points, end) - valueAtTime(other.item.points, start);
  if (!Number.isFinite(difference)) return "";
  if (difference <= -0.1) {
    return ` Na maior queda da Globo, ${analysisArrow(difference, positiveForGlobo)}${label} também recuou ${numberFormatter.format(Math.abs(difference))} pts.`;
  }
  if (difference >= 0.1) {
    return ` Na maior queda da Globo, ${analysisArrow(difference, positiveForGlobo)}${label} subiu ${numberFormatter.format(difference)} pts.`;
  }
  return ` Na maior queda da Globo, ${label} ficou praticamente estável.`;
}

function nearestAnalysisPoint(points, seconds, timelineStart, side = "nearest") {
  let nearest = null;
  let distance = Infinity;
  points.forEach((point) => {
    const pointTime = timeToTimelineSeconds(point.label, timelineStart);
    if (side === "atOrAfter" && pointTime < seconds) return;
    if (side === "atOrBefore" && pointTime > seconds) return;
    const difference = Math.abs(pointTime - seconds);
    if (difference < distance) {
      nearest = { ...point, time: pointTime };
      distance = difference;
    }
  });
  return distance <= 60 ? nearest : null;
}

function describeParallelChange(item, start, end, timelineStart, label, positiveForGlobo = null, showStable = true) {
  if (!item) return "";
  const first = nearestAnalysisPoint(item.points, start, timelineStart);
  const last = nearestAnalysisPoint(item.points, end, timelineStart);
  if (!first || !last || first.time === last.time) return "";
  const delta = last.value - first.value;
  if (Math.abs(delta) < 0.1) return showStable ? `${label} ficou praticamente estável` : "";
  return `${analysisArrow(delta, positiveForGlobo)}${label} ${delta > 0 ? "subiu" : "caiu"} ${numberFormatter.format(Math.abs(delta))} pts`;
}

function describeGloboTleShare(globo, tle, start, end, timelineStart) {
  if (!globo || !tle) return "";
  const globoFirst = nearestAnalysisPoint(globo.item.points, start, timelineStart);
  const globoLast = nearestAnalysisPoint(globo.item.points, end, timelineStart);
  const tleFirst = nearestAnalysisPoint(tle.item.points, start, timelineStart);
  const tleLast = nearestAnalysisPoint(tle.item.points, end, timelineStart);
  if (!globoFirst || !globoLast || !tleFirst?.value || !tleLast?.value) return "";
  const firstShare = globoFirst.value / tleFirst.value;
  const lastShare = globoLast.value / tleLast.value;
  if (!firstShare) return "";
  const delta = (lastShare / firstShare - 1) * 100;
  if (Math.abs(delta) < 0.1) return "a participação da Globo no TLE ficou praticamente estável";
  return `${analysisArrow(delta, true)}a participação da Globo no TLE ${delta > 0 ? "subiu" : "caiu"} ${numberFormatter.format(Math.abs(delta))}%`;
}

function buildRetrancaAnalysis(globo, nic, tle, record, sbt) {
  if (!globo || !inewsItems.length) return "";
  const timelineStart = timeToSeconds(globo.item.points[0].label);
  const timelineEnd = timeToTimelineSeconds(globo.item.points.at(-1).label, timelineStart);
  const intervals = getIntervals().map(({ start, end }) => ({
    start: timeToTimelineSeconds(start, timelineStart),
    end: timeToTimelineSeconds(end, timelineStart),
  }));
  const localWindows = isBomDiaBrasil() ? getLocalParticipations().map(({ start, end }) => ({
    start: timeToTimelineSeconds(start, timelineStart),
    end: timeToTimelineSeconds(end, timelineStart),
  })) : [];
  const timelineItems = inewsItems
    .map((item) => ({
      ...item,
      start: timeToTimelineSeconds(item.time, timelineStart),
      durationSeconds: durationToSeconds(item.duration),
    }))
    .filter((item) =>
      Number.isFinite(item.start) && item.start >= timelineStart && item.start <= timelineEnd &&
      shouldShowInewsItem(item.title) &&
      !intervals.some(({ start, end }) => item.start >= start && item.start < end) &&
      (!isBomDiaBrasil() || localWindows.some(({ start, end }) => item.start >= start && item.start < end)),
    )
    .sort((a, b) => a.start - b.start);
  const events = timelineItems.filter((item) =>
      item.durationSeconds > 0 &&
      !/(^|\W)(vinheta|abertura|passagem|chamada)(\W|$)/i.test(item.title) &&
      item.durationSeconds <= 30 * 60,
    );
  const candidates = [];

  events.forEach((event) => {
    const nextStart = timelineItems.find((item) => item.start > event.start)?.start ?? Infinity;
    const intervalStart = Math.min(
      Infinity,
      ...intervals.filter(({ start }) => start > event.start).map(({ start }) => start),
    );
    const localEnd = isBomDiaBrasil()
      ? localWindows.find(({ start, end }) => event.start >= start && event.start < end)?.end ?? Infinity
      : Infinity;
    const end = Math.min(event.start + event.durationSeconds, nextStart, intervalStart, localEnd, timelineEnd);
    const windows = [
      { phase: "during", start: event.start, end },
      { phase: "after", start: end, end: Math.min(end + 120, nextStart, intervalStart, localEnd, timelineEnd) },
    ];
    windows.forEach((window) => {
      if (window.end - window.start < 60) return;
      const first = nearestAnalysisPoint(globo.item.points, window.start, timelineStart, "atOrAfter");
      const last = nearestAnalysisPoint(globo.item.points, window.end, timelineStart, "atOrBefore");
      if (!first || !last || last.time - first.time < 60) return;
      const delta = last.value - first.value;
      if (Math.abs(delta) < 0.3) return;
      candidates.push({ event, ...window, first, last, delta });
    });
  });

  const selected = [];
  for (const direction of ["rise", "fall"]) {
    const ranked = candidates
      .filter(({ delta }) => direction === "rise" ? delta > 0 : delta < 0)
      .sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
    ranked.forEach((candidate) => {
      if (selected.filter(({ delta }) => direction === "rise" ? delta > 0 : delta < 0).length >= 2) return;
      if (selected.some((item) => item.event === candidate.event ||
        (candidate.start < item.end && item.start < candidate.end))) return;
      selected.push(candidate);
    });
  }
  if (!selected.length) {
    return `<div class="analysis-events"><h3>Retrancas e audiência</h3><p>O iNews foi conectado, mas não há trechos isolados e longos o suficiente para relacionar uma retranca a uma variação de pelo menos 0,3 ponto no gráfico minuto a minuto.</p></div>`;
  }

  const itemsHtml = selected.sort((a, b) => a.start - b.start).map(({ event, phase, first, last, delta }) => {
    const relation = phase === "during" ? "Durante" : "Logo após";
    const parallel = [
      describeParallelChange(nic?.item, first.time, last.time, timelineStart, "o NIC", false, false),
      describeParallelChange(record?.item, first.time, last.time, timelineStart, "a Record", false, false),
      describeParallelChange(sbt?.item, first.time, last.time, timelineStart, "o SBT", false, false),
      describeParallelChange(tle?.item, first.time, last.time, timelineStart, "o TLE", null, false),
      describeGloboTleShare(globo, tle, first.time, last.time, timelineStart),
    ].filter((text) => text && !text.includes("ficou praticamente estável"));
    return `<li><strong>${relation} “${escapeHtml(event.title)}”</strong> (${escapeHtml(first.label)}–${escapeHtml(last.label)}), ` +
      `${analysisArrow(delta, true)}a Globo ${delta > 0 ? "subiu" : "caiu"} ${numberFormatter.format(Math.abs(delta))} pts ` +
      `(${numberFormatter.format(first.value)} para ${numberFormatter.format(last.value)}).` +
      `${parallel.length ? ` No mesmo trecho: ${parallel.join("; ")}.` : ""}</li>`;
  }).join("");
  return `<div class="analysis-events"><h3>Retrancas e audiência</h3><p>Os trechos abaixo são associações de horário, não uma medida do efeito de cada matéria.</p><ul>${itemsHtml}</ul></div>`;
}

function buildAudienceAnalysisHtml(series) {
  const analyzedSeries = series.filter((item) => item.points.length > 1);
  if (!analyzedSeries.length) return "";

  const analyses = analyzedSeries.map((item) => ({
    item,
    rise: getStrongestAnalysisMovement(item.points, "rise"),
    fall: getStrongestAnalysisMovement(item.points, "fall"),
    average: item.points.reduce((sum, point) => sum + point.value, 0) / item.points.length,
    overall: item.points[item.points.length - 1].value - item.points[0].value,
  }));
  const globo = analyses.find(({ item }) => item.key === "GLOBO");
  if (!globo) return "";
  const nic = analyses.find(({ item }) => item.key === "CONTEUDO TV");
  const tle = analyses.find(({ item }) => item.key === "TOTAL LIGADOS");
  const record = analyses.find(({ item }) => item.key === "RECORD");
  const sbt = analyses.find(({ item }) => item.key === "SBT");
  const firstTime = timeToSeconds(globo?.item.points[0]?.label ?? "00:00");
  const lastTime = globo ? timeToTimelineSeconds(globo.item.points.at(-1).label, firstTime) : firstTime;
  const globoFirst = globo.item.points[0].value;
  const globoLast = globo.item.points.at(-1).value;
  const globoMovement = (movement, label) => {
    if (!movement) return "";
    const start = globo.item.points[movement.startIndex ?? movement.index - 1]?.label;
    const end = globo.item.points[movement.index]?.label;
    if (!start || !end) return "";
    return `<li><span>${label}</span><strong>${analysisArrow(movement.delta, true)}${numberFormatter.format(Math.abs(movement.delta))} pts</strong><small>${escapeHtml(start)}–${escapeHtml(end)}</small></li>`;
  };
  const competitorRows = [
    ["NIC", nic], ["Record", record], ["SBT", sbt],
  ].filter(([, analysis]) => analysis).map(([label, analysis]) => `<tr>
    <th scope="row">${label}</th>
    <td>${numberFormatter.format(analysis.item.points[0].value)}</td>
    <td>${numberFormatter.format(analysis.item.points.at(-1).value)}</td>
    <td>${Math.abs(analysis.overall) < 0.1 ? "Estável" : `${analysisArrow(analysis.overall, false)}${analysis.overall > 0 ? "+" : "−"}${numberFormatter.format(Math.abs(analysis.overall))} pts`}</td>
  </tr>`).join("");
  const tleShare = describeGloboTleShare(globo, tle, firstTime, lastTime, firstTime);

  return `<section class="analysis" aria-labelledby="analysis-title">
    <h2 id="analysis-title">Análise de audiência</h2>
    <p class="analysis-intro">Resumo do período, seguido dos movimentos observados nas retrancas. Variações simultâneas não indicam, por si só, causa ou transferência de público.</p>
    <div class="analysis-legend" aria-label="Legenda das setas"><span>${analysisArrow(1, true)} Favorável à Globo</span><span>${analysisArrow(-1, true)} Desfavorável à Globo</span><small>A ponta mostra se o índice subiu ou caiu. Para concorrentes, queda aparece em verde e alta em vermelho; isso não significa que o público migrou entre emissoras. No TLE, a cor acompanha a participação da Globo no total ligado.</small></div>
    <div class="analysis-overview">
      <div class="analysis-card analysis-card-globo">
        <h3>Globo</h3>
        <p class="analysis-card-value">${numberFormatter.format(globoFirst)} <span>→</span> ${numberFormatter.format(globoLast)} <small>pts</small></p>
        <p>${Math.abs(globo.overall) < 0.1 ? "Terminou praticamente no mesmo patamar em que começou." : `${analysisArrow(globo.overall, true)}${globo.overall > 0 ? "Ganhou" : "Perdeu"} ${numberFormatter.format(Math.abs(globo.overall))} pts entre o início e o fim.`}</p>
        <ul class="analysis-movements">${globoMovement(globo.rise, "Maior subida")}${globoMovement(globo.fall, "Maior queda")}</ul>
      </div>
      <div class="analysis-card analysis-card-competition">
        <h3>Concorrência</h3>
        <p>Variação entre a entrada e a saída do programa, em pontos de audiência.</p>
        <table><thead><tr><th>Emissora</th><th>Início</th><th>Fim</th><th>Saldo</th></tr></thead><tbody>${competitorRows}</tbody></table>
      </div>
      <div class="analysis-card analysis-card-tle">
        <h3>Total ligado (TLE)</h3>
        ${tle ? `<p class="analysis-card-value">${numberFormatter.format(tle.item.points[0].value)} <span>→</span> ${numberFormatter.format(tle.item.points.at(-1).value)} <small>pts</small></p><p>${tleShare ? `${tleShare}.` : "Não foi possível calcular a variação da participação da Globo no TLE."}</p><small>O percentual compara a participação da Globo no total ligado no início e no fim.</small>` : "<p>Sem dados de TLE suficientes.</p>"}
      </div>
    </div>
    ${buildRetrancaAnalysis(globo, nic, tle, record, sbt)}
    <details class="analysis-details"><summary>Ver movimentos detalhados por emissora</summary>
      <div class="analysis-summary">
        <p><strong>Globo.</strong> ${describeAnalysisSeries(globo, true)}</p>
        <p><strong>NIC.</strong> ${describeAnalysisSeries(nic, false)}${describeSeriesDuringGloboFall(globo, nic, "o NIC", false)}</p>
        <p><strong>Record.</strong> ${describeAnalysisSeries(record, false)}${describeSeriesDuringGloboFall(globo, record, "a Record", false)}</p>
        <p><strong>SBT.</strong> ${describeAnalysisSeries(sbt, false)}${describeSeriesDuringGloboFall(globo, sbt, "o SBT", false)}</p>
        <p><strong>TLE.</strong> ${describeAnalysisSeries(tle)}${describeSeriesDuringGloboFall(globo, tle, "o TLE")} ${tleShare ? `${tleShare}.` : ""}</p>
      </div>
    </details>
  </section>`;
}

function buildIntervalAnalysisHtml(series) {
  const globo = series.find((item) => item.key === "GLOBO" && item.points.length > 1);
  if (!globo) return "";
  const intervals = getIntervals();
  const timelineStart = timeToSeconds(globo.points[0].label);
  const timelineEnd = timeToTimelineSeconds(globo.points.at(-1).label, timelineStart);
  const competitors = [
    ["CONTEUDO TV", "o NIC"],
    ["RECORD", "a Record"],
    ["SBT", "o SBT"],
  ];
  const analyzed = intervals.map((interval) => {
    const start = timeToTimelineSeconds(interval.start, timelineStart);
    const end = timeToTimelineSeconds(interval.end, timelineStart);
    if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start ||
        start < timelineStart || end > timelineEnd) return null;
    const first = nearestAnalysisPoint(globo.points, start, timelineStart, "atOrAfter");
    const last = nearestAnalysisPoint(globo.points, end, timelineStart, "atOrBefore");
    if (!first || !last || last.time <= first.time) return null;
    const delta = last.value - first.value;
    const parallel = competitors.map(([key, label]) =>
      describeParallelChange(series.find((item) => item.key === key), first.time, last.time, timelineStart, label, false),
    ).filter(Boolean);
    const tle = series.find((item) => item.key === "TOTAL LIGADOS");
    const tleShare = tle ? describeGloboTleShare({ item: globo }, { item: tle }, first.time, last.time, timelineStart) : "";
    const globoMovement = Math.abs(delta) < 0.1
      ? "ficou praticamente estável"
      : `${delta > 0 ? "subiu" : "caiu"} ${numberFormatter.format(Math.abs(delta))} pts`;
    return {
      delta,
      html: `<li><strong>Intervalo ${escapeHtml(interval.start)}–${escapeHtml(interval.end)}.</strong> ` +
        `Entre ${escapeHtml(first.label)} e ${escapeHtml(last.label)}, ` +
        `${analysisArrow(delta, true)}a Globo ${globoMovement} ` +
        `(${numberFormatter.format(first.value)} para ${numberFormatter.format(last.value)} pts).` +
        `${parallel.length ? ` No mesmo período, ${parallel.join("; ")}.` : ""}` +
        `${tleShare ? ` ${tleShare}.` : ""}</li>`,
    };
  }).filter(Boolean);
  let content;
  if (!intervals.length) {
    content = "<p>Nenhum intervalo foi informado para este programa.</p>";
  } else if (!analyzed.length) {
    content = "<p>Os intervalos informados não têm pontos minuto a minuto suficientes dentro dos seus limites para uma comparação segura.</p>";
  } else {
    const falls = analyzed.filter(({ delta }) => delta <= -0.1).length;
    const rises = analyzed.filter(({ delta }) => delta >= 0.1).length;
    const stable = analyzed.length - falls - rises;
    const summary = `Nos ${analyzed.length} intervalos com dados suficientes, a Globo caiu em ${falls}, subiu em ${rises} e ficou praticamente estável em ${stable}.`;
    content = `<p>${summary}</p><ol>${analyzed.map(({ html }) => html).join("")}</ol>`;
  }
  return `<section class="analysis interval-analysis" aria-labelledby="interval-analysis-title">
    <h2 id="interval-analysis-title">Análise de intervalos</h2>
    <p class="analysis-intro">Comparação dos pontos de audiência dentro de cada intervalo informado. Os valores refletem a resolução minuto a minuto da base, não os segundos exatos do início e do fim. Coincidência de horário não demonstra causa e efeito.</p>
    ${content}
  </section>`;
}

function createInteractiveHtml(imageDataUrl, interactiveData, highlights, globotipoFontSource, series) {
  const safeDots = JSON.stringify(interactiveData.dots).replace(/</g, "\\u003c");
  const safeHighlights = JSON.stringify(highlights).replace(/</g, "\\u003c");
  const timeline = interactiveData.timeline;
  const globoPolyline = (interactiveData.globoPoints ?? [])
    .map((point) => `${point.x},${point.y}`)
    .join(" ");
  const interactivePlot = interactiveData.plot ?? { y: 0, height: 0 };
  const pageTitle = `Desempenho prévio — ${marketBadge.value || "Programa"}`;
  const legendX = 24 + 22 + (getLocalParticipations().length ? 292 : 126);
  const audienceAnalysisHtml = buildAudienceAnalysisHtml(series);
  const intervalAnalysisHtml = buildIntervalAnalysisHtml(series);
  return `<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(pageTitle)}</title>
  <style>
    @font-face{font-family:Globotipo;src:url("${globotipoFontSource}") format("truetype");font-weight:400;font-style:normal;font-display:swap}*{box-sizing:border-box}body{margin:0;background:#eef1f6;font-family:Arial,sans-serif;color:#17213b}.page{width:min(1600px,100%);margin:0 auto;padding:18px}.hint{margin:0 0 10px;font-family:Globotipo,Arial,sans-serif;font-size:14px;color:#526079}.visual{position:relative;width:100%;aspect-ratio:16/9;background:#fff;box-shadow:0 12px 38px #17213b22}.visual>img,.visual>svg{position:absolute;inset:0;width:100%;height:100%;display:block}.dot{cursor:pointer;stroke:#fff;stroke-width:2;transition:filter .15s ease}.dot:hover,.dot:focus{filter:drop-shadow(0 2px 3px #0005);outline:none}.retranca-band{fill:#6fbfff;opacity:0;transition:opacity .15s ease}.retranca-band.active{opacity:.22}.globo-highlight{fill:none;stroke:#6fbfff;stroke-width:8;stroke-linecap:round;stroke-linejoin:round;opacity:0;filter:drop-shadow(0 0 5px #6fbfff);transition:opacity .15s ease}.connector{stroke:#6fbfff;stroke-width:2;stroke-dasharray:7 5;opacity:0;transition:opacity .15s ease}.globo-highlight.active,.connector.active{opacity:1}.tooltip{position:fixed;z-index:5;display:none;max-width:420px;padding:10px 12px;border-radius:8px;background:#071d55;color:#fff;box-shadow:0 8px 24px #0004;pointer-events:none}.tooltip strong{display:block;margin-bottom:4px}.tooltip span{font-size:12px;color:#cdddff}.analysis{margin-top:22px;padding:24px;border-radius:16px;background:#fff;box-shadow:0 8px 28px #17213b18;font-family:Globotipo,Arial,sans-serif}.analysis h2{margin:0 0 8px;color:#0b36a8}.analysis-intro{margin:0 0 20px;color:#526079}.analysis-table-wrap{overflow-x:auto}.analysis table{width:100%;border-collapse:collapse}.analysis th,.analysis td{padding:12px 14px;border-bottom:1px solid #e3e8f1;text-align:left}.analysis th{background:#f3f6fb;color:#34415c}.analysis .rise{color:#1684c5;font-weight:700}.analysis .fall{color:#d64036;font-weight:700}.analysis-summary{margin-top:20px;padding:16px 18px;border-left:4px solid #6fbfff;border-radius:8px;background:#eef7ff;line-height:1.55}.analysis-summary p{margin:0 0 8px}.analysis-summary p:last-child{margin-bottom:0}@media(max-width:700px){.analysis{padding:18px}.analysis th,.analysis td{padding:10px 8px;font-size:13px}}
    .analysis-events{margin-top:22px}.analysis-events h3{margin:0 0 8px;color:#0b36a8}.analysis-events p{margin:0 0 12px;color:#526079}.analysis-events ul{margin:0;padding-left:22px;line-height:1.55}.analysis-events li+li{margin-top:10px}.impact-arrow{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;margin:0 4px 0 2px;border-radius:6px;font-family:Arial,sans-serif;font-size:18px;font-weight:900;line-height:1;vertical-align:-3px;white-space:nowrap}.impact-arrow.positive{color:#08742d;background:#d6f4df}.impact-arrow.negative{color:#c72d24;background:#ffe1dd}.analysis-legend{display:flex;align-items:center;flex-wrap:wrap;gap:8px 22px;margin:0 0 14px;color:#34415c}.analysis-legend>span{display:inline-flex;align-items:center;font-weight:700}.analysis-legend small{flex-basis:100%;font-size:12px;color:#526079}.interval-analysis>p:not(.analysis-intro){line-height:1.55}.interval-analysis ol{margin:14px 0 0;padding-left:24px;line-height:1.6}.interval-analysis li+li{margin-top:14px}
    .analysis-overview{display:grid;grid-template-columns:minmax(0,1.1fr) minmax(0,1.3fr) minmax(0,1fr);gap:14px}.analysis-card{min-width:0;padding:18px;border:1px solid #dfe7f2;border-radius:12px;background:#f8fbff;line-height:1.45}.analysis-card h3{margin:0 0 10px;color:#0b36a8;font-size:17px}.analysis-card p{margin:0 0 10px}.analysis-card p:last-child{margin-bottom:0}.analysis-card-value{font-size:24px;font-weight:700;color:#0b36a8}.analysis-card-value span{color:#8c9ab0}.analysis-card-value small{font-size:14px}.analysis-card table{font-size:13px}.analysis-card th,.analysis-card td{padding:8px 5px}.analysis-card td:last-child{white-space:nowrap}.analysis-card-tle>small{display:block;color:#526079}.analysis-movements{list-style:none;margin:14px 0 0;padding:0}.analysis-movements li{display:grid;grid-template-columns:1fr auto;gap:2px 8px;padding:8px 0;border-top:1px solid #dfe7f2}.analysis-movements li strong{white-space:nowrap}.analysis-movements li small{grid-column:1/-1;color:#526079}.analysis-details{margin-top:22px}.analysis-details summary{cursor:pointer;color:#0b36a8;font-weight:700}.analysis-details .analysis-summary{margin-top:12px}@media(max-width:900px){.analysis-overview{grid-template-columns:1fr}.analysis-card table{font-size:14px}}
    .highlight-hit{fill:transparent;stroke:transparent;stroke-width:2;cursor:help}.highlight-hit:hover,.highlight-hit:focus{stroke:#6fbfff;outline:none}
  </style>
</head>
<body>
  <main class="page">
    <p class="hint">Passe o mouse ou use Tab sobre bolinhas e índices para ver os detalhes.</p>
    <div class="visual">
      <img src="${imageDataUrl}" alt="Gráfico de desempenho prévio">
      <svg viewBox="0 0 1600 900" role="group" aria-label="Retrancas e índices interativos sobre o gráfico" id="markers">
        <rect id="retrancaBand" class="retranca-band" y="${interactivePlot.y}" height="${interactivePlot.height}"></rect>
        <polyline id="globoHighlight" class="globo-highlight" points="${globoPolyline}"></polyline>
        <line id="connector" class="connector"></line>
        ${timeline ? `<line x1="${timeline.x1}" y1="${timeline.y}" x2="${timeline.x2}" y2="${timeline.y}" stroke="#b9ddf8" stroke-width="2"></line>` : ""}
        <g aria-label="Legenda: Retrancas">
          <rect x="${legendX}" y="240" width="118" height="28" fill="#eef7ff"></rect>
          <circle cx="${legendX + 17}" cy="254" r="5" fill="${inewsDotColor}"></circle>
          <text x="${legendX + 70}" y="258" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" fill="#27648f">Retrancas</text>
        </g>
        <g id="highlightHitLayer" aria-label="Índices de destaque"></g>
      </svg>
    </div>
    ${audienceAnalysisHtml}
    ${intervalAnalysisHtml}
  </main>
  <div class="tooltip" id="tooltip" role="tooltip"><strong></strong><span></span></div>
  <script>
    const dots=${safeDots};
    const highlights=${safeHighlights};
    const svg=document.querySelector('#markers');
    const tooltip=document.querySelector('#tooltip');
    const title=tooltip.querySelector('strong');
    const details=tooltip.querySelector('span');
    const globoHighlight=document.querySelector('#globoHighlight');
    const connector=document.querySelector('#connector');
    const retrancaBand=document.querySelector('#retrancaBand');
    const ns='http://www.w3.org/2000/svg';
    function show(event,dot){title.textContent=dot.title;details.textContent=dot.time+(dot.duration?' · '+dot.duration:'');retrancaBand.setAttribute('x',dot.bandX);retrancaBand.setAttribute('width',Math.max(6,dot.bandEndX-dot.bandX));retrancaBand.classList.add('active');connector.setAttribute('x1',dot.x);connector.setAttribute('y1',dot.y);connector.setAttribute('x2',dot.x);connector.setAttribute('y2',dot.globoY);connector.classList.add('active');globoHighlight.classList.add('active');tooltip.style.display='block';move(event)}
    function move(event){const x='clientX'in event?event.clientX:event.target.getBoundingClientRect().left;const y='clientY'in event?event.clientY:event.target.getBoundingClientRect().top;tooltip.style.left=Math.min(innerWidth-tooltip.offsetWidth-12,x+14)+'px';tooltip.style.top=Math.max(8,y-tooltip.offsetHeight-12)+'px'}
    function hide(){tooltip.style.display='none';retrancaBand.classList.remove('active');connector.classList.remove('active');globoHighlight.classList.remove('active')}
    dots.forEach(dot=>{const circle=document.createElementNS(ns,'circle');const baseRadius=dot.radius||6;circle.setAttribute('class','dot');circle.setAttribute('cx',dot.x);circle.setAttribute('cy',dot.y);circle.setAttribute('r',baseRadius);circle.setAttribute('fill',dot.color);circle.setAttribute('tabindex','0');circle.setAttribute('aria-label',dot.title+' às '+dot.time);circle.addEventListener('mouseenter',event=>{circle.setAttribute('r',baseRadius+3);show(event,dot)});circle.addEventListener('mousemove',move);circle.addEventListener('mouseleave',()=>{circle.setAttribute('r',baseRadius);hide()});circle.addEventListener('focus',event=>{circle.setAttribute('r',baseRadius+3);show(event,dot)});circle.addEventListener('blur',()=>{circle.setAttribute('r',baseRadius);hide()});svg.appendChild(circle)});
    const highlightLayer=document.querySelector('#highlightHitLayer');
    highlights.forEach(hit=>{const rect=document.createElementNS(ns,'rect');rect.setAttribute('class','highlight-hit');rect.setAttribute('x',hit.x);rect.setAttribute('y',hit.y);rect.setAttribute('width',hit.width);rect.setAttribute('height',hit.height);rect.setAttribute('rx','8');rect.setAttribute('tabindex','0');rect.setAttribute('aria-label',hit.caption+': '+hit.value+' pontos às '+hit.time);const showDetail=event=>{title.textContent=hit.caption;details.textContent=hit.value+' pontos · '+hit.time;tooltip.style.display='block';move(event)};const hideDetail=()=>{tooltip.style.display='none'};rect.addEventListener('mouseenter',showDetail);rect.addEventListener('mousemove',move);rect.addEventListener('mouseleave',hideDetail);rect.addEventListener('focus',showDetail);rect.addEventListener('blur',hideDetail);highlightLayer.appendChild(rect)});
  <\/script>
</body>
</html>`;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

async function drawExportHeader(context, width) {
  const headerY = 26;
  const headerHeight = 42;
  const headerCenterY = headerY + headerHeight / 2;

  context.fillStyle = "#0b36a8";
  context.font = "900 32px Globotipo, Arial, sans-serif";
  context.textBaseline = "middle";
  context.textAlign = "left";
  const title = "DESEMPENHO PRÉVIO";
  context.fillText(title, 28, 48);
  const titleWidth = context.measureText(title).width;

  const badgeText = marketBadge.value || "PROGRAMA";
  const programLogo = getProgramLogoSource(badgeText);
  const programName = normalize(badgeText).replace(/\s+/g, "");
  const isGloboEsporte = programName === "globoesporte" || programName === "ge";
  const isGloboComunidade = programName === "globocomunidade" || programName === "gco";
  let badgeX = 28 + titleWidth + (isGloboEsporte ? -28 : isGloboComunidade ? 4 : 14);
  let badgeWidth = 0;
  let presenterAnchorWidth = 0;

  if (programLogo) {
    const logoSize = getProgramLogoSize(badgeText);
    badgeWidth = logoSize.width;
    presenterAnchorWidth = getPresenterLogoAnchorWidth(badgeText, badgeWidth);
    await drawLogoImageContain(
      context,
      programLogo,
      badgeX,
      headerY,
      logoSize.width,
      logoSize.height,
    );
  } else {
    context.font = "900 21px Globotipo, Arial, sans-serif";
    badgeWidth = Math.max(72, context.measureText(badgeText).width + 36);
    roundRect(context, badgeX, headerY, badgeWidth, headerHeight, 21);
    context.fill();
    context.fillStyle = "#ffffff";
    context.textAlign = "center";
    context.fillText(badgeText, badgeX + badgeWidth / 2, headerCenterY);
    presenterAnchorWidth = badgeWidth;
  }

  const presenters = getSelectedPresenterCards();
  if (presenters.length) {
    let presenterX = badgeX + presenterAnchorWidth + 14;
    for (const presenter of presenters) {
      presenterX += await drawPresenterBadge(context, presenter, presenterX, 25) + 14;
    }
  }

  context.fillStyle = "#0b36a8";
  context.textBaseline = "alphabetic";
  context.font = "900 15px Globotipo, Arial, sans-serif";
  context.font = "900 21px Globotipo, Arial, sans-serif";
  context.textAlign = "center";
  context.fillText(reportDate.textContent, 94, 142);
  roundRect(context, 28, 154, 132, 44, 22);
  context.fill();
  context.fillStyle = "#ffffff";
  context.font = "900 17px Globotipo, Arial, sans-serif";
  context.fillText(reportRange.textContent, 94, 182);
}

async function drawPresenterBadge(context, presenter, x, y) {
  const hasPhoto = Boolean(presenter.photoKey && window.presenterData?.[presenter.photoKey]);
  const photoSize = 54;
  const pillX = x + (hasPhoto ? 30 : 0);
  context.font = "700 14px Globotipo, Arial, sans-serif";
  const textWidth = context.measureText(presenter.name).width;
  const pillWidth = Math.min(320, Math.max(hasPhoto ? 95 : 72, Math.ceil(textWidth + (hasPhoto ? 54 : 36))));
  context.fillStyle = "#0b36a8";
  roundRect(context, pillX, y + 5, pillWidth, 44, 22);
  context.fill();

  context.fillStyle = "#ffffff";
  context.textAlign = "left";
  context.textBaseline = "middle";
  context.fillText(presenter.name, pillX + (hasPhoto ? 31 : 18), y + 27, pillWidth - (hasPhoto ? 49 : 36));

  if (hasPhoto) {
    await new Promise((resolve) => {
      const image = new Image();
      image.onload = () => {
        context.save();
        context.beginPath();
        context.arc(x + photoSize / 2, y + photoSize / 2, photoSize / 2, 0, Math.PI * 2);
        context.clip();
        context.drawImage(image, x, y, photoSize, photoSize);
        context.restore();
        context.strokeStyle = "#ffffff";
        context.lineWidth = 3;
        context.beginPath();
        context.arc(x + photoSize / 2, y + photoSize / 2, photoSize / 2, 0, Math.PI * 2);
        context.stroke();
        resolve();
      };
      image.onerror = resolve;
      image.src = window.presenterData[presenter.photoKey];
    });
  }
  return pillX + pillWidth - x;
}

async function drawExportMetrics(context, series) {
  const startX = 190;
  const top = 130;
  const entries = buildExportMetricEntries(series);
  let x = startX;

  for (const item of entries) {
    const calculatedAverage =
      item.points.reduce((sum, point) => sum + point.value, 0) / item.points.length;
    const average = summary[item.header]?.average;
    const share = summary[item.header]?.share;

    context.fillStyle = item.color;
    context.beginPath();
    context.arc(x + 25, top + 32, 24, 0, Math.PI * 2);
    context.fill();

    const logoSource = getLogoSource(item);
    if (logoSource) {
      await drawLogoImage(context, logoSource, x + 8, top + 15, 34, 34);
    } else {
      context.fillStyle = "#ffffff";
      context.font = "900 13px Globotipo, Arial, sans-serif";
      context.textAlign = "center";
      context.fillText(item.mark, x + 25, top + 37);
    }

    context.fillStyle = item.color;
    context.fillRect(x + 60, top + 2, 3, 74);
    context.font = "900 17px Globotipo, Arial, sans-serif";
    context.textAlign = "left";
    context.fillText(formatMetricValue(average, calculatedAverage), x + 72, top + 18);
    if (item.key === "GLOBO") {
      drawExportHistoricalBadge(context, x + 134, top + 1);
      context.textAlign = "left";
    }
    context.font = "400 12px Globotipo, Arial, sans-serif";
    context.fillStyle = "#9b9b9b";
    context.fillText("Audiência", x + 72, top + 36);
    context.fillStyle = item.color;
    context.font = "900 15px Globotipo, Arial, sans-serif";
    context.fillText(formatShareValue(share), x + 72, top + 57);
    context.fillStyle = "#9b9b9b";
    context.font = "400 12px Globotipo, Arial, sans-serif";
    context.fillText("Share", x + 72, top + 74);
    x += item.key === "GLOBO" ? 280 : 170;
  }
}

function buildExportMetricEntries(series) {
  return [...series].slice(0, 7);
}

function drawExportHistoricalBadge(context, x, y) {
  const badge = getHistoricalBadgeData();
  if (!badge) return;

  const width = 120;
  const height = 40;
  context.fillStyle = badge.tone === "negative" ? "#e95145" : "#6fbfff";
  roundRect(context, x, y, width, height, 9);
  context.fill();
  context.fillStyle = "#ffffff";
  context.textAlign = "center";
  context.font = "700 11px Globotipo, Arial, sans-serif";
  context.fillText("Média histórica", x + width / 2, y + 15);
  context.font = "900 11px Globotipo, Arial, sans-serif";
  context.fillText(`${badge.audience} pts | ${badge.variation}`, x + width / 2, y + 31);
}

function drawExportHistoricalMetric(context, x, top) {
  const audience = historicalAudience.value.trim()
    ? formatSignedMetricValue(historicalAudience.value)
    : "-";
  const variation = historicalShare.value.trim()
    ? formatVariationInput(historicalShare.value)
    : "-";
  const variationNumber = parseNumber(variation);
  const accentColor = getSignedColor(variationNumber);

  context.fillStyle = accentColor;
  context.fillRect(x + 12, top + 2, 3, 74);
  context.font = "900 17px Globotipo, Arial, sans-serif";
  context.textAlign = "left";
  context.fillText(audience, x + 24, top + 18);
  context.font = "400 12px Globotipo, Arial, sans-serif";
  context.fillStyle = "#9b9b9b";
  context.fillText("Média histórica", x + 24, top + 36);
  context.fillStyle = accentColor;
  context.font = "900 15px Globotipo, Arial, sans-serif";
  context.fillText(variation, x + 24, top + 57);
  context.fillStyle = "#9b9b9b";
  context.font = "400 12px Globotipo, Arial, sans-serif";
  context.fillText("Variação", x + 24, top + 74);
}

function drawLogoImage(context, src, x, y, width, height) {
  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => {
      context.imageSmoothingEnabled = true;
      context.imageSmoothingQuality = "high";
      context.drawImage(image, x, y, width, height);
      resolve();
    };
    image.onerror = resolve;
    image.src = src;
  });
}

function drawLogoImageContain(context, src, x, y, width, height) {
  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => {
      context.imageSmoothingEnabled = true;
      context.imageSmoothingQuality = "high";
      const scale = Math.min(width / image.width, height / image.height);
      const drawWidth = image.width * scale;
      const drawHeight = image.height * scale;
      const drawX = x + (width - drawWidth) / 2;
      const drawY = y + (height - drawHeight) / 2;
      context.drawImage(image, drawX, drawY, drawWidth, drawHeight);
      resolve();
    };
    image.onerror = resolve;
    image.src = src;
  });
}

function getLogoSource(item) {
  return item.imageKey ? window.logoData?.[item.imageKey] : null;
}

async function drawTvGloboLogo(context, width) {
  const src = window.logoData?.TVGLOBO;
  if (!src) return;
  await drawLogoImage(context, src, width - 202, 26, 174, 42);
}

function getMissingDownloadFields() {
  const program = normalize(marketBadge.value);
  const missing = [];
  if (!program || program === "programa") missing.push("o nome do PROGRAMA");
  if (!getPresenterSelection(presenterName, presenterOtherName)) {
    missing.push("o nome do Apresentador");
  }
  return missing;
}

function getProgramLogoSource(programName) {
  const normalizedName = normalize(programName);
  const compactName = normalizedName.replace(/\s+/g, "");
  if (compactName === "df1") return window.logoData?.PROGRAMA_DF1;
  if (compactName === "df2") return window.logoData?.PROGRAMA_DF2;
  if (normalizedName === "bom dia df" || compactName === "bddf") {
    return window.logoData?.PROGRAMA_BOM_DIA_DF;
  }
  if (normalizedName === "globo comunidade" || normalizedName === "gco") {
    return window.logoData?.PROGRAMA_GLOBO_COMUNIDADE;
  }
  if (normalizedName === "globo esporte" || compactName === "ge") {
    return window.logoData?.PROGRAMA_GLOBO_ESPORTE;
  }
  if (normalizedName === "boletim df2") return window.logoData?.PROGRAMA_BOLETIM_DF2;
  if (normalizedName === "bom dia brasil" || normalizedName === "bdbr") {
    return window.logoData?.PROGRAMA_BOM_DIA_BRASIL;
  }
  return null;
}

function getProgramLogoSize(programName) {
  const normalizedName = normalize(programName);
  const compactName = normalizedName.replace(/\s+/g, "");
  if (normalizedName === "bom dia df" || compactName === "bddf") return { width: 153, height: 42 };
  if (normalizedName === "globo esporte" || compactName === "ge") return { width: 174, height: 42 };
  if (normalizedName === "boletim df2") return { width: 174, height: 42 };
  if (normalizedName === "bom dia brasil" || normalizedName === "bdbr") {
    return { width: 174, height: 42 };
  }
  if (normalizedName === "globo comunidade" || normalizedName === "gco") {
    return { width: 130, height: 42 };
  }
  return { width: 174, height: 42 };
}

function getPresenterLogoAnchorWidth(programName, logoWidth) {
  const normalizedName = normalize(programName);
  const compactName = normalizedName.replace(/\s+/g, "");
  if (normalizedName === "globo esporte" || compactName === "ge") return 136;
  if (normalizedName === "bom dia brasil" || compactName === "bdbr") return 127;
  return logoWidth;
}

function drawExportFooter(context, width, height) {
  context.fillStyle = "#8a8a8a";
  context.font = "italic 12px Globotipo, Arial, sans-serif";
  context.textAlign = "left";
  context.fillText(
    "Modelo desenvolvido pela área de Programação da TV Globo DF",
    28,
    height - 18,
  );
  context.textAlign = "right";
  context.fillText(
    "Fonte: Ibope. DF. Dados prévios. 1 pt de audiência domiciliar no DF equivale a 10.034 domicílios.",
    width - 28,
    height - 18,
  );
}

function drawExportHistorical(context) {
  const audience = historicalAudience.value.trim();
  const share = historicalShare.value.trim();
  if (!audience && !share) return;

  context.fillStyle = "#8a8a8a";
  context.font = "900 12px Globotipo, Arial, sans-serif";
  context.textAlign = "center";
  context.fillText("Média histórica", 94, 184);

  drawHistoricalValue(context, audience || "-", 52, 207);
  context.fillStyle = "#9b9b9b";
  context.font = "400 11px Globotipo, Arial, sans-serif";
  context.fillText("Audiência", 52, 225);

  drawHistoricalValue(context, share || "-", 136, 207);
  context.fillStyle = "#9b9b9b";
  context.font = "400 11px Globotipo, Arial, sans-serif";
  context.fillText("Variação", 136, 225);
}

function drawHistoricalValue(context, value, x, y) {
  const parsed = parseNumber(value);
  context.fillStyle = getSignedColor(parsed);
  context.font = "900 17px Globotipo, Arial, sans-serif";
  context.textAlign = "center";
  context.fillText(value, x, y);
}

function getSignedColor(value) {
  if (!Number.isFinite(value) || value === 0) return "#0b36a8";
  return value < 0 ? "#e95145" : "#6fbfff";
}

function updateSignedInput(input) {
  const value = parseNumber(input.value);
  input.classList.toggle("negative", Number.isFinite(value) && value < 0);
  input.classList.toggle("positive", Number.isFinite(value) && value > 0);
}

function formatVariationInput(value) {
  const number = parseNumber(value);
  if (!Number.isFinite(number)) return value.trim();

  const formatted = numberFormatter.format(Math.abs(number));
  const sign = number > 0 ? "+" : number < 0 ? "-" : "";
  return `${sign}${formatted}%`;
}

function formatSignedMetricValue(value) {
  const number = parseNumber(value);
  if (!Number.isFinite(number)) return value.trim();

  const formatted = numberFormatter.format(Math.abs(number));
  const sign = number > 0 ? "+" : number < 0 ? "-" : "";
  return `${sign}${formatted}`;
}

downloadImageButton.addEventListener("click", downloadCleanImage);

const approximateProgramWindows = {
  "BOM DIA DF": ["06:00:00", "08:30:00"],
  "BOM DIA BRASIL": ["08:30:00", "09:30:00"],
  DF1: ["11:45:00", "13:00:00"],
  "GLOBO ESPORTE": ["13:00:00", "13:25:00"],
  "BOLETIM DF2": ["17:00:00", "17:20:00"],
  DF2: ["19:00:00", "19:45:00"],
  "GLOBO COMUNIDADE": ["06:00:00", "08:00:00"],
};

function getMaestroProgramBounds(programName, sourceRows) {
  const maestroName = getExhibitionProgramName(programName);
  const production = sourceRows
    .filter((row) =>
      getExhibitionProgramName(getExhibitionField(row, "Programa")) === maestroName &&
      /^PD\d+$/i.test(String(getExhibitionField(row, "Segmento")).trim()),
    )
    .map((row) => ({
      start: timeToSeconds(getExhibitionField(row, "HoraInicio")),
      duration: durationToSeconds(getExhibitionField(row, "Duracao")),
    }))
    .filter(({ start, duration }) => Number.isFinite(start) && duration > 0);
  if (!production.length) return null;
  const start = Math.min(...production.map((block) => block.start));
  const end = Math.max(...production.map((block) => block.start + block.duration));
  return { start, end, source: "maestro" };
}

function getProgramSelectionBounds(programName, maestroRows) {
  const exact = getMaestroProgramBounds(programName, maestroRows);
  if (exact) return exact;
  const approximate = approximateProgramWindows[programName];
  return approximate
    ? { start: timeToSeconds(approximate[0]), end: timeToSeconds(approximate[1]), source: "approximate" }
    : null;
}

function getAudienceRowsForSelection(parsed, selection) {
  const minuteHeader = parsed.parsedHeaders.find((header) => normalize(header) === "min") ?? "MIN";
  return parsed.parsedRows.filter((row) => {
    const minute = timeToTimelineSeconds(row[minuteHeader], selection.start);
    return Number.isFinite(minute) && minute < selection.end && minute + 60 > selection.start;
  });
}

function askProgramSelection(maestroRows) {
  return new Promise((resolve) => {
    programSelectionForm.reset();
    otherProgramFields.hidden = true;
    const updateSelection = () => {
      const name = programSelection.value;
      const isOther = name === "OUTRO";
      otherProgramFields.hidden = !isOther;
      otherProgramName.required = isOther;
      otherProgramStart.required = isOther;
      otherProgramEnd.required = isOther;
      const bounds = getProgramSelectionBounds(name, maestroRows);
      programSelectionHint.textContent = isOther
        ? "Informe o nome e os horários exatos de início e fim."
        : bounds?.source === "maestro"
          ? `Maestro: ${secondsToClock(bounds.start)} a ${secondsToClock(bounds.end)}.`
          : bounds
            ? `Sem horário exato no Maestro. Faixa aproximada: ${secondsToClock(bounds.start)} a ${secondsToClock(bounds.end)}.`
            : "Selecione o programa para recortar a audiência.";
    };
    const finish = (result) => {
      programSelection.removeEventListener("change", updateSelection);
      programSelectionForm.removeEventListener("submit", submit);
      cancelProgramSelection.removeEventListener("click", cancel);
      programSelectionDialog.removeEventListener("cancel", cancel);
      if (programSelectionDialog.open) programSelectionDialog.close();
      resolve(result);
    };
    const cancel = (event) => {
      event?.preventDefault();
      finish(null);
    };
    const submit = (event) => {
      event.preventDefault();
      const name = programSelection.value;
      if (!name) return;
      if (name === "OUTRO") {
        const start = timeToSeconds(otherProgramStart.value);
        let end = timeToSeconds(otherProgramEnd.value);
        if (!Number.isFinite(start) || !Number.isFinite(end) || !otherProgramName.value.trim()) return;
        if (end <= start) end += 86400;
        finish({ name: otherProgramName.value.trim().toUpperCase(), start, end, source: "manual" });
        return;
      }
      finish({ name, ...getProgramSelectionBounds(name, maestroRows) });
    };
    programSelection.addEventListener("change", updateSelection);
    programSelectionForm.addEventListener("submit", submit);
    cancelProgramSelection.addEventListener("click", cancel);
    programSelectionDialog.addEventListener("cancel", cancel);
    updateSelection();
    programSelectionDialog.showModal();
  });
}

function applyAudienceSelection(parsed, selection) {
  downloadWarning.textContent = "";
  if (!parsed.parsedHeaders.some((header) => normalize(header) === "min") ||
      !parsed.parsedHeaders.some((header) => normalize(header).includes("globo"))) {
    throw new Error("Não encontrei as colunas de minuto e audiência da Globo no Realtime.");
  }
  const filteredRows = getAudienceRowsForSelection(parsed, selection);
  if (!filteredRows.length) {
    throw new Error(`Não há minutos de audiência para ${selection.name} no período selecionado.`);
  }
  headers = parsed.parsedHeaders;
  rows = filteredRows;
  metadata = {
    ...parsed.parsedMetadata,
    intervalo: `${secondsToClock(selection.start)} - ${secondsToClock(selection.end)}`,
  };
  const tleHeader = headers.find((header) => normalize(header) === "total ligados");
  const tleValues = filteredRows.map((row) => parseSeriesNumber(row[tleHeader])).filter(Number.isFinite);
  const tleAverage = tleValues.length
    ? tleValues.reduce((sum, value) => sum + value, 0) / tleValues.length
    : NaN;
  summary = {};
  headers.filter((header) => header !== "MIN").forEach((header) => {
    const values = filteredRows.map((row) => parseSeriesNumber(row[header])).filter(Number.isFinite);
    const average = values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : NaN;
    summary[header] = {
      average,
      share: Number.isFinite(average) && tleAverage > 0 ? average / tleAverage * 100 : NaN,
    };
  });
  marketBadge.value = canonicalProgramName(selection.name);
  marketBadge.dataset.userEdited = selection.source === "manual" ? "true" : "auto";
  resizeMarketBadge();
  updateImportWorkflow();
  updateChart();
}

function getRealtimeBounds() {
  const labelKey = getLabelColumn();
  const first = timeToSeconds(rows[0]?.[labelKey]);
  const times = rows
    .map((row) => timeToTimelineSeconds(row[labelKey], first))
    .filter(Number.isFinite)
    .sort((a, b) => a - b);
  return times.length ? { start: times[0], end: times[times.length - 1] } : null;
}

function applyProgramNameFromRealtime() {
  const bounds = getRealtimeBounds();
  if (!bounds) return;
  if (isSundayDate(metadata.data) && bounds.start >= 6 * 3600 && bounds.start <= 8 * 3600) {
    marketBadge.value = "GLOBO COMUNIDADE";
    marketBadge.dataset.userEdited = "auto";
    resizeMarketBadge();
    return;
  }
  const candidates = [
    { name: "BOM DIA DF", start: 6 * 3600, end: 8 * 3600 + 30 * 60 },
    { name: "BOM DIA BRASIL", start: 8 * 3600 + 30 * 60, end: 9 * 3600 + 30 * 60 },
    { name: "DF1", start: 11 * 3600 + 45 * 60, end: 13 * 3600 },
    { name: "GLOBO ESPORTE", start: 13 * 3600, end: 13 * 3600 + 25 * 60 },
  ]
    .map((candidate) => ({
      ...candidate,
      startDifference: Math.abs(bounds.start - candidate.start),
      endDifference: Math.abs(bounds.end - candidate.end),
    }))
    .filter((candidate) => candidate.startDifference <= 20 * 60 && candidate.endDifference <= 20 * 60)
    .sort(
      (a, b) =>
        a.startDifference + a.endDifference - (b.startDifference + b.endDifference),
    );
  if (!candidates.length) return;

  marketBadge.value = candidates[0].name;
  marketBadge.dataset.userEdited = "auto";
  resizeMarketBadge();
}

function getSelectedPresenter() {
  const firstPresenter = getPresenterSelection(presenterName, presenterOtherName);
  const secondPresenter = getPresenterSelection(presenterName2, presenterOtherName2);
  return [firstPresenter, secondPresenter].filter(Boolean).join(" E ");
}

function getSelectedPresenterCards() {
  const photoKeys = {
    "FRED FERREIRA": "fred",
    "LUÍSA DOYLE": "luisa",
    "HERALDO PEREIRA": "heraldo",
    "NATALIA GODOY": "natalia",
    "ANDRÉ BARROSO": "andreb",
    "KARINA AZEVEDO": "karina",
    "ANTÔNIO DE CASTRO": "antonio",
    "RAIANE SENA": "raiane",
  };
  return [
    [presenterName, presenterOtherName],
    [presenterName2, presenterOtherName2],
  ]
    .map(([select, otherInput]) => ({
      name: getPresenterSelection(select, otherInput),
      photoKey: photoKeys[select.value] ?? null,
    }))
    .filter((presenter) => presenter.name);
}

function renderPresenterPreview() {
  presenterPreview.replaceChildren();
  getSelectedPresenterCards().forEach((presenter) => {
    const card = document.createElement("div");
    card.className = "presenter-preview-card";
    const photo = presenter.photoKey && window.presenterData?.[presenter.photoKey];
    if (photo) {
      const image = document.createElement("img");
      image.src = photo;
      image.alt = "";
      card.appendChild(image);
    } else {
      card.classList.add("without-photo");
    }
    const name = document.createElement("span");
    name.textContent = presenter.name;
    card.appendChild(name);
    presenterPreview.appendChild(card);
  });
}

function getPresenterSelection(select, otherInput) {
  return select.value === "OUTRO" ? otherInput.value.trim().toUpperCase() : select.value;
}

function isMaestroCsv(text) {
  const delimiter = detectDelimiter(text);
  return String(text).replace(/^\uFEFF/, "").split(/\r?\n/).slice(0, 25).some((line) => {
    const columns = parseCsvLine(line, delimiter).map(normalize);
    return columns.includes("programa") && columns.includes("segmento") && columns.includes("horainicio");
  });
}

function localParticipationClock(value) {
  if (typeof value === "number" && Number.isFinite(value) && value >= 0 && value < 1) {
    return secondsToClock(value * 86400);
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return [value.getHours(), value.getMinutes(), value.getSeconds()]
      .map((part) => String(part).padStart(2, "0")).join(":");
  }
  const text = String(value ?? "").trim();
  const match = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match || Number(match[1]) > 23 || Number(match[2]) > 59 || Number(match[3] ?? 0) > 59) return "";
  return [match[1], match[2], match[3] ?? "00"]
    .map((part) => part.padStart(2, "0")).join(":");
}

async function parseLocalParticipationWorkbook(file) {
  if (!window.XLSX) throw new Error("Leitor de planilhas indisponível. Verifique a conexão e recarregue a página.");
  const workbook = window.XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
  const sheetName = workbook.SheetNames.find((name) => normalize(name).includes("entradas bdb"));
  if (!sheetName) throw new Error("A planilha precisa ter a aba 'Entradas BDB'.");
  const grid = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: "" });
  const headerIndex = grid.findIndex((row) => normalize(row[1]) === "entrada" &&
    (normalize(row[2]) === "saida" || normalize(row[2]).startsWith("sa")));
  if (headerIndex < 0) throw new Error("Não encontrei as colunas Entrada e Saída na aba 'Entradas BDB'.");
  const seen = new Set();
  const periods = [];
  grid.slice(headerIndex + 1).forEach((row) => {
    const start = localParticipationClock(row[1]);
    const end = localParticipationClock(row[2]);
    if (!start || !end) return;
    const duration = timeToSeconds(end) - timeToSeconds(start);
    if (duration <= 0 || duration > 30 * 60) return;
    const key = `${start}|${end}`;
    if (seen.has(key)) return;
    seen.add(key);
    periods.push({ start, end });
  });
  if (!periods.length) throw new Error("Não encontrei pares válidos de Entrada e Saída na planilha.");
  return periods.sort((a, b) => timeToSeconds(a.start) - timeToSeconds(b.start));
}

async function handleBaseFilesSelection(event) {
  const files = [...event.target.files];
  if (!files.length) return;
  baseStatus.textContent = "Lendo bases...";
  downloadWarning.textContent = "";

  try {
    const audienceFiles = [];
    const maestroFiles = [];
    const inewsFiles = [];
    const localParticipationFiles = [];
    for (const file of files) {
      if (/\.pdf$/i.test(file.name)) {
        inewsFiles.push(file);
      } else if (/\.xlsx$/i.test(file.name)) {
        localParticipationFiles.push(file);
      } else if (/\.csv$/i.test(file.name)) {
        const text = await file.text();
        (isMaestroCsv(text) ? maestroFiles : audienceFiles).push({ file, text });
      } else {
        throw new Error(`Formato não reconhecido: ${file.name}. Use CSV, PDF ou XLSX.`);
      }
    }
    if ([audienceFiles, maestroFiles, inewsFiles, localParticipationFiles].some((group) => group.length > 1)) {
      throw new Error("Selecione no máximo um Realtime, um Maestro, um iNews e uma planilha de participação local por envio.");
    }
    if (!rows.length && !audienceFiles.length) {
      throw new Error("Inclua a base de audiência minuto a minuto para montar o gráfico.");
    }
    let audiencePreview = null;
    if (audienceFiles.length) {
      audiencePreview = adaptDailyAudienceCsv(
        parseCsv(audienceFiles[0].text, "auto"),
        audienceFiles[0].file.name,
      );
      if (!audiencePreview.parsedHeaders.some((header) => normalize(header) === "min") ||
          !audiencePreview.parsedHeaders.some((header) => normalize(header).includes("globo"))) {
        throw new Error("O Realtime precisa conter minuto e audiência da Globo.");
      }
    }
    const parsedMaestro = maestroFiles.length ? parseExhibitionCsv(maestroFiles[0].text) : null;
    const selection = audiencePreview ? await askProgramSelection(parsedMaestro ?? []) : null;
    if (audiencePreview && !selection) {
      baseStatus.textContent = "Seleção cancelada. Nenhuma base nova foi aplicada.";
      return;
    }
    if (audiencePreview && !getAudienceRowsForSelection(audiencePreview, selection).length) {
      throw new Error(`Não há minutos de audiência para ${selection.name} no período selecionado.`);
    }

    if (audiencePreview) {
      manualProductionTotal.value = "";
      manualProductionTotal.setCustomValidity("");
      inewsItems = [];
      inewsSourceItems = [];
      exhibitionRows = [];
      exhibitionFileLoaded = false;
      inewsFileLoaded = false;
      localParticipationFileLoaded = false;
      intervalList.replaceChildren();
      addIntervalRow();
      localParticipationList.replaceChildren();
      addLocalParticipationRow();
      delete marketBadge.dataset.userEdited;
      marketBadge.value = "PROGRAMA";
      fullAudienceData = audiencePreview;
      selectedProgram = selection;
      applyAudienceSelection(fullAudienceData, selectedProgram);
    }
    if (parsedMaestro) {
      exhibitionRows = parsedMaestro;
      exhibitionFileLoaded = true;
      if (selectedProgram && selectedProgram.source === "approximate" && fullAudienceData) {
        const exact = getMaestroProgramBounds(selectedProgram.name, exhibitionRows);
        if (exact) {
          selectedProgram = { name: selectedProgram.name, ...exact };
          applyAudienceSelection(fullAudienceData, selectedProgram);
        }
      }
      refreshExhibitionIntervals();
    }
    const optionalErrors = [];
    if (localParticipationFiles.length) {
      try {
        if (!isBomDiaBrasil()) throw new Error("a planilha de participação local é destinada ao Bom Dia Brasil.");
        const periods = await parseLocalParticipationWorkbook(localParticipationFiles[0]);
        localParticipationList.replaceChildren();
        periods.forEach(({ start, end }) => addLocalParticipationRow(start, end));
        localParticipationFileLoaded = true;
      } catch (error) {
        optionalErrors.push(`Participação local não aplicada: ${error.message}`);
      }
    }
    if (inewsFiles.length) {
      const previousInews = inewsSourceItems;
      const previousInewsLoaded = inewsFileLoaded;
      try {
        const parsed = await parseInewsPdf(inewsFiles[0]);
        if (!parsed.some((item) => shouldShowInewsItem(item.title))) {
          throw new Error("O PDF do iNews não contém retrancas com horários válidos.");
        }
        inewsSourceItems = parsed;
        inewsFileLoaded = true;
      } catch (error) {
        inewsSourceItems = previousInews;
        inewsFileLoaded = previousInewsLoaded;
        optionalErrors.push(`iNews não aplicado: ${error.message}`);
      }
    }
    updateImportWorkflow();
    updateChart();
    if (optionalErrors.length) baseStatus.textContent += ` · ${optionalErrors.join(" · ")}`;
  } catch (error) {
    baseStatus.textContent = error.message;
    if (rows.length) updateChart();
  } finally {
    event.target.value = "";
  }
}

baseFilesInput.addEventListener("change", handleBaseFilesSelection);

function groupPdfItemsByRow(items) {
  const rows = [];
  [...items]
    .filter((item) => item.str?.trim())
    .sort((a, b) => b.transform[5] - a.transform[5] || a.transform[4] - b.transform[4])
    .forEach((item) => {
      const y = item.transform[5];
      let row = rows.find((candidate) => Math.abs(candidate.y - y) <= 1.5);
      if (!row) {
        row = { y, items: [] };
        rows.push(row);
      }
      row.items.push({ text: item.str.trim(), x: item.transform[4] });
    });
  rows.forEach((row) => row.items.sort((a, b) => a.x - b.x));
  return rows;
}

function getPdfColumnX(headerRow, name) {
  return headerRow.items.find((item) => normalize(item.text) === normalize(name))?.x;
}

function getPdfCellText(row, fromX, toX) {
  return row.items
    .filter((item) => item.x >= fromX && item.x < toX)
    .map((item) => item.text)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

async function parseInewsPdf(file) {
  if (!window.pdfjsLib) {
    throw new Error("Não foi possível carregar o leitor de PDF. Verifique a conexão e tente novamente.");
  }
  window.pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  const documentTask = window.pdfjsLib.getDocument({ data: await file.arrayBuffer() });
  const pdf = await documentTask.promise;
  const extracted = [];
  let retrancaPage = null;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const pageRows = groupPdfItemsByRow(textContent.items);
    const retrancaHeader = pageRows.find((row) =>
      Number.isFinite(getPdfColumnX(row, "RETRANCA")) &&
      Number.isFinite(getPdfColumnX(row, "tMat")));
    const tempoHeader = pageRows.find((row) =>
      Number.isFinite(getPdfColumnX(row, "TEMPO")) &&
      Number.isFinite(getPdfColumnX(row, "STATUS")));

    if (retrancaHeader) {
      const retrancaX = getPdfColumnX(retrancaHeader, "RETRANCA");
      const repX = getPdfColumnX(retrancaHeader, "Rep") ?? getPdfColumnX(retrancaHeader, "LOC");
      const tMatX = getPdfColumnX(retrancaHeader, "tMat");
      if (![retrancaX, repX, tMatX].every(Number.isFinite)) continue;
      retrancaPage = pageRows
        .filter((row) => row.y < retrancaHeader.y - 2)
        .map((row) => ({
          y: row.y,
          title: getPdfCellText(row, retrancaX - 3, repX - 1),
          duration: getPdfCellText(row, tMatX - 3, Infinity).match(/\d{1,2}:\d{2}(?::\d{2})?/)?.[0] ?? "",
        }));
    }

    // O PDF do iNews pode imprimir as colunas largas em páginas consecutivas:
    // retranca/tMat na primeira e TEMPO na segunda, com as linhas na mesma altura.
    if (tempoHeader && retrancaPage) {
      const timeRows = pageRows.filter((row) => row.y < tempoHeader.y - 2);
      for (const item of retrancaPage) {
        const matchingRow = timeRows.find((row) => Math.abs(row.y - item.y) <= 3);
        // Em algumas linhas o PDF une MODI, APV e TEMPO em um único item de texto.
        const time = matchingRow?.items.map((part) => part.text).join(" ")
          .match(/\d{1,2}:\d{2}:\d{2}/)?.[0];
        if (item.title && time) extracted.push({ title: item.title, time, duration: item.duration });
      }
      retrancaPage = null;
    }
  }

  return selectInewsMirror(extracted);
}

function refreshInewsItems() {
  if (!inewsSourceItems.length) {
    inewsItems = [];
    return;
  }
  const aligned = isBomDiaBrasil()
    ? alignBdbInewsToLocalParticipations(inewsSourceItems)
    : alignInewsTimelineToRealtime(inewsSourceItems);
  const uniqueItems = new Map();
  aligned
    .filter((item) => shouldShowInewsItem(item.title))
    .forEach((item) => uniqueItems.set(`${normalize(item.title)}|${item.time}`, item));
  const timelineStart = getRealtimeBounds()?.start ?? 0;
  inewsItems = [...uniqueItems.values()].sort(
    (a, b) => timeToTimelineSeconds(a.time, timelineStart) - timeToTimelineSeconds(b.time, timelineStart),
  );
}

function alignBdbInewsToLocalParticipations(items) {
  const windows = getLocalParticipations()
    .map(({ start, end }) => ({ start: timeToSeconds(start), end: timeToSeconds(end) }))
    .filter(({ start, end }) => Number.isFinite(start) && Number.isFinite(end) && end > start)
    .sort((a, b) => a.start - b.start);
  if (!windows.length) return [];

  const blocks = [];
  let block = null;
  let section = "none";
  const closeBlock = (end) => {
    if (block?.items.length) {
      block.end = Number.isFinite(end) ? end : Math.max(...block.items.map((item) =>
        timeToSeconds(item.time) + durationToSeconds(item.duration),
      ));
      blocks.push(block);
    }
    block = null;
  };
  items.forEach((item) => {
    const title = normalize(item.title);
    const time = timeToSeconds(item.time);
    if (title.includes("bloco local")) {
      closeBlock(time);
      block = { start: time, items: [] };
      section = "local";
    } else if (title.includes("intervalo")) {
      closeBlock(time);
      section = "after-interval";
    } else if (title.includes("bloco rede") || title.includes("encerramento")) {
      closeBlock(time);
      section = "network";
    } else if (shouldShowInewsItem(item.title) && Number.isFinite(time)) {
      if (section === "after-interval") {
        block = { start: time, items: [] };
        section = "local";
      }
      if (section !== "local" || !block) return;
      block.items.push(item);
    }
  });
  closeBlock(NaN);

  // O espelho pode misturar tempo decorrido e horário de relógio: reserve primeiro
  // as faixas que coincidem com os blocos em horário absoluto.
  const assignedWindows = new Set();
  const matchedBlocks = [];
  blocks.filter((block) => block.start >= 4 * 3600).forEach((source) => {
    const match = windows
      .map((window, index) => ({
        index,
        overlap: Math.max(0, Math.min(window.end, source.end) - Math.max(window.start, source.start)),
      }))
      .sort((a, b) => b.overlap - a.overlap)[0];
    if (!match?.overlap) return;
    assignedWindows.add(match.index);
    matchedBlocks.push({ source, target: windows[match.index], absolute: true });
  });
  const freeWindows = windows.filter((_, index) => !assignedWindows.has(index));
  blocks.filter((block) => block.start < 4 * 3600).slice(0, freeWindows.length)
    .forEach((source, index) => matchedBlocks.push({ source, target: freeWindows[index], absolute: false }));

  return matchedBlocks.flatMap(({ source, target, absolute }) => {
    const sourceSpan = source.end - source.start;
    if (!Number.isFinite(sourceSpan) || sourceSpan <= 0) return [];
    const scale = absolute ? 1 : (target.end - target.start) / sourceSpan;
    return source.items.map((item) => {
      const sourceTime = timeToSeconds(item.time);
      const mappedTime = absolute
        ? sourceTime
        : Math.min(target.end, Math.max(target.start,
            target.start + (sourceTime - source.start) * scale));
      const duration = Math.min(target.end - mappedTime, durationToSeconds(item.duration) * scale);
      return {
        ...item,
        time: secondsToClock(mappedTime),
        duration: item.duration ? formatProductionDuration(duration) : "",
      };
    }).filter((item) => {
      const time = timeToSeconds(item.time);
      return time >= target.start && time < target.end;
    });
  });
}

function alignInewsTimelineToRealtime(items) {
  const realtimeBounds = getRealtimeBounds();
  if (!realtimeBounds || !items.length) return items;

  const itemTimes = items.map((item) => timeToSeconds(item.time)).filter(Number.isFinite);
  if (!itemTimes.length) return items;

  const maestroAligned = alignElapsedInewsToMaestro(items, realtimeBounds);
  if (maestroAligned) return maestroAligned;

  // Alguns espelhos começam com tempo decorrido (00:00:00) e passam a mostrar
  // horários de relógio no mesmo PDF. Ajuste somente o trecho inicial.
  if (itemTimes[0] === 0) {
    const anchorCandidates = [];
    const maestroIntervals = getIntervals().map(({ start, end }) => ({
      start: timeToSeconds(start),
      duration: timeToSeconds(end) - timeToSeconds(start),
    }));
    items.forEach((item) => {
      if (!normalize(item.title).includes("intervalo")) return;
      const elapsed = timeToSeconds(item.time);
      const duration = durationToSeconds(item.duration);
      if (!Number.isFinite(elapsed) || elapsed >= realtimeBounds.start || !duration) return;
      maestroIntervals.forEach((interval) => {
        if (!Number.isFinite(interval.start) || Math.abs(interval.duration - duration) > 5) return;
        const anchor = interval.start - elapsed;
        // O CSV de audiência tem resolução de um minuto; aceite uma pequena
        // variação no início, sem casar intervalos de outro programa.
        if (Math.abs(anchor - realtimeBounds.start) <= 120) anchorCandidates.push(anchor);
      });
    });
    anchorCandidates.sort((a, b) => a - b);
    const anchor = anchorCandidates.length
      ? anchorCandidates[Math.floor(anchorCandidates.length / 2)]
      : realtimeBounds.start;
    let usingClockTime = false;
    return items.map((item) => {
      const time = timeToSeconds(item.time);
      if (!Number.isFinite(time)) return item;
      if (time >= realtimeBounds.start) usingClockTime = true;
      return usingClockTime
        ? item
        : { ...item, time: secondsToClock(anchor + time) };
    });
  }

  const itemsInsideRealtime = itemTimes.filter(
    (time) => time >= realtimeBounds.start && time <= realtimeBounds.end,
  ).length;
  if (itemsInsideRealtime > 0) return items;

  const earliestTime = Math.min(...itemTimes);
  const latestTime = Math.max(...itemTimes);
  const realtimeDuration = Math.max(0, realtimeBounds.end - realtimeBounds.start);
  const looksLikeElapsedTime =
    earliestTime < 4 * 3600 &&
    latestTime <= Math.max(4 * 3600, realtimeDuration + 2 * 3600);
  if (!looksLikeElapsedTime) return items;

  return items.map((item) => {
    const elapsedSeconds = timeToSeconds(item.time);
    return Number.isFinite(elapsedSeconds)
      ? { ...item, time: secondsToClock(realtimeBounds.start + elapsedSeconds) }
      : item;
  });
}

function alignElapsedInewsToMaestro(items, realtimeBounds) {
  const elapsedTimes = items.map((item) => timeToSeconds(item.time));
  if (elapsedTimes.some((time) => !Number.isFinite(time) || time >= 4 * 3600)) return null;

  const breakIndexes = items
    .map((item, index) => (normalize(item.title).includes("intervalo") ? index : -1))
    .filter((index) => index >= 0);
  if (!breakIndexes.length) return null;

  const maestroIntervals = identifyProgramIntervals(exhibitionRows, marketBadge.value)
    .map(({ start, end }) => ({ start: timeToSeconds(start), end: timeToSeconds(end) }));
  if (maestroIntervals.length !== breakIndexes.length) return null;

  const productionBlocks = getMaestroProductionBlocks();
  const targetBlocks = productionBlocks.length === breakIndexes.length + 1
    ? productionBlocks
    : [
        { start: realtimeBounds.start, end: maestroIntervals[0].start },
        ...maestroIntervals.slice(0, -1).map((interval, index) => ({
          start: interval.end,
          end: maestroIntervals[index + 1].start,
        })),
        { start: maestroIntervals[maestroIntervals.length - 1].end, end: realtimeBounds.end + 60 },
      ];
  if (targetBlocks.some(({ start, end }) => !Number.isFinite(start) || !Number.isFinite(end) || end <= start)) {
    return null;
  }

  const aligned = items.map((item) => ({ ...item }));
  breakIndexes.forEach((index, breakNumber) => {
    aligned[index].time = secondsToClock(maestroIntervals[breakNumber].start);
  });

  targetBlocks.forEach((block, blockNumber) => {
    const firstIndex = blockNumber === 0 ? 0 : breakIndexes[blockNumber - 1] + 1;
    const endIndex = blockNumber < breakIndexes.length ? breakIndexes[blockNumber] : items.length;
    if (firstIndex >= endIndex) return;
    // O bloco após o intervalo começa na primeira retranca seguinte (geralmente
    // a VINHETA VOLTA), não no início do intervalo anterior. O tempo decorrido
    // do iNews já inclui a duração do intervalo.
    const sourceStart = elapsedTimes[firstIndex];
    const sourceEnd = blockNumber < breakIndexes.length
      ? elapsedTimes[breakIndexes[blockNumber]]
      : Math.max(...items.slice(firstIndex).map((item, index) =>
          elapsedTimes[firstIndex + index] + durationToSeconds(item.duration),
        ));
    const sourceSpan = sourceEnd - sourceStart;
    const scale = sourceSpan > 0 ? (block.end - block.start) / sourceSpan : 0;
    for (let index = firstIndex; index < endIndex; index += 1) {
      const mappedTime = Math.min(block.end, Math.max(block.start,
        block.start + (elapsedTimes[index] - sourceStart) * scale));
      aligned[index].time = secondsToClock(mappedTime);
      if (items[index].duration) {
        const duration = Math.min(block.end - mappedTime, durationToSeconds(items[index].duration) * scale);
        aligned[index].duration = formatProductionDuration(duration);
      }
    }
  });
  return aligned;
}

function askInewsJournal() {
  return new Promise((resolve) => {
    const buttons = [...inewsJournalDialog.querySelectorAll("[data-inews-journal]")];
    const finish = (choice) => {
      buttons.forEach((button) => button.removeEventListener("click", handleChoice));
      inewsJournalDialog.close();
      resolve(choice);
    };
    const handleChoice = (event) => finish(event.currentTarget.dataset.inewsJournal);
    buttons.forEach((button) => button.addEventListener("click", handleChoice));
    inewsJournalDialog.addEventListener("cancel", (event) => event.preventDefault(), { once: true });
    inewsJournalDialog.showModal();
  });
}

async function selectInewsMirror(items) {
  const closingIndexes = items
    .map((item, index) => (normalize(item.title).includes("encerr") ? index : -1))
    .filter((index) => index >= 0);
  const firstClosing = closingIndexes[0] ?? items.length;
  const bulletinStart = items.findIndex(
    (item, index) => index > firstClosing && normalize(item.title).includes("boletim"),
  );
  const bulletinClosing = closingIndexes.find((index) => index > bulletinStart);
  const hasDf2AndBulletin =
    firstClosing < items.length && bulletinStart > firstClosing && Number.isFinite(bulletinClosing);

  if (!hasDf2AndBulletin) return items.slice(0, firstClosing);

  const choice = await askInewsJournal();
  let selectedItems =
    choice === "boletim-df2"
      ? items.slice(bulletinStart, bulletinClosing)
      : items.slice(0, firstClosing);
  if (choice === "boletim-df2") {
    selectedItems = alignBulletinMirrorToRealtime(selectedItems);
  }
  const journalName = choice === "boletim-df2" ? "BOLETIM DF2" : "DF2";
  marketBadge.value = journalName;
  marketBadge.dataset.userEdited = "true";
  resizeMarketBadge();
  return selectedItems;
}

function getRealtimeStartSeconds() {
  return getRealtimeBounds()?.start;
}

function alignBulletinMirrorToRealtime(items) {
  const openingIndex = items.findIndex((item) => normalize(item.title).includes("abertura"));
  if (openingIndex < 0) {
    throw new Error("Não encontrei a VINHETA ABERTURA no espelho do Boletim DF2.");
  }

  const realtimeStart = getRealtimeStartSeconds();
  if (!Number.isFinite(realtimeStart)) {
    throw new Error("Importe primeiro o CSV de realtime para alinhar o Boletim DF2.");
  }

  let currentTime = realtimeStart;
  return items.slice(openingIndex).map((item) => {
    const alignedItem = { ...item, time: secondsToClock(currentTime) };
    currentTime += durationToSeconds(item.duration);
    return alignedItem;
  });
}

function shouldShowInewsItem(title) {
  const normalizedTitle = normalize(title);
  if (
    normalizedTitle.includes("intervalo") ||
    normalizedTitle.includes("encerr") ||
    normalizedTitle.includes("bloco local") ||
    normalizedTitle.includes("bloco rede") ||
    /^\s*\(av\)/.test(normalizedTitle)
  ) {
    return false;
  }
  return !/(^|[^a-z0-9])vh([^a-z0-9]|$)/.test(normalizedTitle);
}

function clearTextOnFocus(input, afterClear) {
  input.addEventListener("focus", () => {
    input.value = "";
    afterClear?.();
  });
}

clearTextOnFocus(marketBadge, () => {
  marketBadge.dataset.userEdited = "true";
  resizeMarketBadge();
});

marketBadge.addEventListener("input", () => {
  downloadWarning.textContent = "";
  marketBadge.dataset.userEdited = "true";
  marketBadge.value = marketBadge.value.toUpperCase();
  resizeMarketBadge();
});

marketBadge.addEventListener("change", () => {
  marketBadge.value = canonicalProgramName(marketBadge.value);
  resizeMarketBadge();
  manualProductionTotal.value = "";
  manualProductionTotal.setCustomValidity("");
  if (exhibitionFileLoaded) refreshExhibitionIntervals();
  if (rows.length) updateChart();
});

function setupPresenterSelection(select, otherInput) {
  select.addEventListener("change", () => {
    downloadWarning.textContent = "";
    const isOther = select.value === "OUTRO";
    otherInput.hidden = !isOther;
    if (isOther) {
      otherInput.focus();
    } else {
      otherInput.value = "";
    }
    renderPresenterPreview();
  });

  otherInput.addEventListener("input", () => {
    downloadWarning.textContent = "";
    otherInput.value = otherInput.value.toUpperCase();
    renderPresenterPreview();
  });
}

setupPresenterSelection(presenterName, presenterOtherName);
setupPresenterSelection(presenterName2, presenterOtherName2);
renderPresenterPreview();

addIntervalButton.addEventListener("click", () => {
  addIntervalRow();
});

addLocalParticipationButton.addEventListener("click", () => {
  addLocalParticipationRow();
});

manualProductionTotal.addEventListener("input", () => {
  const duration = parseManualProductionDuration(manualProductionTotal.value);
  manualProductionTotal.setCustomValidity(
    manualProductionTotal.value.trim() && !Number.isFinite(duration)
      ? "Use MM:SS ou HH:MM:SS."
      : "",
  );
  downloadWarning.textContent = "";
  if (rows.length) updateChart();
});

[historicalAudience, historicalShare].forEach((input) => {
  input.addEventListener("input", () => {
    updateSignedInput(input);
    if (rows.length) updateChart();
  });
});

historicalShare.addEventListener("blur", () => {
  historicalShare.value = formatVariationInput(historicalShare.value);
  updateSignedInput(historicalShare);
  if (rows.length) updateChart();
});

window.addEventListener("resize", () => {
  if (!rows.length || !emptyState.hidden) return;
  updateChart();
});

resizeMarketBadge();
addIntervalRow();
addLocalParticipationRow();
