/**
 * 从飞书多维表格同步数据到 PPT1.html 和 page-data.js
 *
 * 用法（在 DataProcess 目录下）：node scripts/sync_pages.mjs
 *
 * 飞书表格格式需与 page_elements_with_defaults.xlsx 一致：
 * 列：Field Key, Element, HTML / Chart Hook, Data Type, 说明, 默认值/列1, 列2, 列3, 列4, 列5
 * - 封面 (Sheet): bfd5ff，页面1: zBgHxt，页面2: jonA37（可通过环境变量覆盖）
 * - 页面 3～13：通过环境变量 PAGE3_SHEET_ID … PAGE13_SHEET_ID 指定 sheet ID，或飞书中将对应工作表命名为 "3"～"13" 或 "Page3"～"Page13"
 * - 若飞书表仍为旧格式（封面为 标题/简介/年份…，页面1 为 标题/总客流/客流趋势…，页面2 为多表），脚本会自动按旧逻辑解析
 */
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { feishuFetch } from "./feishu_auth.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");

const {
  FEISHU_SHEET_TOKEN = "KcGusWwSuhBMcot3Deucjd9Inix",
  COVER_RANGE = "bfd5ff!A1:J20",
  PAGE1_RANGE = "zBgHxt!A1:Z200",
  PAGE2_SHEET_NAME = "2",
  PAGE2_SHEET_ID = "jonA37",
  PAGE2_RANGE = "A1:Z200",
  HTML_TARGET = "Maanshan3DMap/PPT1.html",
  PAGE_DATA_FILE = "Maanshan3DMap/page-data.js"
} = process.env;

const PAGE_SHEET_IDS = [
  process.env.PAGE3_SHEET_ID,
  process.env.PAGE4_SHEET_ID,
  process.env.PAGE5_SHEET_ID,
  process.env.PAGE6_SHEET_ID,
  process.env.PAGE7_SHEET_ID,
  process.env.PAGE8_SHEET_ID,
  process.env.PAGE9_SHEET_ID,
  process.env.PAGE10_SHEET_ID,
  process.env.PAGE11_SHEET_ID,
  process.env.PAGE12_SHEET_ID,
  process.env.PAGE13_SHEET_ID
].map(s => (s || "").trim()).filter(Boolean);

const HTML_PATH = path.join(rootDir, HTML_TARGET);
const PAGE_DATA_PATH = path.join(rootDir, PAGE_DATA_FILE);

async function fetchRange(range) {
  const url = `https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/${FEISHU_SHEET_TOKEN}/values/${encodeURIComponent(
    range
  )}`;
  const resp = await feishuFetch(url);
  const json = await resp.json();
  if (json.code !== 0) {
    throw new Error(`飞书 API 返回异常：${json.code} - ${json.msg}`);
  }
  return json.data?.valueRange?.values || [];
}

let cachedSheetsMeta = null;

async function fetchSheetsMeta() {
  if (cachedSheetsMeta) return cachedSheetsMeta;
  const url = `https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/${FEISHU_SHEET_TOKEN}/sheets/query`;
  const resp = await feishuFetch(url);
  const json = await resp.json();
  if (json.code !== 0) {
    throw new Error(`获取 sheet 列表失败：${json.code} - ${json.msg}`);
  }
  cachedSheetsMeta = json.data?.sheets ?? [];
  return cachedSheetsMeta;
}

async function resolveSheetRange({ sheetTitle, sheetId, cellRange }) {
  const cleanRange = (cellRange || "A1:Z200").replace(/^!/, "");
  if (sheetId) return `${sheetId}!${cleanRange}`;
  if (!sheetTitle) throw new Error("需要提供 sheet 的标题或 sheet_id");
  const sheets = await fetchSheetsMeta();
  const target = sheets.find(sheet => (sheet?.title ?? "").trim() === sheetTitle.trim());
  if (!target) {
    const available = sheets.map(s => s?.title ?? "").join(", ");
    throw new Error(`未找到名称为 "${sheetTitle}" 的 sheet，可用：${available}`);
  }
  return `${target.sheet_id}!${cleanRange}`;
}

async function resolvePageRange(pageNum) {
  const n = pageNum - 3;
  if (n >= 0 && n < PAGE_SHEET_IDS.length && PAGE_SHEET_IDS[n]) {
    return `${PAGE_SHEET_IDS[n]}!A1:Z300`;
  }
  const sheets = await fetchSheetsMeta();
  const titles = [String(pageNum), `Page${pageNum}`];
  const target = sheets.find(sheet => titles.includes((sheet?.title ?? "").trim()));
  if (target) return `${target.sheet_id}!A1:Z300`;
  return null;
}

function getColumnIndices(headers) {
  const h = (headers || []).map(c => (c || "").toString().trim());
  const find = (...candidates) => {
    for (const c of candidates) {
      const i = h.findIndex(
        x => x === c || x.toLowerCase() === c.toLowerCase() || (c.length > 2 && x.includes(c))
      );
      if (i !== -1) return i;
    }
    return -1;
  };
  return {
    fieldKey: find("Field Key", "FieldKey") >= 0 ? find("Field Key", "FieldKey") : 0,
    element: find("Element") >= 0 ? find("Element") : 1,
    htmlHook: find("HTML / Chart Hook", "HTML / Chart Hook", "HTML") >= 0 ? find("HTML / Chart Hook", "HTML") : 2,
    dataType: find("Data Type", "DataType") >= 0 ? find("Data Type", "DataType") : 3,
    col1: h.findIndex(x => /列1|默认值/.test(x)) >= 0 ? h.findIndex(x => /列1|默认值/.test(x)) : 5,
    col2: h.findIndex(x => /列2/.test(x)) >= 0 ? h.findIndex(x => /列2/.test(x)) : 6,
    col3: h.findIndex(x => /列3/.test(x)) >= 0 ? h.findIndex(x => /列3/.test(x)) : 7,
    col4: h.findIndex(x => /列4/.test(x)) >= 0 ? h.findIndex(x => /列4/.test(x)) : 8,
    col5: h.findIndex(x => /列5/.test(x)) >= 0 ? h.findIndex(x => /列5/.test(x)) : 9
  };
}

function cell(row, colIdx) {
  if (colIdx < 0 || !row) return "";
  return (row[colIdx] ?? "").toString().trim();
}

const COVER_FIELD_MAP = [
  { header: "标题", id: "coverTitle" },
  { header: "简介", id: "coverSummary" },
  { header: "年份", id: "coverYearBadge" },
  { header: "年度", id: "coverYearLabel", prefix: "年度：" },
  { header: "主题", id: "coverThemeLabel", prefix: "主题：" },
  { header: "地市", id: "coverCityLabel", prefix: "地市：" },
  { header: "报告周期", id: "coverPeriodLabel", prefix: "报告周期：" },
  { header: "客流峰值", id: "coverFlowValue" },
  { header: "热点商圈", id: "coverBizValue" },
  { header: "网络覆盖", id: "coverNetValue" },
  { header: "城市场景", id: "coverSceneName" }
];

function buildRecordLegacy(values) {
  if (!values.length) return {};
  const headers = values[0].map(c => (c || "").trim());
  const row = values.slice(1).find(r => r.some(c => c));
  if (!row) return {};
  const record = {};
  headers.forEach((h, i) => { if (h) record[h] = (row[i] ?? "").toString().trim(); });
  return record;
}

function buildPage1DataLegacy(values) {
  if (!values.length) return null;
  const headers = values[0].map(c => (c || "").trim());
  const rows = values.slice(1);
  const col = (name) => { const i = headers.indexOf(name); return i === -1 ? [] : rows.map(r => (r[i] ?? "").toString().trim()); };
  const cell = (name, rowIndex = 0) => { const arr = col(name); return arr[rowIndex] ?? ""; };
  return {
    title: cell("标题"),
    summary: cell("简介"),
    total: { value: cell("总客流"), desc: (col("总客流")[1] ?? "").trim() },
    peak: { value: cell("峰值时段"), desc: (col("峰值时段")[1] ?? "").trim() },
    cross: { value: cell("跨省出行"), desc: (col("跨省出行")[1] ?? "").trim() },
    flow: { times: col("客流趋势（时间）").filter(Boolean), values: col("客流趋势（人数）").map(v => Number(v) || 0) },
    travel: { labels: col("出行方式对比（方式）").filter(Boolean), values: col("出行方式对比（人数）").map(v => Number(v) || 0) }
  };
}

function isRowEmpty(row) {
  return !row || row.every(c => !c || !String(c).trim());
}

function parsePage2DataLegacy(values) {
  const tables = [];
  let i = 0;
  while (i < values.length) {
    while (i < values.length && isRowEmpty(values[i])) i++;
    if (i >= values.length) break;
    const header = (values[i] || []).map(c => (c ?? "").toString().trim());
    i++;
    const rows = [];
    while (i < values.length && !isRowEmpty(values[i])) rows.push(values[i++].map(c => (c ?? "").toString().trim()));
    tables.push({ header, rows });
  }
  const find = (kw) => tables.find(t => (t.header?.[0] || "").includes(kw));
  const base = find("基础信息") ?? tables[0];
  const heatmap = (find("热力") ?? tables[1])?.rows?.map(r => ({ name: r?.[0] ?? "", value: Number(r?.[1]) || 0 })) ?? [];
  const topProvinces = (find("Top5") ?? tables[2])?.rows?.map(r => ({ name: r?.[1] ?? "", value: Number(r?.[2]) || 0 })) ?? [];
  const baseRecord = {};
  base?.rows?.forEach(r => { const k = r?.[0]?.trim(); if (k) baseRecord[k] = r?.[1] ?? ""; });
  const ring = { core: [], potential: [], position: [] };
  (find("圈层") ?? tables[3])?.rows?.forEach(r => {
    const type = r?.[0] ?? "", name = r?.[1] ?? "", value = Number(r?.[2]) || 0;
    if (!name) return;
    if (type.includes("核心")) ring.core.push({ name, value });
    else if (type.includes("潜力")) ring.potential.push({ name, value });
    else ring.position.push({ name, value });
  });
  return {
    title: baseRecord["页面标题"] ?? "",
    summary: baseRecord["摘要描述"] ?? "",
    highlights: [baseRecord["高亮 1"], baseRecord["高亮 2"], baseRecord["高亮 3"]].filter(Boolean),
    heatmap,
    topProvinces,
    ring
  };
}

function parseSheetTable(values, sheetKind = "cover") {
  const empty = { textUpdates: [], page1Data: null, page2Data: null, genericPageData: null };
  if (!values || values.length < 2) return empty;
  const headers = values[0].map(c => (c ?? "").toString().trim());
  const rows = values.slice(1);
  const hasNewFormat = headers.some(h => (h || "").includes("Field Key") || (h || "").includes("FieldKey"));
  if (!hasNewFormat) {
    if (sheetKind === "cover") {
      const record = buildRecordLegacy(values);
      const textUpdates = [];
      COVER_FIELD_MAP.forEach(({ header, id, prefix = "" }) => {
        let v = (record[header] ?? "").toString().trim();
        if (header === "年度" && !v && record["年份"]) v = record["年份"];
        if (header === "报告周期" && !v && record["报告期"]) v = record["报告期"];
        if (v) textUpdates.push({ id, value: prefix ? prefix + v : v });
      });
      return { ...empty, textUpdates };
    }
    if (sheetKind === "page1") {
      return { ...empty, page1Data: buildPage1DataLegacy(values) };
    }
    if (sheetKind === "page2") {
      return { ...empty, page2Data: parsePage2DataLegacy(values) };
    }
    return empty;
  }
  const idx = getColumnIndices(headers);
  const textUpdates = [];
  let page1Data = null;
  let page2Data = { title: "", summary: "", highlights: [], topProvinces: [], heatmap: [], ring: { core: [], potential: [], position: [] } };
  const isGenericPage = /^page([3-9]|1[0-3])$/.test(sheetKind);
  const genericPageData = isGenericPage ? { title: "", summary: "" } : null;
  const setGeneric = (key, value) => {
    if (genericPageData) genericPageData[key] = value;
  };
  const skipHeader = (v) => /^(值|name|label|默认表头|时间|人数|rank|activity|desc|tag|theme|stops)$/i.test((v || "").trim());

  let i = 0;
  while (i < rows.length) {
    const row = rows[i];
    const fieldKey = cell(row, idx.fieldKey);
    const htmlHook = cell(row, idx.htmlHook);
    const dataType = cell(row, idx.dataType);
    const v1 = cell(row, idx.col1);
    const v2 = cell(row, idx.col2);
    const v3 = cell(row, idx.col3);

    if (fieldKey) {
      const id = htmlHook.replace(/^#/, "").trim();
      if (dataType === "text" && id) {
        if (v1) textUpdates.push({ id, value: v1 });
      } else if (fieldKey === "flow.times+values" && (dataType || "").includes("chart")) {
        const times = [];
        const vals = [];
        i++;
        const skipChartHeader = (t) => /^(时间|人数|默认表头)$/i.test((t || "").trim());
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          const t = cell(rows[i], idx.col1);
          const v = cell(rows[i], idx.col2);
          if (t && !skipChartHeader(t)) {
            times.push(t);
            vals.push(Number(v) || 0);
          }
          i++;
        }
        if (!page1Data) page1Data = { title: "", summary: "", total: {}, peak: {}, cross: {}, flow: { times: [], values: [] }, travel: { labels: [], values: [] } };
        page1Data.flow = { times, values: vals };
        continue;
      } else if (fieldKey === "travel.labels+values" && (dataType || "").includes("chart")) {
        const labels = [];
        const vals = [];
        i++;
        const skipChartHeader = (x) => /^(出行方式|占比|默认表头|人数)$/i.test((x || "").trim());
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          const l = cell(rows[i], idx.col1);
          const v = cell(rows[i], idx.col2);
          if (l && !skipChartHeader(l)) {
            labels.push(l);
            vals.push(Number(v) || 0);
          }
          i++;
        }
        if (!page1Data) page1Data = { title: "", summary: "", total: {}, peak: {}, cross: {}, flow: { times: [], values: [] }, travel: { labels: [], values: [] } };
        page1Data.travel = { labels, values: vals };
        continue;
      } else if (fieldKey === "highlights" && (dataType || "").includes("list")) {
        const list = [];
        i++;
        const skipHeader = (v) => /^(值|name|label|默认表头)$/i.test((v || "").trim());
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          const a = cell(rows[i], idx.col1);
          const b = cell(rows[i], idx.col2);
          const c = cell(rows[i], idx.col3);
          if (a && !skipHeader(a)) list.push(a);
          if (b && !skipHeader(b)) list.push(b);
          if (c && !skipHeader(c)) list.push(c);
          i++;
        }
        page2Data.highlights = list;
        continue;
      } else if (fieldKey === "topProvinces" && (dataType || "").includes("list")) {
        i++;
        const list = [];
        const isHeaderRow = (r) => /^(name|value|默认表头|省份|占比)$/i.test((cell(r, idx.col1) || "").trim());
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          if (isHeaderRow(rows[i])) { i++; continue; }
          const name = cell(rows[i], idx.col1);
          const value = cell(rows[i], idx.col2);
          if (name) list.push({ name, value: Number(value) || 0 });
          i++;
        }
        page2Data.topProvinces = list;
        continue;
      } else if (fieldKey === "heatmap" && (dataType || "").includes("chart")) {
        i++;
        const list = [];
        const isHeaderRow = (r) => /^(name|value|默认表头)$/i.test((cell(r, idx.col1) || "").trim());
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          if (isHeaderRow(rows[i])) { i++; continue; }
          const name = cell(rows[i], idx.col1);
          const value = cell(rows[i], idx.col2);
          if (name) list.push({ name, value: Number(value) || 0 });
          i++;
        }
        page2Data.heatmap = list;
        continue;
      } else if (fieldKey === "flow" && (dataType || "").includes("chart") && isGenericPage) {
        const times = [];
        const vals = [];
        i++;
        const skipChartHeader = (t) => /^(时间|人数|默认表头)$/i.test((t || "").trim());
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          const t = cell(rows[i], idx.col1);
          const v = cell(rows[i], idx.col2);
          if (t && !skipChartHeader(t)) {
            times.push(t);
            vals.push(Number(v) || 0);
          }
          i++;
        }
        setGeneric("flow", { times, values: vals });
        continue;
      } else if ((dataType || "").includes("list(label,value,desc)") && (fieldKey === "cards" || fieldKey === "extras") && isGenericPage) {
        i++;
        const list = [];
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          const label = cell(rows[i], idx.col1);
          const value = cell(rows[i], idx.col2);
          const desc = cell(rows[i], idx.col3);
          if (label && !skipHeader(label)) list.push({ label, value, desc });
          i++;
        }
        setGeneric(fieldKey, list);
        continue;
      } else if (fieldKey === "trend.categories+series" && (dataType || "").includes("chart") && isGenericPage) {
        i++;
        const categories = [];
        const series = [];
        const headerRow = rows[i];
        if (headerRow) {
          const names = [];
          for (let c = idx.col1 + 1; c <= idx.col5; c++) {
            const n = cell(headerRow, c);
            if (n) names.push(n);
          }
          names.forEach(name => series.push({ name, data: [] }));
          i++;
        }
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          const cat = cell(rows[i], idx.col1);
          if (cat && !skipHeader(cat)) {
            categories.push(cat);
            for (let si = 0; si < series.length; si++) {
              const v = cell(rows[i], idx.col2 + si);
              series[si].data.push(Number(v) || 0);
            }
          }
          i++;
        }
        setGeneric("trend", { categories, series });
        continue;
      } else if (fieldKey === "ranking" && (dataType || "").includes("table") && isGenericPage) {
        i++;
        const list = [];
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          const rank = cell(rows[i], idx.col1);
          const name = cell(rows[i], idx.col2);
          const activity = cell(rows[i], idx.col3);
          const value = cell(rows[i], idx.col4);
          if ((rank || name) && !skipHeader(rank)) list.push({ rank, name, activity, value });
          i++;
        }
        setGeneric("ranking", list);
        continue;
      } else if ((dataType || "").includes("list(label,value,tag)") && fieldKey === "preferences" && isGenericPage) {
        i++;
        const list = [];
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          const label = cell(rows[i], idx.col1);
          const value = cell(rows[i], idx.col2);
          const tag = cell(rows[i], idx.col3);
          if (label && !skipHeader(label)) list.push({ label, value, tag });
          i++;
        }
        setGeneric("preferences", list);
        continue;
      } else if ((dataType || "").includes("list(") && fieldKey === "routes" && isGenericPage) {
        i++;
        const list = [];
        while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
          const name = cell(rows[i], idx.col1);
          const theme = cell(rows[i], idx.col2);
          const stops = cell(rows[i], idx.col3) || cell(rows[i], idx.col4) || "";
          if (name && !skipHeader(name)) list.push({ name, theme, stops: stops ? [stops] : [] });
          i++;
        }
        setGeneric("routes", list);
        continue;
      } else if (fieldKey === "highlights" && (dataType || "").includes("list") && isGenericPage) {
        const isLabelValue = (dataType || "").includes("label") && (dataType || "").includes("value");
        const list = [];
        i++;
        if (isLabelValue) {
          const skipHeaderRow = (r) => /^(label|value|desc|默认表头)$/i.test((cell(r, idx.col1) || "").trim());
          while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
            if (skipHeaderRow(rows[i])) { i++; continue; }
            const label = cell(rows[i], idx.col1);
            const value = cell(rows[i], idx.col2);
            const desc = cell(rows[i], idx.col3);
            if (value || label) list.push({ label, value, desc });
            i++;
          }
        } else {
          while (i < rows.length && !cell(rows[i], idx.fieldKey)) {
            const a = cell(rows[i], idx.col1);
            const b = cell(rows[i], idx.col2);
            const c = cell(rows[i], idx.col3);
            if (a && !skipHeader(a)) list.push(a);
            if (b && !skipHeader(b)) list.push(b);
            if (c && !skipHeader(c)) list.push(c);
            i++;
          }
        }
        setGeneric("highlights", list);
        continue;
      } else if (fieldKey === "title" && id) {
        textUpdates.push({ id, value: v1 });
        if (id.startsWith("page1") && !page1Data) page1Data = { title: "", summary: "", total: {}, peak: {}, cross: {}, flow: { times: [], values: [] }, travel: { labels: [], values: [] } };
        if (page1Data && id === "page1Title") page1Data.title = v1;
        if (page2Data && id === "page2Title") page2Data.title = v1;
        if (genericPageData && /^page([3-9]|1[0-3])Title$/.test(id)) genericPageData.title = v1;
      } else if (fieldKey === "summary" && id) {
        textUpdates.push({ id, value: v1 });
        if (page1Data && id === "page1Summary") page1Data.summary = v1;
        if (page2Data && id === "page2Summary") page2Data.summary = v1;
        if (genericPageData && /^page([3-9]|1[0-3])Summary$/.test(id)) genericPageData.summary = v1;
      } else if (fieldKey === "total.value" && id) {
        textUpdates.push({ id, value: v1 });
        if (!page1Data) page1Data = { title: "", summary: "", total: {}, peak: {}, cross: {}, flow: { times: [], values: [] }, travel: { labels: [], values: [] } };
        page1Data.total = page1Data.total || {}; page1Data.total.value = v1;
      } else if (fieldKey === "total.desc" && id) {
        textUpdates.push({ id, value: v1 });
        if (!page1Data) page1Data = { title: "", summary: "", total: {}, peak: {}, cross: {}, flow: { times: [], values: [] }, travel: { labels: [], values: [] } };
        page1Data.total = page1Data.total || {}; page1Data.total.desc = v1;
      } else if (fieldKey === "peak.value" && id) {
        textUpdates.push({ id, value: v1 });
        if (!page1Data) page1Data = { title: "", summary: "", total: {}, peak: {}, cross: {}, flow: { times: [], values: [] }, travel: { labels: [], values: [] } };
        page1Data.peak = page1Data.peak || {}; page1Data.peak.value = v1;
      } else if (fieldKey === "peak.desc" && id) {
        textUpdates.push({ id, value: v1 });
        if (!page1Data) page1Data = { title: "", summary: "", total: {}, peak: {}, cross: {}, flow: { times: [], values: [] }, travel: { labels: [], values: [] } };
        page1Data.peak = page1Data.peak || {}; page1Data.peak.desc = v1;
      } else if (fieldKey === "cross.value" && id) {
        textUpdates.push({ id, value: v1 });
        if (!page1Data) page1Data = { title: "", summary: "", total: {}, peak: {}, cross: {}, flow: { times: [], values: [] }, travel: { labels: [], values: [] } };
        page1Data.cross = page1Data.cross || {}; page1Data.cross.value = v1;
      } else if (fieldKey === "cross.desc" && id) {
        textUpdates.push({ id, value: v1 });
        if (!page1Data) page1Data = { title: "", summary: "", total: {}, peak: {}, cross: {}, flow: { times: [], values: [] }, travel: { labels: [], values: [] } };
        page1Data.cross = page1Data.cross || {}; page1Data.cross.desc = v1;
      } else if (id && v1 && (dataType === "text" || !dataType)) {
        textUpdates.push({ id, value: v1 });
      }
    }
    i++;
  }

  return { textUpdates, page1Data, page2Data, genericPageData };
}

function replaceContentById(html, id, value) {
  if (value == null || value === "") return html;
  const pattern = new RegExp(`(<[^>]*id="${id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*>)([\\s\\S]*?)(</[^>]+>)`);
  if (!pattern.test(html)) {
    console.warn(`未找到 id="${id}" 的元素`);
    return html;
  }
  const safe = value
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return html.replace(pattern, `$1${safe}$3`);
}

function applyPage1Constant(html, data) {
  if (!data) return html;
  const pattern = /const PAGE1_DATA = \/\* PAGE1_DATA_START \*\/[\s\S]*?\/\* PAGE1_DATA_END \*\/;/;
  if (!pattern.test(html)) {
    console.warn("未找到 PAGE1_DATA 常量，跳过页面 1 数据更新");
    return html;
  }
  const serialized = JSON.stringify(data, null, 2);
  return html.replace(pattern, `const PAGE1_DATA = /* PAGE1_DATA_START */ ${serialized} /* PAGE1_DATA_END */;`);
}

async function loadPageData() {
  try {
    const raw = await fs.readFile(PAGE_DATA_PATH, "utf8");
    const match = raw.match(/window\.PAGES_DATA\s*=\s*(\{[\s\S]*\});?/);
    if (!match) throw new Error("PARSE_ERROR");
    return JSON.parse(match[1]);
  } catch (e) {
    if (e.code === "ENOENT" || e.message === "PARSE_ERROR") return {};
    throw e;
  }
}

async function writePageData(data) {
  const payload = `window.PAGES_DATA = ${JSON.stringify(data, null, 2)};
`;
  await fs.writeFile(PAGE_DATA_PATH, payload, "utf8");
}

async function main() {
  try {
    const coverValues = await fetchRange(COVER_RANGE);
    const page1Values = await fetchRange(PAGE1_RANGE);
    const page2Range = await resolveSheetRange({
      sheetTitle: PAGE2_SHEET_NAME,
      sheetId: PAGE2_SHEET_ID,
      cellRange: PAGE2_RANGE
    });
    const page2Values = await fetchRange(page2Range);

    const coverParsed = parseSheetTable(coverValues, "cover");
    const page1Parsed = parseSheetTable(page1Values, "page1");
    const page2Parsed = parseSheetTable(page2Values, "page2");

    const pageParsed = [];
    for (let n = 3; n <= 13; n++) {
      const range = await resolvePageRange(n);
      if (!range) continue;
      try {
        const values = await fetchRange(range);
        const parsed = parseSheetTable(values, `page${n}`);
        pageParsed.push({ n, parsed });
      } catch (e) {
        console.warn(`页面 ${n} 拉取或解析跳过：${e.message}`);
      }
    }

    let html = await fs.readFile(HTML_PATH, "utf8");

    for (const { id, value } of coverParsed.textUpdates) {
      html = replaceContentById(html, id, value);
    }
    for (const { id, value } of page1Parsed.textUpdates) {
      html = replaceContentById(html, id, value);
    }
    for (const { id, value } of page2Parsed.textUpdates) {
      html = replaceContentById(html, id, value);
    }
    for (const { parsed } of pageParsed) {
      for (const { id, value } of parsed.textUpdates) {
        html = replaceContentById(html, id, value);
      }
    }

    const page1Data = page1Parsed.page1Data || {
      title: "",
      summary: "",
      total: { value: "", desc: "" },
      peak: { value: "", desc: "" },
      cross: { value: "", desc: "" },
      flow: { times: [], values: [] },
      travel: { labels: [], values: [] }
    };
    html = applyPage1Constant(html, page1Data);

    const pageData = await loadPageData();
    pageData.page2 = {
      title: page2Parsed.page2Data?.title ?? pageData.page2?.title ?? "",
      summary: page2Parsed.page2Data?.summary ?? pageData.page2?.summary ?? "",
      highlights: (page2Parsed.page2Data?.highlights?.length ? page2Parsed.page2Data.highlights : pageData.page2?.highlights) ?? [],
      topProvinces: (page2Parsed.page2Data?.topProvinces?.length ? page2Parsed.page2Data.topProvinces : pageData.page2?.topProvinces) ?? [],
      heatmap: (page2Parsed.page2Data?.heatmap?.length ? page2Parsed.page2Data.heatmap : pageData.page2?.heatmap) ?? [],
      ring: pageData.page2?.ring ?? { core: [], potential: [], position: [] }
    };
    for (const { n, parsed } of pageParsed) {
      if (parsed.genericPageData) {
        pageData[`page${n}`] = { ...(pageData[`page${n}`] || {}), ...parsed.genericPageData };
      }
    }
    await writePageData(pageData);

    await fs.writeFile(HTML_PATH, html, "utf8");
    const pagesDone = 2 + pageParsed.length;
    console.log(`封面、页面 1～${pagesDone} 已从飞书表格同步到 HTML 与 page-data.js。`);
  } catch (error) {
    console.error(error.message || error);
    process.exit(1);
  }
}

main();
