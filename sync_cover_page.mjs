import fs from "node:fs/promises";
import { feishuFetch } from "./feishu_auth.mjs";

const {
  FEISHU_SHEET_TOKEN = "KcGusWwSuhBMcot3Deucjd9Inix",
  FEISHU_RANGE = "bfd5ff!A1:J20",
  HTML_TARGET = "Maanshan3DMap/PPT1.html"
} = process.env;

const FIELD_MAP = {
  标题: { id: "coverTitle" },
  简介: { id: "coverSummary" },
  年份: { id: "coverYearBadge" },
  年度: { id: "coverYearLabel", prefix: "年度：" },
  主题: { id: "coverThemeLabel", prefix: "主题：" },
  地市: { id: "coverCityLabel", prefix: "地市：" },
  报告周期: { id: "coverPeriodLabel", prefix: "报告周期：" },
  客流峰值: { id: "coverFlowValue" },
  热点商圈: { id: "coverBizValue" },
  网络覆盖: { id: "coverNetValue" },
  城市场景: { id: "coverSceneName" }
};

async function fetchSheetData() {
  const url = `https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/${FEISHU_SHEET_TOKEN}/values/${encodeURIComponent(
    FEISHU_RANGE
  )}`;
  const resp = await feishuFetch(url);
  const json = await resp.json();
  if (json.code !== 0) {
    throw new Error(`飞书 API 返回异常：${json.code} - ${json.msg}`);
  }
  const values = json.data?.valueRange?.values || [];
  if (values.length < 2) throw new Error("表格中没有可用的数据行");
  const headers = values[0].map(cell => (cell || "").trim());
  const row = values.slice(1).find(row => row.some(cell => cell));
  if (!row) throw new Error("未找到包含数据的行");

  const record = {};
  headers.forEach((header, idx) => {
    if (!header) return;
    record[header] = row[idx] ?? "";
  });
  return record;
}

function replaceContentById(html, id, text) {
  const pattern = new RegExp(`(<[^>]*id="${id}"[^>]*>)([\\s\\S]*?)(</[^>]+>)`);
  if (!pattern.test(html)) {
    console.warn(`警告：未找到 id="${id}" 的元素`);
    return html;
  }
  const safeText = text
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return html.replace(pattern, `$1${safeText}$3`);
}

async function updateCover(record) {
  let html = await fs.readFile(HTML_TARGET, "utf8");

  Object.entries(FIELD_MAP).forEach(([field, { id, prefix = "" }]) => {
    const value = (record[field] ?? "").toString().trim();
    if (!value) return;
    html = replaceContentById(html, id, prefix ? `${prefix}${value}` : value);
  });

  // 如果表格提供了“年份”，同步到年度标签
  if (record["年份"] && !record["年度"]) {
    html = replaceContentById(html, "coverYearLabel", `年度：${record["年份"]}`);
  }

  await fs.writeFile(HTML_TARGET, html, "utf8");
  console.log("封面数据已同步飞书表格。");
}

async function main() {
  try {
    const record = await fetchSheetData();
    await updateCover(record);
  } catch (err) {
    console.error(err.message || err);
    process.exit(1);
  }
}

main();
