import fs from "node:fs/promises";
import { feishuFetch } from "./feishu_auth.mjs";

const {
  FEISHU_SHEET_TOKEN = "KcGusWwSuhBMcot3Deucjd9Inix",
  FEISHU_PAGE1_RANGE = "1!A1:Z200",
  HTML_TARGET = "Maanshan3DMap/PPT1.html"
} = process.env;

async function fetchSheet(range) {
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

function buildPage1Data(values) {
  if (!values.length) throw new Error("表格中没有数据");
  const headers = values[0].map(cell => (cell || "").trim());
  const rows = values.slice(1);

  const getColumnIndex = header => headers.indexOf(header);
  const getCell = (header, rowIndex = 0) => {
    const idx = getColumnIndex(header);
    if (idx === -1) return "";
    return (rows[rowIndex]?.[idx] ?? "").toString().trim();
  };
  const getColumn = (header, startRow = 0) => {
    const idx = getColumnIndex(header);
    if (idx === -1) return [];
    return rows.slice(startRow).map(row => (row[idx] ?? "").toString().trim()).filter(Boolean);
  };

  return {
    title: getCell("标题"),
    summary: getCell("简介"),
    total: {
      value: getCell("总客流"),
      desc: getCell("总客流", 1)
    },
    peak: {
      value: getCell("峰值时段"),
      desc: getCell("峰值时段", 1)
    },
    cross: {
      value: getCell("跨省出行"),
      desc: getCell("跨省出行", 1)
    },
    flow: {
      times: getColumn("客流趋势（时间）"),
      values: getColumn("客流趋势（人数）")
    },
    travel: {
      labels: getColumn("出行方式对比（方式）"),
      values: getColumn("出行方式对比（人数）")
    }
  };
}

async function updatePage1Data(data) {
  const html = await fs.readFile(HTML_TARGET, "utf8");
  const pattern =
    /const PAGE1_DATA = \/\\* PAGE1_DATA_START \\*\/[\s\S]*?\/\\* PAGE1_DATA_END \\*\/;/;
  const replacement = `const PAGE1_DATA = /* PAGE1_DATA_START */ ${JSON.stringify(
    data,
    null,
    2
  )} /* PAGE1_DATA_END */;`;

  if (!pattern.test(html)) {
    throw new Error("未找到 PAGE1_DATA 常量块");
  }

  const updated = html.replace(pattern, replacement);
  await fs.writeFile(HTML_TARGET, updated, "utf8");
  console.log("页面 1 数据已同步飞书表格。");
}

async function main() {
  try {
    const values = await fetchSheet(FEISHU_PAGE1_RANGE);
    const page1Data = buildPage1Data(values);
    await updatePage1Data(page1Data);
  } catch (error) {
    console.error(error.message || error);
    process.exit(1);
  }
}

main();
