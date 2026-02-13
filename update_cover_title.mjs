import fs from "node:fs/promises";
import { feishuFetch } from "./feishu_auth.mjs";

const {
  FEISHU_SHEET_TOKEN = "KcGusWwSuhBMcot3Deucjd9Inix",
  FEISHU_TITLE_RANGE = "bfd5ff!A1:D20",
  HTML_TARGET = "Maanshan3DMap/PPT1.html"
} = process.env;

async function fetchSheetTitle() {
  const url = `https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/${FEISHU_SHEET_TOKEN}/values/${encodeURIComponent(FEISHU_TITLE_RANGE)}`;
  const resp = await feishuFetch(url);
  const json = await resp.json();
  if (json.code !== 0) {
    throw new Error(`飞书 API 调用失败：${json.code} - ${json.msg}`);
  }
  const values = json.data?.valueRange?.values || [];
  if (values.length < 2) {
    throw new Error("未找到任何数据行");
  }
  const headers = values[0].map(cell => (cell || "").trim());
  const titleIndex = headers.findIndex(name => name.includes("标题"));
  if (titleIndex === -1) {
    throw new Error("表头中未找到“标题”列");
  }
  const row = values.find((row, idx) => idx > 0 && row[titleIndex]);
  if (!row) {
    throw new Error("未找到包含标题数据的行");
  }
  return row[titleIndex].toString();
}

async function updateCoverTitle(newTitle) {
  const html = await fs.readFile(HTML_TARGET, "utf8");
  const pattern = /(<h1[^>]*id="coverTitle"[^>]*>)([\s\S]*?)(<\/h1>)/;
  if (!pattern.test(html)) {
    throw new Error("无法在 HTML 中找到 coverTitle 元素");
  }
  const updated = html.replace(pattern, `$1${newTitle}$3`);
  await fs.writeFile(HTML_TARGET, updated, "utf8");
  console.log(`已将封面标题替换为：${newTitle}`);
}

async function main() {
  try {
    const title = await fetchSheetTitle();
    await updateCoverTitle(title);
  } catch (err) {
    console.error(err.message || err);
    process.exit(1);
  }
}

main();
