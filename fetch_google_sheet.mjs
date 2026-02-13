import fs from "node:fs/promises";

const {
  GOOGLE_SHEET_ID,
  GOOGLE_SHEET_NAME = "Sheet1",
  GOOGLE_SHEET_QUERY = "",
  GOOGLE_OUTPUT = ""
} = process.env;

if (!GOOGLE_SHEET_ID) {
  console.error("缺少环境变量：GOOGLE_SHEET_ID");
  process.exit(1);
}

const BASE_URL = "https://docs.google.com/spreadsheets/d";

function buildUrl(sheetId, sheetName, query) {
  const url = new URL(`${BASE_URL}/${sheetId}/gviz/tq`);
  url.searchParams.set("sheet", sheetName);
  if (query) {
    url.searchParams.set("tq", query);
  }
  return url.toString();
}

function parseGvizResponse(text) {
  const prefix = "google.visualization.Query.setResponse(";
  const suffix = ");";
  const start = text.indexOf(prefix);
  const end = text.lastIndexOf(suffix);
  if (start === -1 || end === -1) {
    throw new Error("无法解析返回内容：未找到 gviz 包裹结构");
  }
  const jsonText = text.slice(start + prefix.length, end + 1);
  const data = JSON.parse(jsonText);
  if (data.status !== "ok") {
    throw new Error(`Google Sheets 返回错误：${data.status}`);
  }
  return data.table;
}

function tableToObjects(table) {
  const cols = table.cols.map(col => col.label || col.id);
  const rows = table.rows.map(row => {
    const obj = {};
    row.c?.forEach((cell, idx) => {
      const key = cols[idx] || `col_${idx}`;
      obj[key] = cell ? cell.v : null;
    });
    return obj;
  });
  return { columns: cols, rows };
}

async function main() {
  try {
    const url = buildUrl(GOOGLE_SHEET_ID, GOOGLE_SHEET_NAME, GOOGLE_SHEET_QUERY);
    const resp = await fetch(url);
    if (!resp.ok) {
      throw new Error(`请求失败：${resp.status} ${resp.statusText}`);
    }
    const text = await resp.text();
    const table = parseGvizResponse(text);
    const result = tableToObjects(table);
    const output = JSON.stringify(result, null, 2);

    if (GOOGLE_OUTPUT) {
      await fs.writeFile(GOOGLE_OUTPUT, output, "utf8");
      console.log(`数据已写入 ${GOOGLE_OUTPUT}`);
    } else {
      console.log(output);
    }
  } catch (error) {
    console.error(error);
    process.exit(1);
  }
}

main();
