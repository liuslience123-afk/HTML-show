import fs from "node:fs/promises";
import { resolveTenantToken } from "./feishu_auth.mjs";

const {
  FEISHU_SHEET_TOKEN,
  FEISHU_RANGE = "Sheet1!A1:Z50",
  FEISHU_OUTPUT = "",
  FEISHU_LIST_SHEETS,
  FEISHU_DIRECT_TOKEN
} = process.env;

if (!FEISHU_SHEET_TOKEN) {
  console.error("缺少环境变量：FEISHU_SHEET_TOKEN");
  process.exit(1);
}

async function getSheetRange(tenantToken, sheetToken, range) {
  const url = new URL(
    `https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/${sheetToken}/values_batch_get`
  );
  url.searchParams.append("ranges", range);

  const resp = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${tenantToken}`,
      "Content-Type": "application/json"
    }
  });
  const data = await resp.json();
  if (data.code !== 0) {
    throw new Error(`读取表格失败：${data.code} - ${data.msg}`);
  }
  return data.data?.valueRanges ?? [];
}

async function getSheetMeta(tenantToken, sheetToken) {
  const resp = await fetch(
    `https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/${sheetToken}`,
    {
      method: "GET",
      headers: {
        Authorization: `Bearer ${tenantToken}`,
        "Content-Type": "application/json"
      }
    }
  );
  const data = await resp.json();
  if (data.code !== 0) {
    throw new Error(`获取表格信息失败：${data.code} - ${data.msg}`);
  }
  return data.data?.sheets ?? [];
}

async function main() {
  try {
    const tenantToken = await resolveTenantToken();

    if (FEISHU_DIRECT_TOKEN) {
      const directData = await getDirectRange(
        FEISHU_DIRECT_TOKEN,
        FEISHU_SHEET_TOKEN,
        FEISHU_RANGE
      );
      const payload = JSON.stringify(directData, null, 2);
      if (FEISHU_OUTPUT) {
        await fs.writeFile(FEISHU_OUTPUT, payload, "utf8");
        console.log(`数据已写入 ${FEISHU_OUTPUT}`);
      } else {
        console.log(payload);
      }
      return;
    }

    if (FEISHU_LIST_SHEETS) {
      const sheets = await getSheetMeta(tenantToken, FEISHU_SHEET_TOKEN);
      console.log("可用 sheet 列表：");
      sheets.forEach(sheet => {
        console.log(
          `- title: ${sheet?.title ?? "未知"}\n  sheet_id: ${sheet?.sheet_id}\n  grid_id: ${sheet?.grid_id}`
        );
      });
      return;
    }

    const ranges = await getSheetRange(tenantToken, FEISHU_SHEET_TOKEN, FEISHU_RANGE);

    if (FEISHU_OUTPUT) {
      await fs.writeFile(
        FEISHU_OUTPUT,
        JSON.stringify(ranges, null, 2),
        "utf8"
      );
      console.log(`已写入 ${FEISHU_OUTPUT}`);
    } else {
      console.log(JSON.stringify(ranges, null, 2));
    }
  } catch (error) {
    console.error(error);
    process.exit(1);
  }
}

async function getDirectRange(token, spreadsheetToken, range) {
  const encodedRange = encodeURIComponent(range);
  const url = `https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${encodedRange}`;
  const resp = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    }
  });
  const data = await resp.json();
  if (data.code !== 0) {
    throw new Error(`读取表格失败：${data.code} - ${data.msg}`);
  }
  return data.data;
}

main();
