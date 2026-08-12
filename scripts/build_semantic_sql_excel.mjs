import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");
const inputPath = path.join(rootDir, "outputs", "semantic_sql_compare_results.json");
const outputDir = path.join(rootDir, "outputs", "semantic_sql_compare_workbook");
const outputPath = path.join(outputDir, "0709_SQL_7月10日_7月13日_AI人工语义匹配完整版.xlsx");

const raw = JSON.parse(await fs.readFile(inputPath, "utf8"));
const detailRows = raw.details.map((row) => ({
  "0709序号": row["0709序号"],
  "0709原表行号": row["0709原表行号"],
  "任务编号": row["任务编号"],
  "应用场景": row["应用场景"],
  "代码位置": row["代码位置"],
  "原修复情况": row["修复情况"],
  "0709 SQL": row["0709 SQL"],
  "7月10日判断": row["7月10日判断"],
  "7月10日置信度": row["7月10日置信度"],
  "7月10日匹配序号": row["7月10日匹配序号"],
  "7月10日原表行号": row["7月10日原表行号"],
  "7月10日指纹": row["7月10日指纹"],
  "7月10日服务": row["7月10日服务"],
  "7月10日候选SQL": row["7月10日候选SQL"],
  "7月10日详细分析": `0709行号：${row["0709行号标注"]}\n7月10行号：${row["7月10日行号标注"]}\n${row["7月10日详细分析"]}`,
  "7月13日判断": row["7月13日判断"],
  "7月13日置信度": row["7月13日置信度"],
  "7月13日匹配序号": row["7月13日匹配序号"],
  "7月13日原表行号": row["7月13日原表行号"],
  "7月13日指纹": row["7月13日指纹"],
  "7月13日服务": row["7月13日服务"],
  "7月13日候选SQL": row["7月13日候选SQL"],
  "7月13日详细分析": `0709行号：${row["0709行号标注"]}\n7月13行号：${row["7月13日行号标注"]}\n${row["7月13日详细分析"]}`,
  "综合结论": row["综合结论"],
}));

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function toMatrix(rows, columns) {
  return rows.map((row) => columns.map((col) => row[col] ?? ""));
}

function setHeader(sheet, columns, title, note) {
  const end = colLetter(columns.length);
  sheet.getRange(`A1:${end}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${end}1`).format.fill = { color: "#1F4E78" };
  sheet.getRange(`A1:${end}1`).format.font = { color: "#FFFFFF", bold: true, size: 14 };
  sheet.getRange(`A1:${end}1`).format.rowHeight = 24;

  sheet.getRange(`A2:${end}2`).merge();
  sheet.getRange("A2").values = [[note]];
  sheet.getRange(`A2:${end}2`).format.fill = { color: "#D9EAF7" };
  sheet.getRange(`A2:${end}2`).format.font = { color: "#244061", italic: true };
  sheet.getRange(`A2:${end}2`).format.wrapText = true;

  sheet.getRange(`A3:${end}3`).values = [columns];
  sheet.getRange(`A3:${end}3`).format.fill = { color: "#5B9BD5" };
  sheet.getRange(`A3:${end}3`).format.font = { color: "#FFFFFF", bold: true };
  sheet.getRange(`A3:${end}3`).format.wrapText = true;
  sheet.getRange(`A3:${end}3`).format.borders = { preset: "all", style: "thin", color: "#D9D9D9" };
  sheet.freezePanes.freezeRows(3);
  sheet.showGridLines = false;
}

function writeTable(sheet, rows, columns) {
  if (!rows.length) return;
  const startRow = 4;
  const endRow = startRow + rows.length - 1;
  const endCol = colLetter(columns.length);
  sheet.getRange(`A${startRow}:${endCol}${endRow}`).values = toMatrix(rows, columns);
  sheet.getRange(`A${startRow}:${endCol}${endRow}`).format.wrapText = true;
  sheet.getRange(`A${startRow}:${endCol}${endRow}`).format.verticalAlignment = "top";
  sheet.getRange(`A${startRow}:${endCol}${endRow}`).format.borders = { preset: "all", style: "thin", color: "#E6E6E6" };
}

function setWidths(sheet, widths) {
  widths.forEach((width, idx) => {
    const col = colLetter(idx + 1);
    sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  });
}

function addDecisionColors(sheet, colIndex, rowCount) {
  const col = colLetter(colIndex);
  const range = sheet.getRange(`${col}4:${col}${rowCount + 3}`);
  range.conditionalFormats.add("containsText", {
    text: "同一业务SQL",
    format: {
      fill: { color: "#E2F0D9" },
      font: { color: "#215E21", bold: true },
    },
  });
  range.conditionalFormats.add("containsText", {
    text: "可能同一业务SQL",
    format: {
      fill: { color: "#FFF2CC" },
      font: { color: "#7F6000", bold: true },
    },
  });
  range.conditionalFormats.add("containsText", {
    text: "非同一业务SQL",
    format: {
      fill: { color: "#FCE4D6" },
      font: { color: "#9E480E", bold: true },
    },
  });
}

const workbook = Workbook.create();

const summarySheet = workbook.worksheets.add("汇总");
const summaryCols = ["统计项", "数量"];
setHeader(
  summarySheet,
  summaryCols,
  "0709 SQL 与 7月10 / 7月13 全量 SQL AI 语义比对汇总",
  "说明：本工作簿按 SQL 语义结构做比对，重点关注核心表、JOIN、WHERE、GROUP BY、子查询、UNION 与过滤逻辑。行号基于本次导出的格式化 SQL。"
);
writeTable(summarySheet, raw.summary, summaryCols);
setWidths(summarySheet, [34, 20]);

const detailSheet = workbook.worksheets.add("逐条语义比对");
const detailCols = [
  "0709序号",
  "0709原表行号",
  "任务编号",
  "应用场景",
  "代码位置",
  "原修复情况",
  "0709 SQL",
  "7月10日判断",
  "7月10日置信度",
  "7月10日匹配序号",
  "7月10日原表行号",
  "7月10日指纹",
  "7月10日服务",
  "7月10日候选SQL",
  "7月10日详细分析",
  "7月13日判断",
  "7月13日置信度",
  "7月13日匹配序号",
  "7月13日原表行号",
  "7月13日指纹",
  "7月13日服务",
  "7月13日候选SQL",
  "7月13日详细分析",
  "综合结论",
];
setHeader(
  detailSheet,
  detailCols,
  "逐条语义比对明细",
  "每条 0709 SQL 都与 7月10 / 7月13 全量 SQL 做了候选打分，明细页保留最佳匹配与解释。"
);
writeTable(detailSheet, detailRows, detailCols);
setWidths(detailSheet, [8, 10, 26, 28, 34, 14, 48, 13, 10, 9, 10, 18, 16, 48, 60, 13, 10, 9, 10, 18, 16, 48, 60, 18]);
addDecisionColors(detailSheet, 8, detailRows.length);
addDecisionColors(detailSheet, 16, detailRows.length);

await fs.mkdir(outputDir, { recursive: true });
const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);

const summaryInspect = await workbook.inspect({
  kind: "table",
  sheetId: "汇总",
  range: "A1:B12",
  include: "values",
  tableMaxRows: 12,
  tableMaxCols: 4,
});
console.log(summaryInspect.ndjson);

const detailInspect = await workbook.inspect({
  kind: "table",
  sheetId: "逐条语义比对",
  range: "A1:H8",
  include: "values",
  tableMaxRows: 8,
  tableMaxCols: 8,
});
console.log(detailInspect.ndjson);

const renderTargets = [
  ["汇总", "A1:B12"],
  ["逐条语义比对", "A1:H8"],
];

for (const [sheetName, range] of renderTargets) {
  const blob = await workbook.render({ sheetName, range, scale: 1.2, format: "png", autoCrop: "all" });
  const bytes = new Uint8Array(await blob.arrayBuffer());
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), bytes);
}

console.log(outputPath);
