import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const DETAIL_COLUMNS = [
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

const DETAIL_WIDTHS = [8, 10, 24, 28, 32, 14, 42, 12, 10, 10, 10, 18, 14, 42, 52, 12, 10, 10, 10, 18, 14, 42, 52, 18];
const SUMMARY_WIDTHS = [28, 14];

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function matrix(rows, columns) {
  return rows.map((row) => columns.map((column) => row[column] ?? ""));
}

function setColumnWidths(sheet, widths) {
  widths.forEach((width, i) => {
    const col = colLetter(i + 1);
    sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  });
}

function paintHeader(range) {
  range.format.fill = { color: "#D9E2F3" };
  range.format.font = { bold: true, color: "#1F1F1F" };
  range.format.borders = { preset: "all", style: "thin", color: "#C9D2E3" };
  range.format.wrapText = true;
}

function paintBody(range) {
  range.format.borders = { preset: "all", style: "thin", color: "#E6E6E6" };
  range.format.wrapText = true;
  range.format.verticalAlignment = "top";
}

function addDecisionColors(sheet, colIndex, rowStart, rowEnd) {
  const col = colLetter(colIndex);
  const range = sheet.getRange(`${col}${rowStart}:${col}${rowEnd}`);
  range.conditionalFormats.add("containsText", {
    text: "同一业务SQL",
    format: { fill: { color: "#E2F0D9" }, font: { color: "#215E21", bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "未匹配",
    format: { fill: { color: "#FCE4D6" }, font: { color: "#9E480E", bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "可能同一业务SQL",
    format: { fill: { color: "#FFF2CC" }, font: { color: "#7F6000", bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "高",
    format: { fill: { color: "#E2F0D9" }, font: { color: "#215E21", bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "极高",
    format: { fill: { color: "#C6E0B4" }, font: { color: "#215E21", bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "低",
    format: { fill: { color: "#FCE4D6" }, font: { color: "#9E480E", bold: true } },
  });
}

export async function generateAiMatchWorkbook({
  inputPath,
  outputPath,
  summaryTitle = "0709 SQL语义匹配汇总",
}) {
  const raw = JSON.parse(await fs.readFile(inputPath, "utf8"));
  const summaryRows = Array.isArray(raw.summary_rows) ? raw.summary_rows : [];
  const detailRows = Array.isArray(raw.detail_rows) ? raw.detail_rows : [];

  const workbook = Workbook.create();

  const summary = workbook.worksheets.add("汇总");
  summary.showGridLines = false;
  summary.getRange("A1").values = [[summaryTitle]];
  summary.getRange("A1").format.font = { bold: true, size: 12 };
  summary.getRange("A2:B2").values = [["统计项", "数量"]];
  paintHeader(summary.getRange("A2:B2"));
  if (summaryRows.length) {
    const endRow = 2 + summaryRows.length;
    summary.getRange(`A3:B${endRow}`).values = matrix(summaryRows, ["统计项", "数量"]);
    paintBody(summary.getRange(`A3:B${endRow}`));
  }
  setColumnWidths(summary, SUMMARY_WIDTHS);
  summary.freezePanes.freezeRows(2);

  const detail = workbook.worksheets.add("逐条语义匹配");
  detail.showGridLines = false;
  detail.getRange(`A1:${colLetter(DETAIL_COLUMNS.length)}1`).values = [DETAIL_COLUMNS];
  paintHeader(detail.getRange(`A1:${colLetter(DETAIL_COLUMNS.length)}1`));
  if (detailRows.length) {
    const endRow = 1 + detailRows.length;
    detail.getRange(`A2:${colLetter(DETAIL_COLUMNS.length)}${endRow}`).values = matrix(detailRows, DETAIL_COLUMNS);
    paintBody(detail.getRange(`A2:${colLetter(DETAIL_COLUMNS.length)}${endRow}`));
    addDecisionColors(detail, 8, 2, endRow);
    addDecisionColors(detail, 9, 2, endRow);
    addDecisionColors(detail, 16, 2, endRow);
    addDecisionColors(detail, 17, 2, endRow);
  }
  setColumnWidths(detail, DETAIL_WIDTHS);
  detail.freezePanes.freezeRows(1);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputPath);

  return outputPath;
}

async function main() {
  const inputArg = process.argv[2];
  const outputArg = process.argv[3];
  if (!inputArg || !outputArg) {
    console.error("Usage: node ai_match_template_generator.mjs <input.json> <output.xlsx>");
    process.exit(1);
  }
  const inputPath = path.resolve(process.cwd(), inputArg);
  const outputPath = path.resolve(process.cwd(), outputArg);
  const saved = await generateAiMatchWorkbook({ inputPath, outputPath });
  console.log(saved);
}

const entryUrl = new URL(process.argv[1], "file://").href;
if (import.meta.url === entryUrl) {
  await main();
}
