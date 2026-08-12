import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

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
  return rows.map((row) => columns.map((column) => row[column] ?? ""));
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

function setColumnWidths(sheet, widths) {
  widths.forEach((width, index) => {
    const col = colLetter(index + 1);
    sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  });
}

function addTextHighlight(sheet, config, rowStart, rowEnd) {
  const col = colLetter(config.colIndex);
  const range = sheet.getRange(`${col}${rowStart}:${col}${rowEnd}`);
  for (const rule of config.rules ?? []) {
    range.conditionalFormats.add("containsText", {
      text: rule.text,
      format: {
        fill: { color: rule.fill },
        font: { color: rule.fontColor, bold: true },
      },
    });
  }
}

function buildSheet(workbook, cfg) {
  const sheet = workbook.worksheets.add(cfg.name);
  sheet.showGridLines = false;

  let rowCursor = 1;
  if (cfg.title) {
    sheet.getRange(`A${rowCursor}`).values = [[cfg.title]];
    sheet.getRange(`A${rowCursor}`).format.font = { bold: true, size: 12 };
    rowCursor += 1;
  }

  if (cfg.summaryColumns && cfg.summaryRows) {
    const end = colLetter(cfg.summaryColumns.length);
    sheet.getRange(`A${rowCursor}:${end}${rowCursor}`).values = [cfg.summaryColumns];
    paintHeader(sheet.getRange(`A${rowCursor}:${end}${rowCursor}`));
    if (cfg.summaryRows.length) {
      const endRow = rowCursor + cfg.summaryRows.length;
      sheet.getRange(`A${rowCursor + 1}:${end}${endRow}`).values = toMatrix(cfg.summaryRows, cfg.summaryColumns);
      paintBody(sheet.getRange(`A${rowCursor + 1}:${end}${endRow}`));
      rowCursor = endRow + 2;
    } else {
      rowCursor += 2;
    }
  }

  const columns = cfg.columns;
  const end = colLetter(columns.length);
  sheet.getRange(`A${rowCursor}:${end}${rowCursor}`).values = [columns];
  paintHeader(sheet.getRange(`A${rowCursor}:${end}${rowCursor}`));
  if (cfg.rows?.length) {
    const endRow = rowCursor + cfg.rows.length;
    sheet.getRange(`A${rowCursor + 1}:${end}${endRow}`).values = toMatrix(cfg.rows, columns);
    paintBody(sheet.getRange(`A${rowCursor + 1}:${end}${endRow}`));
    for (const highlight of cfg.highlights ?? []) {
      addTextHighlight(sheet, highlight, rowCursor + 1, endRow);
    }
  }
  sheet.freezePanes.freezeRows(rowCursor);
  setColumnWidths(sheet, cfg.widths);
}

export async function buildWorkbookFromJson({ inputPath, outputPath }) {
  const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
  const workbook = Workbook.create();
  for (const sheet of payload.sheets) {
    buildSheet(workbook, sheet);
  }

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputPath);

  const firstSheet = payload.sheets[0];
  const inspect = await workbook.inspect({
    kind: "table",
    sheetId: firstSheet.name,
    range: `A1:${colLetter(Math.min(firstSheet.columns.length || firstSheet.summaryColumns?.length || 6, 8))}${Math.min((firstSheet.rows?.length || 0) + 3, 8)}`,
    include: "values",
    tableMaxRows: 8,
    tableMaxCols: 8,
  });
  console.log(inspect.ndjson);

  for (const sheet of payload.sheets.slice(0, 2)) {
    const width = Math.min(sheet.columns.length || 6, 8);
    const height = Math.min((sheet.rows?.length || 0) + 3, 8);
    const blob = await workbook.render({
      sheetName: sheet.name,
      range: `A1:${colLetter(width)}${height}`,
      scale: 1.2,
      format: "png",
      autoCrop: "all",
    });
    const bytes = new Uint8Array(await blob.arrayBuffer());
    await fs.writeFile(path.join(path.dirname(outputPath), `${sheet.name}.png`), bytes);
  }
}

async function main() {
  const inputArg = process.argv[2];
  const outputArg = process.argv[3];
  if (!inputArg || !outputArg) {
    console.error("Usage: node generic_workbook_builder.mjs <input.json> <output.xlsx>");
    process.exit(1);
  }
  await buildWorkbookFromJson({
    inputPath: path.resolve(process.cwd(), inputArg),
    outputPath: path.resolve(process.cwd(), outputArg),
  });
  console.log(path.resolve(process.cwd(), outputArg));
}

const entryUrl = new URL(process.argv[1], "file://").href;
if (import.meta.url === entryUrl) {
  await main();
}
