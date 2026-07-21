import ExcelJS from "exceljs";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";

const sourceFile =
  process.argv[2] ??
  "C:\\Users\\chompoopan.jan\\Downloads\\sigmas_20260512T122904.xlsx";
const outputFile = path.join(process.cwd(), "public", "ai-highlighted-data.json");

function text(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "object") {
    if ("text" in value) {
      return String(value.text ?? "");
    }
    if ("result" in value) {
      return text(value.result);
    }
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text ?? "").join("");
    }
  }

  return String(value);
}

function isYellow(cell) {
  const color = cell.fill?.fgColor?.argb ?? cell.fill?.bgColor?.argb ?? "";
  return color.toUpperCase().endsWith("FFFF00");
}

function numberValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (value && typeof value === "object" && typeof value.result === "number") {
    return value.result;
  }

  return null;
}

function findHeader(headers, name) {
  return headers.findIndex((header) => header === name) + 1;
}

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(sourceFile);

const records = [];
const sheets = [];

for (const worksheet of workbook.worksheets) {
  const headerRow = worksheet.getRow(1);
  const headers = [];
  const yellowColumns = [];

  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const header = text(cell.value).trim();
    headers[colNumber - 1] = header;

    if (header && isYellow(cell)) {
      yellowColumns.push(colNumber);
    }
  });

  if (yellowColumns.length === 0) {
    continue;
  }

  const factoryCol = findHeader(headers, "WarehouseForPlan1");
  const sourceFactoryCol = findHeader(headers, "SourceWarehouseForPlan1");
  const destinationFactoryCol = findHeader(headers, "DestinationWarehouseForPlan1");
  const weekCol = findHeader(headers, "weekNo");
  const dayCol = findHeader(headers, "DayKey") || findHeader(headers, "วันที่โอน");

  const map = new Map();

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }

    let factory = "ไม่ระบุโรงงาน";
    if (factoryCol) {
      factory = text(row.getCell(factoryCol).value).trim() || factory;
    } else if (sourceFactoryCol && destinationFactoryCol) {
      const source = text(row.getCell(sourceFactoryCol).value).trim();
      const destination = text(row.getCell(destinationFactoryCol).value).trim();
      factory = `${source || "ไม่ระบุ"} -> ${destination || "ไม่ระบุ"}`;
    } else if (sourceFactoryCol) {
      factory = text(row.getCell(sourceFactoryCol).value).trim() || factory;
    } else if (destinationFactoryCol) {
      factory = text(row.getCell(destinationFactoryCol).value).trim() || factory;
    }

    const week = weekCol ? text(row.getCell(weekCol).value).trim() : "";
    const day = dayCol ? text(row.getCell(dayCol).value).trim() : "";

    for (const colNumber of yellowColumns) {
      const metric = headers[colNumber - 1];
      const cell = row.getCell(colNumber);
      const valueText = text(cell.value).trim();
      const numericValue = numberValue(cell.value);

      if (!valueText && numericValue === null) {
        continue;
      }

      const key = `${worksheet.name}|${factory}|${metric}`;
      if (!map.has(key)) {
        map.set(key, {
          sheet: worksheet.name,
          factory,
          metric,
          rows: 0,
          numericCount: 0,
          sum: 0,
          min: null,
          max: null,
          examples: new Set(),
          weeks: new Set(),
          days: new Set(),
        });
      }

      const item = map.get(key);
      item.rows += 1;

      if (week) {
        item.weeks.add(week);
      }
      if (day && item.days.size < 10) {
        item.days.add(day);
      }

      if (numericValue !== null) {
        item.numericCount += 1;
        item.sum += numericValue;
        item.min = item.min === null ? numericValue : Math.min(item.min, numericValue);
        item.max = item.max === null ? numericValue : Math.max(item.max, numericValue);
      } else if (item.examples.size < 3) {
        item.examples.add(valueText);
      }
    }
  });

  sheets.push({
    name: worksheet.name,
    rowCount: Math.max(worksheet.rowCount - 1, 0),
    yellowMetrics: yellowColumns.map((colNumber) => headers[colNumber - 1]),
  });

  for (const item of map.values()) {
    const average =
      item.numericCount > 0 ? Number((item.sum / item.numericCount).toFixed(4)) : null;
    const aiValue = item.numericCount > 0 ? Number(item.sum.toFixed(4)) : null;

    records.push({
      id: `${item.sheet}|${item.factory}|${item.metric}`,
      sheet: item.sheet,
      factory: item.factory,
      metric: item.metric,
      kind: item.numericCount > 0 ? "number" : "text",
      aiValue,
      average,
      min: item.min,
      max: item.max,
      rows: item.rows,
      numericCount: item.numericCount,
      examples: [...item.examples],
      weeks: [...item.weeks].sort(),
      days: [...item.days].sort(),
    });
  }
}

records.sort((a, b) =>
  [a.sheet, a.factory, a.metric].join("|").localeCompare(
    [b.sheet, b.factory, b.metric].join("|"),
    "th",
  ),
);

await mkdir(path.dirname(outputFile), { recursive: true });
await writeFile(
  outputFile,
  JSON.stringify(
    {
      generatedAt: new Date().toISOString(),
      sourceFile: path.basename(sourceFile),
      sheets,
      records,
    },
    null,
    2,
  ),
);

console.log(`Exported ${records.length} records to ${outputFile}`);
