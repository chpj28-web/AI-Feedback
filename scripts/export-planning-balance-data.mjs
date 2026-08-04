import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

const sourceDir = path.resolve("result", "WK29");
const outputFile = path.resolve("public", "planning-balance-data.json");

function text(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "object") {
    if ("text" in value) return String(value.text ?? "");
    if ("result" in value) return text(value.result);
    if (Array.isArray(value.richText)) return value.richText.map((part) => String(part.text ?? "")).join("");
  }
  return String(value);
}

function numeric(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value && typeof value === "object" && typeof value.result === "number" && Number.isFinite(value.result)) {
    return value.result;
  }
  return null;
}

function clean(value) {
  return text(value).replace(/\s+/g, " ").trim();
}

function findBalanceSheet(workbook) {
  return (
    workbook.worksheets.find((sheet) => sheet.name.toLowerCase().includes("balance")) ??
    workbook.worksheets.find((sheet) => sheet.name.includes("บาลาน")) ??
    null
  );
}

function inferWeek(fileName, worksheet) {
  const fileMatch = fileName.match(/(?:wk|w|week)\s*\.?\s*0?([1-9]|[1-4]\d|5[0-3])/i);
  if (fileMatch) return String(Number(fileMatch[1]));

  for (let rowNumber = 1; rowNumber <= Math.min(10, worksheet.rowCount); rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (let col = 1; col <= Math.min(10, worksheet.columnCount); col += 1) {
      const value = clean(row.getCell(col).value);
      if (/^(?:[1-9]|[1-4]\d|5[0-3])$/.test(value)) return String(Number(value));
    }
  }

  return "";
}

function inferFactory(fileName, worksheet) {
  for (let rowNumber = 1; rowNumber <= Math.min(8, worksheet.rowCount); rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (let col = 1; col <= Math.min(20, worksheet.columnCount); col += 1) {
      const value = clean(row.getCell(col).value);
      if (value.includes("โรง") || value.includes("รชล") || value.includes("ตัดแต่ง")) return value;
    }
  }

  return fileName
    .replace(/\.(xlsm|xlsx)$/i, "")
    .replace(/balance[_\s-]*/i, "")
    .replace(/wk\s*\.?\s*\d+/i, "")
    .replace(/\([^)]*\)/g, "")
    .trim() || fileName;
}

function columnText(worksheet, col, fromRow = 1, toRow = 15) {
  const parts = [];
  for (let rowNumber = fromRow; rowNumber <= Math.min(toRow, worksheet.rowCount); rowNumber += 1) {
    const value = clean(worksheet.getRow(rowNumber).getCell(col).value);
    if (value && value !== "[object Object]" && !parts.includes(value)) parts.push(value);
  }
  return parts.join(" ");
}

function findColumns(worksheet) {
  const columns = {};
  const maxCol = worksheet.columnCount;

  for (let col = 1; col <= maxCol; col += 1) {
    const label = columnText(worksheet, col);
    const lower = label.toLowerCase();
    const set = (key) => {
      if (!columns[key]) columns[key] = col;
    };

    if (label.includes("SAP")) set("sap");
    if (label.includes("ชิ้นส่วน") || label.includes("กลุ่มสินค้า")) set("product");
    if (label.includes("%Yield") || label.includes("% Yield")) set("yieldFg");
    if (label.includes("ผลิต") && !label.includes("ผลิตรวม")) set("production");
    if (label.includes("ยกมา") || lower.includes("stock")) set("stock");
    if (label.includes("รับโอน")) set("transferIn");
    if (label.includes("รวมผลิต") || lower.includes("total supply")) set("totalSupply");
    if (label.includes("FC Total") || label.includes("FC (kg/สัปดาห์)")) set("fcTotal");
    if (label.includes("QT total") || lower.includes("quota")) set("qtTotal");
    if (label.includes("สินค้าที่ขาด/เหลือ") || label.includes("เหลือ/ขาด") || label.includes("ของขาด/เหลือ")) {
      set("shortageSurplus");
    }
  }

  return columns;
}

function rowText(row, col) {
  return clean(row.getCell(col).value);
}

function parseRows(fileName, worksheet) {
  const columns = findColumns(worksheet);
  if (!columns.product) return [];

  const factory = inferFactory(fileName, worksheet);
  const week = inferWeek(fileName, worksheet);
  const rows = [];

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber <= 7) return;
    const product = rowText(row, columns.product);
    if (!product || product === "ชิ้นส่วน" || product === "กลุ่มสินค้า" || product.includes("วัตถุประสงค์")) return;
    if (product.length > 80) return;

    const values = {
      sap: columns.sap ? rowText(row, columns.sap) : "",
      yieldFg: columns.yieldFg ? numeric(row.getCell(columns.yieldFg).value) : null,
      production: columns.production ? numeric(row.getCell(columns.production).value) : null,
      stock: columns.stock ? numeric(row.getCell(columns.stock).value) : null,
      transferIn: columns.transferIn ? numeric(row.getCell(columns.transferIn).value) : null,
      totalSupply: columns.totalSupply ? numeric(row.getCell(columns.totalSupply).value) : null,
      fcTotal: columns.fcTotal ? numeric(row.getCell(columns.fcTotal).value) : null,
      qtTotal: columns.qtTotal ? numeric(row.getCell(columns.qtTotal).value) : null,
      shortageSurplus: columns.shortageSurplus ? numeric(row.getCell(columns.shortageSurplus).value) : null,
    };

    const hasNumber = Object.values(values).some((value) => typeof value === "number");
    if (!hasNumber) return;

    rows.push({
      id: `${fileName}|${worksheet.name}|${rowNumber}`,
      fileName,
      sheetName: worksheet.name,
      factory,
      week,
      rowNumber,
      product,
      ...values,
    });
  });

  return rows;
}

const records = [];
const files = fs.existsSync(sourceDir)
  ? fs.readdirSync(sourceDir).filter((file) => /^balance/i.test(file) && /\.(xlsm|xlsx)$/i.test(file))
  : [];

for (const fileName of files) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.join(sourceDir, fileName));
  const worksheet = findBalanceSheet(workbook);
  if (!worksheet) continue;
  records.push(...parseRows(fileName, worksheet));
}

fs.writeFileSync(
  outputFile,
  `${JSON.stringify(
    {
      generatedAt: new Date().toISOString(),
      sourceFolder: sourceDir,
      records,
    },
    null,
    2,
  )}\n`,
);

console.log(`Exported ${records.length} planning balance rows to ${outputFile}`);
