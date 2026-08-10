"use client";

import {
  AlertCircle,
  BarChart3,
  Brain,
  CalendarDays,
  CheckCircle2,
  ChevronDown,
  Database,
  Download,
  Edit3,
  Factory,
  Gauge,
  HelpCircle,
  LogOut,
  Menu,
  MessageCircle,
  Paperclip,
  Save,
  Search,
  Settings,
  Sparkles,
  Target,
  TrendingUp,
  Upload,
  User,
  Users,
} from "lucide-react";
import Image from "next/image";
import { type CSSProperties, type ReactNode, useEffect, useMemo, useState } from "react";

type AiRecord = {
  id: string;
  sheet: string;
  factory: string;
  metric: string;
  kind: "number" | "text";
  aiValue: number | null;
  average: number | null;
  min: number | null;
  max: number | null;
  rows: number;
  examples: string[];
  weeks: string[];
};

type TransferRecord = {
  id: string;
  route: string;
  source: string;
  destination: string;
  productGroup: string;
  productType: string;
  aiTransfer: number | null;
  plannedTransfer: number | null;
  fourWheel: number | null;
  sixWheel: number | null;
  tenWheel: number | null;
  weeks: string[];
};

type BalanceRecord = {
  id: string;
  tableType: "factory" | "product-group";
  sourceSheet: string;
  factory: string;
  week: string;
  productGroup: string;
  metric: string;
  aiMetric: string;
  aiValue: number | null;
  balanceValue: number | null;
};

type BalanceData = {
  generatedAt: string;
  sourceFile: string;
  aiFile: string;
  mapFile?: string;
  records: BalanceRecord[];
};

type PlanComparisonHistoryItem = {
  id: string;
  label: string;
  uploadedAt: string;
  aiFile: string;
  planFile: string;
  weeks: string[];
  recordCount: number;
};

type StoredPlanComparison = PlanComparisonHistoryItem & {
  data: BalanceData;
};

type PlanningBalanceRow = {
  id: string;
  fileName: string;
  sheetName: string;
  factory: string;
  week: string;
  rowNumber: number;
  product: string;
  sap: string;
  yieldFg: number | null;
  production: number | null;
  stock: number | null;
  transferIn: number | null;
  totalSupply: number | null;
  fcTotal: number | null;
  qtTotal: number | null;
  shortageSurplus: number | null;
};

type PlanningBalanceData = {
  generatedAt: string;
  sourceFolder: string;
  records: PlanningBalanceRow[];
};

type UnifiedPlanningRow = {
  product: string;
  sap: string;
  yieldFg: number | null;
  aiProduction: number | null;
  aiStock: number | null;
  aiTransferIn: number | null;
  aiTransferOut: number | null;
  aiTotalSupply: number | null;
  aiFcTotal: number | null;
  aiQtTotal: number | null;
  aiShortageSurplus: number | null;
};

type UnifiedPlanningNumberField = keyof Pick<
  UnifiedPlanningRow,
  | "aiProduction"
  | "aiStock"
  | "aiTransferIn"
  | "aiTransferOut"
  | "aiTotalSupply"
  | "aiFcTotal"
  | "aiQtTotal"
  | "aiShortageSurplus"
>;

type AiData = {
  generatedAt: string;
  sourceFile: string;
  records: AiRecord[];
  transferRecords?: TransferRecord[];
};

type Feedback = {
  actual: string;
  accuracy: string;
  comment: string;
};

type UploadHistoryItem = {
  type: "AI" | "Actual" | "Balance";
  name: string;
  uploadedAt: string;
};

type AppTab = "planning" | "analyze" | "factoryFeedback" | "feedback" | "transfer" | "balance" | "upload";

type CellLike = {
  value?: unknown;
  fill?: {
    fgColor?: { argb?: string };
    bgColor?: { argb?: string };
  };
};

type RowLike = {
  getCell(column: number): CellLike;
  eachCell(
    options: { includeEmpty: boolean },
    callback: (cell: CellLike, colNumber: number) => void,
  ): void;
};

type WorksheetLike = {
  name: string;
  getRow(row: number): RowLike;
  eachRow(
    options: { includeEmpty: boolean },
    callback: (row: RowLike, rowNumber: number) => void,
  ): void;
};

type WorkbookLike = {
  xlsx: {
    load(buffer: ArrayBuffer): Promise<unknown>;
  };
  worksheets: WorksheetLike[];
};

type Aggregate = {
  sheet: string;
  factory: string;
  metric: string;
  rows: number;
  numericCount: number;
  sum: number;
  min: number | null;
  max: number | null;
  examples: Set<string>;
  weeks: Set<string>;
};

type TransferAggregate = {
  route: string;
  source: string;
  destination: string;
  productGroup: string;
  productType: string;
  aiTransfer: number;
  plannedTransfer: number;
  weeks: Set<string>;
};

type TransferActualAggregate = {
  actualTransfer: number;
  productTypes: Set<string>;
};

type TransferActualPayload = Record<
  string,
  {
    actualTransfer: number;
    productTypes: string[];
  }
>;

const allSheets = "ทั้งหมด";
const storageKey = "ai-feedback-review-v1";
const uploadedDataKey = "ai-feedback-uploaded-ai-data-v3";
const uploadedActualKey = "ai-feedback-uploaded-actual-feedback-v2";
const uploadedActualTransferKey = "ai-feedback-uploaded-actual-transfer-v1";
const uploadedBalanceComparisonKey = "ai-feedback-uploaded-balance-comparison-v1";
const planComparisonHistoryKey = "ai-feedback-plan-comparison-history-v1";
const planComparisonDbName = "ai-feedback-plan-comparison-db";
const planComparisonStoreName = "planComparisons";
const uploadHistoryKey = "ai-feedback-upload-history-v1";
const numberFormatter = new Intl.NumberFormat("th-TH", {
  maximumFractionDigits: 2,
});

const rememberedFeedbackMetrics: Record<string, string[]> = {
  "1. ปริมาณตัดแต่ง": ["จำนวนหมูเข้าตัดแต่ง (head)"],
  "2. ปริมาณ Supply": [
    "Production (kg)",
    "Stock ยกมา (kg)",
    "แปลงไปเป็น SKU อื่น (kg)",
    "Transfer in (kg)",
    "Transfer out (kg)",
    "Net Transfer (kg)",
    "Total Supply (kg)",
    "ของขาด-เหลือ (kg)",
    "ของเหลือ (kg)",
    "Buffer วันถัดไป (kg)",
    "Quota (kg)",
    "ผลักดัน (kg)",
  ],
  "3. FC,QT": [
    "กำไร/ขาดทุนต่อหน่วย (Baht/kg)",
    "คาดการณ์ยอดขาย (Baht)",
    "กำไร/ขาดทุน รวม (Baht)",
    "% VDP",
  ],
  "4. โอน": [
    "SourceWarehouseForPlan1",
    "DestinationWarehouseForPlan1",
    "ปริมาณแนะนำโอน (kg)",
  ],
  "5. การใช้รถ": [
    "ปริมาณแนะนำโอน (kg)",
    "จำนวนรถ_4Wheels",
    "จำนวนรถ_6Wheels",
    "จำนวนรถ_10Wheels",
  ],
};

const actualMetricAliases: Record<string, string[]> = {
  "จำนวนหมูเข้าตัดแต่ง (head)": ["จำนวนหมูเข้าตัดแต่ง (head)"],
  "Production (kg)": ["ProductionWeight"],
  "Stock ยกมา (kg)": ["StockWeight"],
  "Transfer in (kg)": ["Received"],
  "Transfer out (kg)": ["Delivered"],
  "Total Supply (kg)": ["Total_Supply"],
  "Forecast (kg)": ["Forecast"],
  "FC (kg)": ["Forecast"],
  "Quota (kg)": ["Quota"],
  "QT (kg)": ["Quota"],
  "% ตอบกลับ Forecast": ["Quota_per_Forecast"],
  "ของขาด-เหลือ (kg)": ["Shortage_Surplus", "Sum Shortage"],
  "ขาด": ["Shortage_Surplus"],
  "เกิน": ["Shortage_Surplus"],
  "ปริมาณโอนทั้งหมด (kg)": ["Weight"],
  "ปริมาณแนะนำโอน (kg)": ["Weight"],
};

const balanceFactoryMetricMap: Record<string, string[]> = {
  "จำนวนหมูทั้งหมด (ตัว/สัปดาห์)": ["จำนวนหมูเข้าตัดแต่ง (head)"],
  "ตัดแต่งต่อวัน": ["จำนวนหมูเข้าตัดแต่ง (head)"],
  "น้ำหนักหมู": ["น้ำหนักเฉลี่ยนต่อตัว (kg)"],
};

const balanceProductMetricMap: Record<string, string[]> = {
  "Yield FG Adjust": ["% Actual Yield", "% Actual Yield"],
  "%Yield FG": ["% Actual Yield"],
  "% VDP": ["% VDP"],
  "ผลิต": ["Production (kg)"],
  "stock": ["Stock ยกมา (kg)"],
  "รับโอน": ["Transfer in (kg)"],
  "Total Supply": ["Total Supply (kg)"],
  "FC Total": ["FC (kg)"],
  "QT total": ["QT (kg)", "Quota (kg)"],
  "สินค้าที่ขาด/เหลือจากการบาล้าน": ["ของขาด-เหลือ (kg)", "ขาด", "เกิน"],
};

const navItems = [
  { label: "วางแผน", icon: Target, tab: "planning" },
  { label: "Feedback โรงงาน", icon: MessageCircle, tab: "factoryFeedback" },
  { label: "วิเคราะห์ผล", icon: TrendingUp, tab: "analyze" },
  { label: "AI vs Actual", icon: Edit3, tab: "feedback" },
  { label: "AI vs Actual (โอน)", icon: Factory, tab: "transfer" },
  { label: "AI vs แผน", icon: BarChart3, tab: "balance" },
  { label: "อัปโหลดผล AI", icon: Upload, tab: "upload" },
  { label: "ประวัติ Feedback", icon: Database, tab: "feedback" },
  { label: "วิเคราะห์ผล", icon: BarChart3 },
  { label: "รายงานสรุป", icon: Brain },
  { label: "ตั้งค่า", icon: Settings },
  { label: "ผู้ใช้งาน", icon: Users },
] as const;

function formatNumber(value: number | null) {
  return value === null ? "-" : numberFormatter.format(value);
}

function cellText(value: unknown): string {
  if (value === null || value === undefined) return "";

  if (typeof value === "object") {
    const objectValue = value as {
      text?: unknown;
      result?: unknown;
      richText?: { text?: unknown }[];
    };

    if ("text" in objectValue) return String(objectValue.text ?? "");
    if ("result" in objectValue) return cellText(objectValue.result);
    if (Array.isArray(objectValue.richText)) {
      return objectValue.richText.map((part) => String(part.text ?? "")).join("");
    }
  }

  return String(value);
}

function numericValue(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  if (value && typeof value === "object") {
    const result = (value as { result?: unknown }).result;
    if (typeof result === "number" && Number.isFinite(result)) return result;
  }

  return null;
}

function isYellow(cell: CellLike) {
  const color = cell.fill?.fgColor?.argb ?? cell.fill?.bgColor?.argb ?? "";
  return color.toUpperCase().endsWith("FFFF00");
}

function headerIndex(headers: string[], name: string) {
  const normalizedName = name.toLowerCase();
  return (
    headers.findIndex((header) => header.trim().toLowerCase() === normalizedName) + 1
  );
}

function firstHeaderIndex(headers: string[], names: string[]) {
  for (const name of names) {
    const index = headerIndex(headers, name);
    if (index > 0) return index;
  }

  return 0;
}

function normalizeWeek(value: unknown) {
  const text = cellText(value).trim();
  if (!text) return "";

  const exactNumber = Number(text);
  if (Number.isInteger(exactNumber) && exactNumber >= 1 && exactNumber <= 53) {
    return String(exactNumber);
  }

  const wkMatch = text.match(/\b(?:wk|week|สัปดาห์)\s*0?([1-9]|[1-4]\d|5[0-3])\b/i);
  if (wkMatch) return String(Number(wkMatch[1]));

  return "";
}

function uploadKindFromName(fileName: string): UploadHistoryItem["type"] | null {
  const lower = fileName.toLowerCase();
  if (lower.includes("actual")) return "Actual";
  if (lower.includes("balance")) return "Balance";
  return null;
}

function sheetFamily(sheetName: string) {
  if (sheetName.includes("ปริมาณตัดแต่ง")) return "1. ปริมาณตัดแต่ง";
  if (sheetName.includes("ปริมาณ Supply")) return "2. ปริมาณ Supply";
  if (sheetName.includes("FC") || sheetName.includes("ActualQuota")) return "3. FC,QT";
  if (sheetName.includes("โอน")) return "4. โอน";
  if (sheetName.includes("การใช้รถ")) return "5. การใช้รถ";
  return sheetName;
}

function factoryFromRow(row: RowLike, headers: string[]) {
  const factoryCol = firstHeaderIndex(headers, [
    "WarehouseForPlan1",
    "WarehouseForplan1",
    "SourceWarehouseForPlan1",
    "SourceWarehouseName",
  ]);
  const destinationCol = firstHeaderIndex(headers, [
    "DestinationWarehouseForPlan1",
    "DestinationWarehouseName",
  ]);

  if (factoryCol && destinationCol) {
    const source = cellText(row.getCell(factoryCol).value).trim() || "ไม่ระบุ";
    const destination = cellText(row.getCell(destinationCol).value).trim() || "ไม่ระบุ";
    return `${source} -> ${destination}`;
  }

  if (factoryCol) {
    return cellText(row.getCell(factoryCol).value).trim() || "ไม่ระบุโรงงาน";
  }

  if (destinationCol) {
    return cellText(row.getCell(destinationCol).value).trim() || "ไม่ระบุโรงงาน";
  }

  return "ไม่ระบุโรงงาน";
}

function score(record: AiRecord, actual: string) {
  if (!actual.trim()) return null;

  if (record.kind === "text") {
    const normalizedActual = actual.trim().toLowerCase();
    const exact = record.examples.some(
      (example) => example.trim().toLowerCase() === normalizedActual,
    );
    return exact ? 100 : 0;
  }

  const actualNumber = Number(actual.replaceAll(",", ""));
  if (!Number.isFinite(actualNumber) || record.aiValue === null) return null;

  const denominator = Math.max(Math.abs(actualNumber), 1);
  const errorRate = Math.abs(record.aiValue - actualNumber) / denominator;
  return Math.max(0, Math.round((1 - errorRate) * 100));
}

function scoreLabel(value: number | null) {
  if (value === null) return "รอข้อมูล";
  if (value >= 80) return "ดี";
  if (value >= 60) return "ต่างปานกลาง";
  return "ต่างกันมาก";
}

function scoreTone(value: number | null) {
  if (value === null) return "bg-slate-100 text-slate-500";
  if (value >= 80) return "bg-emerald-50 text-emerald-700";
  if (value >= 60) return "bg-orange-50 text-orange-700";
  return "bg-rose-50 text-rose-700";
}

function difference(record: AiRecord, actual: string) {
  const actualNumber = Number(actual.replaceAll(",", ""));
  if (record.aiValue === null || !Number.isFinite(actualNumber)) return null;
  return actualNumber - record.aiValue;
}

function isTransferRecord(record: AiRecord) {
  return record.sheet === "4. โอน" || record.sheet === "5. การใช้รถ";
}

function transferKey(source: string, destination: string, productGroup: string) {
  return [source, destination, productGroup].join("|");
}

function transferRowId(source: string, destination: string, productGroup: string, productType: string) {
  return [source, destination, productGroup, productType].join("|");
}

async function parseAiWorkbook(file: File): Promise<AiData> {
  const ExcelJS = await import("exceljs");
  const workbook = new ExcelJS.Workbook() as unknown as WorkbookLike;
  await workbook.xlsx.load(await file.arrayBuffer());

  const records: AiRecord[] = [];
  const transferMap = new Map<string, TransferAggregate>();
  const vehicleMap = new Map<
    string,
    {
      fourWheel: number;
      sixWheel: number;
      tenWheel: number;
    }
  >();

  for (const worksheet of workbook.worksheets) {
    const headers: string[] = [];
    const yellowColumns: number[] = [];

    worksheet.getRow(1).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const header = cellText(cell.value).trim();
      headers[colNumber - 1] = header;
      if (header && isYellow(cell)) yellowColumns.push(colNumber);
    });

    const family = sheetFamily(worksheet.name);
    const rememberedMetrics = rememberedFeedbackMetrics[family] ?? [];
    const metricColumns =
      rememberedMetrics.length > 0
        ? rememberedMetrics
            .map((metric) => ({
              column: headerIndex(headers, metric),
              metric,
            }))
            .filter((item) => item.column > 0)
        : yellowColumns.map((column) => ({
            column,
            metric: headers[column - 1],
          }));

    if (metricColumns.length === 0) continue;

    const weekCol = headerIndex(headers, "weekNo");
    const map = new Map<string, Aggregate>();
    const isTransferSheet = family === "4. โอน";
    const isVehicleSheet = family === "5. การใช้รถ";
    const sourceCol = firstHeaderIndex(headers, ["SourceWarehouseForPlan1", "SourceWarehouseName"]);
    const destinationCol = firstHeaderIndex(headers, [
      "DestinationWarehouseForPlan1",
      "DestinationWarehouseName",
    ]);
    const productGroupCol = firstHeaderIndex(headers, ["ProductForPlan19", "ProductForPlan10"]);
    const productTypeCol = firstHeaderIndex(headers, ["ProductForPlan19Custom", "ProductName"]);
    const aiTransferCol = firstHeaderIndex(headers, ["ปริมาณแนะนำโอน (kg)"]);
    const plannedTransferCol = firstHeaderIndex(headers, ["ปริมาณโอนทั้งหมด (kg)"]);
    const fourWheelCol = firstHeaderIndex(headers, ["จำนวนรถ_4Wheels"]);
    const sixWheelCol = firstHeaderIndex(headers, ["จำนวนรถ_6Wheels"]);
    const tenWheelCol = firstHeaderIndex(headers, ["จำนวนรถ_10Wheels"]);

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;

      const factory = factoryFromRow(row, headers);
      const week = weekCol ? normalizeWeek(row.getCell(weekCol).value) : "";

      if (isTransferSheet && sourceCol && destinationCol && productGroupCol && aiTransferCol) {
        const source = cellText(row.getCell(sourceCol).value).trim() || "ไม่ระบุ";
        const destination = cellText(row.getCell(destinationCol).value).trim() || "ไม่ระบุ";
        const productGroup = cellText(row.getCell(productGroupCol).value).trim() || "ไม่ระบุกลุ่ม";
        const productType =
          productTypeCol > 0
            ? cellText(row.getCell(productTypeCol).value).trim() || productGroup
            : productGroup;
        const key = transferRowId(source, destination, productGroup, productType);
        const current =
          transferMap.get(key) ??
          ({
            route: `${source} -> ${destination}`,
            source,
            destination,
            productGroup,
            productType,
            aiTransfer: 0,
            plannedTransfer: 0,
            weeks: new Set<string>(),
          } satisfies TransferAggregate);
        const aiTransfer = numericValue(row.getCell(aiTransferCol).value);
        const plannedTransfer =
          plannedTransferCol > 0 ? numericValue(row.getCell(plannedTransferCol).value) : null;
        if (aiTransfer !== null) current.aiTransfer += aiTransfer;
        if (plannedTransfer !== null) current.plannedTransfer += plannedTransfer;
        if (week) current.weeks.add(week);
        transferMap.set(key, current);
      }

      if (isVehicleSheet && sourceCol && destinationCol) {
        const source = cellText(row.getCell(sourceCol).value).trim() || "ไม่ระบุ";
        const destination = cellText(row.getCell(destinationCol).value).trim() || "ไม่ระบุ";
        const key = `${source} -> ${destination}`;
        const current =
          vehicleMap.get(key) ??
          ({
            fourWheel: 0,
            sixWheel: 0,
            tenWheel: 0,
          });
        current.fourWheel += fourWheelCol > 0 ? numericValue(row.getCell(fourWheelCol).value) ?? 0 : 0;
        current.sixWheel += sixWheelCol > 0 ? numericValue(row.getCell(sixWheelCol).value) ?? 0 : 0;
        current.tenWheel += tenWheelCol > 0 ? numericValue(row.getCell(tenWheelCol).value) ?? 0 : 0;
        vehicleMap.set(key, current);
      }

      for (const { column, metric } of metricColumns) {
        const cell = row.getCell(column);
        const textValue = cellText(cell.value).trim();
        const value = numericValue(cell.value);
        if (!textValue && value === null) continue;

        const key = `${family}|${factory}|${metric}`;
        if (!map.has(key)) {
          map.set(key, {
            sheet: family,
            factory,
            metric,
            rows: 0,
            numericCount: 0,
            sum: 0,
            min: null,
            max: null,
            examples: new Set<string>(),
            weeks: new Set<string>(),
          });
        }

        const item = map.get(key);
        if (!item) continue;

        item.rows += 1;
        if (week) item.weeks.add(week);

        if (value !== null) {
          item.numericCount += 1;
          item.sum += value;
          item.min = item.min === null ? value : Math.min(item.min, value);
          item.max = item.max === null ? value : Math.max(item.max, value);
        } else if (item.examples.size < 3) {
          item.examples.add(textValue);
        }
      }
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
        examples: [...item.examples],
        weeks: [...item.weeks].sort(),
      });
    }
  }

  if (records.length === 0) {
    throw new Error(
      "ไฟล์นี้ไม่พบหัวข้อ feedback ที่ระบบจำไว้ กรุณาตรวจว่าหัวคอลัมน์ตรงกับไฟล์ config สีเหลืองชุดแรก",
    );
  }

  records.sort((a, b) =>
    [a.sheet, a.factory, a.metric]
      .join("|")
      .localeCompare([b.sheet, b.factory, b.metric].join("|"), "th"),
  );

  const transferRecords = [...transferMap.values()]
    .map((item) => {
      const vehicle = vehicleMap.get(item.route);
      return {
        id: transferRowId(item.source, item.destination, item.productGroup, item.productType),
        route: item.route,
        source: item.source,
        destination: item.destination,
        productGroup: item.productGroup,
        productType: item.productType,
        aiTransfer: Number(item.aiTransfer.toFixed(4)),
        plannedTransfer: Number(item.plannedTransfer.toFixed(4)),
        fourWheel: vehicle ? Number(vehicle.fourWheel.toFixed(4)) : null,
        sixWheel: vehicle ? Number(vehicle.sixWheel.toFixed(4)) : null,
        tenWheel: vehicle ? Number(vehicle.tenWheel.toFixed(4)) : null,
        weeks: [...item.weeks].sort((a, b) => Number(a) - Number(b)),
      } satisfies TransferRecord;
    })
    .sort((a, b) =>
      [a.source, a.destination, a.productGroup, a.productType]
        .join("|")
        .localeCompare([b.source, b.destination, b.productGroup, b.productType].join("|"), "th"),
    );

  return {
    generatedAt: new Date().toISOString(),
    sourceFile: file.name,
    records,
    transferRecords,
  };
}

async function parseActualWorkbook(
  file: File,
  aiRecords: AiRecord[],
): Promise<Record<string, Feedback>> {
  const ExcelJS = await import("exceljs");
  const workbook = new ExcelJS.Workbook() as unknown as WorkbookLike;
  await workbook.xlsx.load(await file.arrayBuffer());
  const values = new Map<string, number>();

  for (const worksheet of workbook.worksheets) {
    const headers: string[] = [];

    worksheet.getRow(1).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      headers[colNumber - 1] = cellText(cell.value).trim();
    });

    const family = sheetFamily(worksheet.name);
    const relevantMetrics = Array.from(
      new Set(
        aiRecords
          .filter((record) => record.sheet === family)
          .map((record) => record.metric),
      ),
    );

    const metricColumns = relevantMetrics
      .map((metric) => ({
        metric,
        column: firstHeaderIndex(headers, actualMetricAliases[metric] ?? [metric]),
      }))
      .filter((item) => item.column > 0);

    if (metricColumns.length === 0) continue;

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;

      const factory = factoryFromRow(row, headers);

      for (const { metric, column } of metricColumns) {
        const value = numericValue(row.getCell(column).value);
        if (value === null) continue;

        const key = `${family}|${factory}|${metric}`;
        values.set(key, (values.get(key) ?? 0) + value);
      }
    });
  }

  const actualFeedback: Record<string, Feedback> = {};

  for (const record of aiRecords) {
    const value = values.get(`${record.sheet}|${record.factory}|${record.metric}`);
    if (value !== undefined) {
      actualFeedback[record.id] = {
        actual: String(Number(value.toFixed(4))),
        accuracy: "",
        comment: "",
      };
    }
  }

  return actualFeedback;
}

async function parseActualTransferWorkbook(file: File): Promise<TransferActualPayload> {
  const ExcelJS = await import("exceljs");
  const workbook = new ExcelJS.Workbook() as unknown as WorkbookLike;
  await workbook.xlsx.load(await file.arrayBuffer());
  const values = new Map<string, TransferActualAggregate>();

  for (const worksheet of workbook.worksheets) {
    if (sheetFamily(worksheet.name) !== "4. โอน") continue;

    const headers: string[] = [];
    worksheet.getRow(1).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      headers[colNumber - 1] = cellText(cell.value).trim();
    });

    const sourceCol = firstHeaderIndex(headers, ["SourceWarehouseForPlan1", "SourceWarehouseName"]);
    const destinationCol = firstHeaderIndex(headers, [
      "DestinationWarehouseForPlan1",
      "DestinationWarehouseName",
    ]);
    const productGroupCol = firstHeaderIndex(headers, ["ProductForPlan19", "ProductForPlan10"]);
    const productTypeCol = firstHeaderIndex(headers, [
      "ProductName",
      "ProductForPlan19Custom",
      "ProductForPlan10",
    ]);
    const weightCol = firstHeaderIndex(headers, ["Weight", "ปริมาณโอนทั้งหมด (kg)"]);
    if (!sourceCol || !destinationCol || !productGroupCol || !weightCol) continue;

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;

      const source = cellText(row.getCell(sourceCol).value).trim() || "ไม่ระบุ";
      const destination = cellText(row.getCell(destinationCol).value).trim() || "ไม่ระบุ";
      const productGroup = cellText(row.getCell(productGroupCol).value).trim() || "ไม่ระบุกลุ่ม";
      const productType =
        productTypeCol > 0
          ? cellText(row.getCell(productTypeCol).value).trim() || productGroup
          : productGroup;
      const weight = numericValue(row.getCell(weightCol).value);
      if (weight === null) return;

      const key = transferKey(source, destination, productGroup);
      const current =
        values.get(key) ??
        ({
          actualTransfer: 0,
          productTypes: new Set<string>(),
        } satisfies TransferActualAggregate);
      current.actualTransfer += weight;
      if (productType) current.productTypes.add(productType);
      values.set(key, current);
    });
  }

  return Object.fromEntries(
    [...values.entries()].map(([key, value]) => [
      key,
      {
        actualTransfer: Number(value.actualTransfer.toFixed(4)),
        productTypes: [...value.productTypes].sort((a, b) => a.localeCompare(b, "th")),
      },
    ]),
  );
}

async function parseBalanceComparisonWorkbook(aiFile: File, balanceFile: File): Promise<BalanceData> {
  const ExcelJS = await import("exceljs");
  const aiWorkbook = new ExcelJS.Workbook() as unknown as WorkbookLike;
  const balanceWorkbook = new ExcelJS.Workbook() as unknown as WorkbookLike;

  await Promise.all([
    aiWorkbook.xlsx.load(await aiFile.arrayBuffer()),
    balanceWorkbook.xlsx.load(await balanceFile.arrayBuffer()),
  ]);

  const aiFactoryValues = new Map<string, { value: number; sourceSheet: string; aiMetric: string }>();
  const aiProductValues = new Map<string, { value: number; sourceSheet: string; aiMetric: string }>();

  for (const worksheet of aiWorkbook.worksheets) {
    const family = sheetFamily(worksheet.name);
    const headers = headersFromWorksheet(worksheet);
    const weekCol = headerIndex(headers, "weekNo");
    const factoryCol = firstHeaderIndex(headers, ["WarehouseForPlan1", "WarehouseForplan1"]);
    const productGroupCol = firstHeaderIndex(headers, ["ProductForPlan19"]);

    if (!weekCol || !factoryCol) continue;

    if (family === "1. ปริมาณตัดแต่ง") {
      const dayCol = firstHeaderIndex(headers, ["DayKey"]);
      const dayKeys = new Map<string, Set<string>>();

      for (const [balanceMetric, aiMetrics] of Object.entries(balanceFactoryMetricMap)) {
        const metricCol = firstHeaderIndex(headers, aiMetrics);
        if (!metricCol) continue;
        const aiMetric = headers[metricCol - 1] || aiMetrics[0];

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber === 1) return;
          const week = normalizeWeek(row.getCell(weekCol).value);
          const factory = cellText(row.getCell(factoryCol).value).trim() || "ไม่ระบุโรงงาน";
          const value = numericValue(row.getCell(metricCol).value);
          if (!week || value === null) return;

          const key = balanceKey("factory", factory, week, "", balanceMetric);
          const current = aiFactoryValues.get(key) ?? { value: 0, sourceSheet: family, aiMetric };
          current.value += value;
          aiFactoryValues.set(key, current);

          if (dayCol) {
            const dayKey = cellText(row.getCell(dayCol).value).trim();
            if (dayKey) {
              const daySetKey = [factory, week].join("|");
              const days = dayKeys.get(daySetKey) ?? new Set<string>();
              days.add(dayKey);
              dayKeys.set(daySetKey, days);
            }
          }
        });
      }

      for (const [key, item] of aiFactoryValues.entries()) {
        const [tableType, factory, week, , metric] = key.split("|");
        if (tableType !== "factory" || metric !== "ตัดแต่งต่อวัน") continue;
        const days = dayKeys.get([factory, week].join("|"))?.size ?? 0;
        if (days > 0) item.value = item.value / days;
      }
    }

    if ((family === "2. ปริมาณ Supply" || family === "3. FC,QT") && productGroupCol) {
      for (const [balanceMetric, aiMetrics] of Object.entries(balanceProductMetricMap)) {
        const metricCol = firstHeaderIndex(headers, aiMetrics);
        if (!metricCol) continue;
        const aiMetric = headers[metricCol - 1] || aiMetrics[0];

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber === 1) return;
          const week = normalizeWeek(row.getCell(weekCol).value);
          const factory = cellText(row.getCell(factoryCol).value).trim() || "ไม่ระบุโรงงาน";
          const productGroup = cellText(row.getCell(productGroupCol).value).trim() || "ไม่ระบุกลุ่ม";
          const value = numericValue(row.getCell(metricCol).value);
          if (!week || value === null) return;

          const key = balanceKey("product-group", factory, week, productGroup, balanceMetric);
          const current = aiProductValues.get(key) ?? { value: 0, sourceSheet: family, aiMetric };
          current.value += value;
          aiProductValues.set(key, current);
        });
      }
    }
  }

  const balanceWorksheet =
    balanceWorkbook.worksheets.find((worksheet) => worksheet.name.includes("Balance")) ?? balanceWorkbook.worksheets[0];
  if (!balanceWorksheet) throw new Error("ไม่พบชีท Balance ในไฟล์แผนที่อัปโหลด");

  const balanceFactoryValues = new Map<string, { value: number; sourceSheet: string }>();
  const balanceProductValues = new Map<string, { value: number; sourceSheet: string }>();
  const balanceHeaders = headersFromWorksheet(balanceWorksheet);
  const balanceFactoryCol = firstHeaderIndex(balanceHeaders, ["โรงงาน"]);
  const balanceWeekCol = firstHeaderIndex(balanceHeaders, ["Forecast สัปดาห์ที่", "weekNo", "Weekno"]);
  const balanceProductGroupCol = firstHeaderIndex(balanceHeaders, ["กลุ่มสินค้า", "ProductForPlan19"]);

  if (!balanceFactoryCol || !balanceWeekCol) {
    throw new Error("ไฟล์แผนไม่มีคอลัมน์ โรงงาน หรือ Forecast สัปดาห์ที่");
  }

  for (const metric of Object.keys(balanceFactoryMetricMap)) {
    const metricCol = headerIndex(balanceHeaders, metric);
    if (!metricCol) continue;

    balanceWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;
      const factory = cellText(row.getCell(balanceFactoryCol).value).trim() || "ไม่ระบุโรงงาน";
      const week = normalizeWeek(row.getCell(balanceWeekCol).value);
      const value = numericValue(row.getCell(metricCol).value);
      if (!week || value === null) return;
      const key = balanceKey("factory", factory, week, "", metric);
      const current = balanceFactoryValues.get(key) ?? { value: 0, sourceSheet: balanceWorksheet.name };
      current.value += value;
      balanceFactoryValues.set(key, current);
    });
  }

  if (balanceProductGroupCol) {
    for (const metric of Object.keys(balanceProductMetricMap)) {
      const metricCol = headerIndex(balanceHeaders, metric);
      if (!metricCol) continue;

      balanceWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber === 1) return;
        const factory = cellText(row.getCell(balanceFactoryCol).value).trim() || "ไม่ระบุโรงงาน";
        const week = normalizeWeek(row.getCell(balanceWeekCol).value);
        const productGroup = cellText(row.getCell(balanceProductGroupCol).value).trim() || "ไม่ระบุกลุ่ม";
        const value = numericValue(row.getCell(metricCol).value);
        if (!week || value === null) return;
        const key = balanceKey("product-group", factory, week, productGroup, metric);
        const current = balanceProductValues.get(key) ?? { value: 0, sourceSheet: balanceWorksheet.name };
        current.value += value;
        balanceProductValues.set(key, current);
      });
    }
  }

  const records: BalanceRecord[] = [];
  const allKeys = new Set([
    ...aiFactoryValues.keys(),
    ...balanceFactoryValues.keys(),
    ...aiProductValues.keys(),
    ...balanceProductValues.keys(),
  ]);

  for (const key of allKeys) {
    const [tableType, factory, week, productGroup, metric] = key.split("|") as [
      BalanceRecord["tableType"],
      string,
      string,
      string,
      string,
    ];
    const ai = tableType === "factory" ? aiFactoryValues.get(key) : aiProductValues.get(key);
    const balance = tableType === "factory" ? balanceFactoryValues.get(key) : balanceProductValues.get(key);

    records.push({
      id: `${key}|${ai?.sourceSheet ?? balance?.sourceSheet ?? ""}`,
      tableType,
      sourceSheet: ai?.sourceSheet ?? balance?.sourceSheet ?? "",
      factory,
      week,
      productGroup,
      metric,
      aiMetric: ai?.aiMetric ?? "",
      aiValue: Number((ai?.value ?? 0).toFixed(4)),
      balanceValue: Number((balance?.value ?? 0).toFixed(4)),
    });
  }

  records.sort((a, b) =>
    [a.tableType, a.factory, a.week, a.productGroup, a.metric]
      .join("|")
      .localeCompare([b.tableType, b.factory, b.week, b.productGroup, b.metric].join("|"), "th"),
  );

  if (records.length === 0) {
    throw new Error("ยังจับคู่ข้อมูล AI กับแผนไม่ได้ กรุณาตรวจหัวคอลัมน์และเลขสัปดาห์ในสองไฟล์");
  }

  return {
    sourceFile: balanceFile.name,
    aiFile: aiFile.name,
    mapFile: "browser upload mapping",
    generatedAt: new Date().toISOString(),
    records,
  };
}

async function parseBalanceComparisonFromAiData(aiData: AiData, balanceFile: File): Promise<BalanceData> {
  const ExcelJS = await import("exceljs");
  const balanceWorkbook = new ExcelJS.Workbook() as unknown as WorkbookLike;
  await balanceWorkbook.xlsx.load(await balanceFile.arrayBuffer());

  const balanceWorksheet =
    balanceWorkbook.worksheets.find((worksheet) => worksheet.name.includes("Balance")) ?? balanceWorkbook.worksheets[0];
  if (!balanceWorksheet) throw new Error("ไม่พบชีท Balance ในไฟล์แผนที่อัปโหลด");

  const headers = headersFromWorksheet(balanceWorksheet);
  const factoryCol = firstHeaderIndex(headers, ["โรงงาน"]);
  const weekCol = firstHeaderIndex(headers, ["Forecast สัปดาห์ที่", "weekNo", "Weekno"]);
  const productGroupCol = firstHeaderIndex(headers, ["กลุ่มสินค้า", "ProductForPlan19"]);
  if (!factoryCol || !weekCol) throw new Error("ไฟล์แผนไม่มีคอลัมน์ โรงงาน หรือ Forecast สัปดาห์ที่");

  const balanceValues = new Map<string, { value: number; sourceSheet: string }>();
  const metricNames = [...Object.keys(balanceFactoryMetricMap), ...Object.keys(balanceProductMetricMap)];

  for (const metric of metricNames) {
    const metricCol = headerIndex(headers, metric);
    if (!metricCol) continue;

    balanceWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;
      const factory = cellText(row.getCell(factoryCol).value).trim() || "ไม่ระบุโรงงาน";
      const week = normalizeWeek(row.getCell(weekCol).value);
      const value = numericValue(row.getCell(metricCol).value);
      if (!week || value === null) return;

      const tableType: BalanceRecord["tableType"] = metric in balanceFactoryMetricMap ? "factory" : "product-group";
      const productGroup =
        tableType === "product-group" && productGroupCol
          ? cellText(row.getCell(productGroupCol).value).trim() || "รวมทุกกลุ่มสินค้า"
          : "";
      const key = balanceKey(tableType, factory, week, productGroup, metric);
      const current = balanceValues.get(key) ?? { value: 0, sourceSheet: balanceWorksheet.name };
      current.value += value;
      balanceValues.set(key, current);
    });
  }

  const aiValues = new Map<string, { value: number; sourceSheet: string; aiMetric: string }>();

  for (const record of aiData.records) {
    if (record.kind !== "number" || record.aiValue === null) continue;

    const factoryMetric = Object.entries(balanceFactoryMetricMap).find(([, aiMetrics]) =>
      aiMetrics.includes(record.metric),
    );
    const productMetric = Object.entries(balanceProductMetricMap).find(([, aiMetrics]) =>
      aiMetrics.includes(record.metric),
    );
    const matchedMetric = factoryMetric?.[0] ?? productMetric?.[0];
    if (!matchedMetric) continue;

    const tableType: BalanceRecord["tableType"] = factoryMetric ? "factory" : "product-group";
    const weeks = record.weeks.length > 0 ? record.weeks : [""];
    for (const week of weeks) {
      if (!week) continue;
      if (tableType === "factory") {
        const key = balanceKey("factory", record.factory, week, "", matchedMetric);
        aiValues.set(key, {
          value: (aiValues.get(key)?.value ?? 0) + record.aiValue,
          sourceSheet: record.sheet,
          aiMetric: record.metric,
        });
        continue;
      }

      const matchingBalanceKeys = [...balanceValues.keys()].filter((key) => {
        const [keyType, keyFactory, keyWeek, , keyMetric] = key.split("|");
        return keyType === "product-group" && keyFactory === record.factory && keyWeek === week && keyMetric === matchedMetric;
      });
      const targetKeys =
        matchingBalanceKeys.length > 0
          ? matchingBalanceKeys
          : [balanceKey("product-group", record.factory, week, "รวมทุกกลุ่มสินค้า", matchedMetric)];
      const splitValue = record.aiValue / targetKeys.length;

      for (const key of targetKeys) {
        aiValues.set(key, {
          value: (aiValues.get(key)?.value ?? 0) + splitValue,
          sourceSheet: record.sheet,
          aiMetric: record.metric,
        });
      }
    }
  }

  const records: BalanceRecord[] = [];
  const allKeys = new Set([...balanceValues.keys(), ...aiValues.keys()]);
  for (const key of allKeys) {
    const [tableType, factory, week, productGroup, metric] = key.split("|") as [
      BalanceRecord["tableType"],
      string,
      string,
      string,
      string,
    ];
    const balance = balanceValues.get(key);
    const ai = aiValues.get(key);
    records.push({
      id: `${key}|fallback`,
      tableType,
      sourceSheet: ai?.sourceSheet ?? balance?.sourceSheet ?? "",
      factory,
      week,
      productGroup,
      metric,
      aiMetric: ai?.aiMetric ?? "",
      aiValue: Number((ai?.value ?? 0).toFixed(4)),
      balanceValue: Number((balance?.value ?? 0).toFixed(4)),
    });
  }

  if (records.length === 0) {
    throw new Error("ยังจับคู่ข้อมูล AI กับแผนไม่ได้ กรุณาตรวจหัวคอลัมน์และเลขสัปดาห์ในสองไฟล์");
  }

  return {
    sourceFile: balanceFile.name,
    aiFile: aiData.sourceFile,
    mapFile: "browser upload mapping from parsed AI data",
    generatedAt: new Date().toISOString(),
    records,
  };
}

function headersFromWorksheet(worksheet: WorksheetLike) {
  const headers: string[] = [];
  worksheet.getRow(1).eachCell({ includeEmpty: true }, (cell, colNumber) => {
    headers[colNumber - 1] = cellText(cell.value).trim();
  });
  return headers;
}

function balanceKey(
  tableType: BalanceRecord["tableType"],
  factory: string,
  week: string,
  productGroup: string,
  metric: string,
) {
  return [tableType, factory, week, productGroup, metric].join("|");
}

function planComparisonMeta(data: BalanceData): PlanComparisonHistoryItem {
  const weeks = Array.from(new Set(data.records.map((record) => normalizeWeek(record.week)).filter(Boolean))).sort(
    (a, b) => Number(b) - Number(a),
  );
  const uploadedAt = data.generatedAt || new Date().toISOString();
  const weekLabel = weeks.length > 0 ? `WK ${weeks.join(", ")}` : "ไม่พบสัปดาห์";

  return {
    id: `${uploadedAt}|${data.aiFile}|${data.sourceFile}`,
    label: `${weekLabel} · ${data.aiFile} vs ${data.sourceFile}`,
    uploadedAt,
    aiFile: data.aiFile,
    planFile: data.sourceFile,
    weeks,
    recordCount: data.records.length,
  };
}

function openPlanComparisonDb(): Promise<IDBDatabase | null> {
  if (typeof window === "undefined" || !("indexedDB" in window)) return Promise.resolve(null);

  return new Promise((resolve, reject) => {
    const request = window.indexedDB.open(planComparisonDbName, 1);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(planComparisonStoreName)) {
        database.createObjectStore(planComparisonStoreName, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function readPlanComparisonHistoryIndex() {
  if (typeof window === "undefined") return [];

  try {
    const saved = window.localStorage.getItem(planComparisonHistoryKey);
    return saved ? (JSON.parse(saved) as PlanComparisonHistoryItem[]) : [];
  } catch {
    return [];
  }
}

function writePlanComparisonHistoryIndex(items: PlanComparisonHistoryItem[]) {
  if (typeof window === "undefined") return;

  window.localStorage.setItem(planComparisonHistoryKey, JSON.stringify(items.slice(0, 50)));
}

function mergePlanComparisonHistory(items: PlanComparisonHistoryItem[]) {
  const byId = new Map<string, PlanComparisonHistoryItem>();
  for (const item of items) byId.set(item.id, item);
  return [...byId.values()].sort(
    (a, b) => new Date(b.uploadedAt).getTime() - new Date(a.uploadedAt).getTime(),
  );
}

async function saveStoredPlanComparison(data: BalanceData) {
  const meta = planComparisonMeta(data);
  const indexedItems = mergePlanComparisonHistory([meta, ...readPlanComparisonHistoryIndex()]);
  writePlanComparisonHistoryIndex(indexedItems);

  const database = await openPlanComparisonDb();
  if (!database) return meta;
  const item: StoredPlanComparison = { ...meta, data };

  await new Promise<void>((resolve, reject) => {
    const transaction = database.transaction(planComparisonStoreName, "readwrite");
    transaction.objectStore(planComparisonStoreName).put(item);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
  database.close();
  return item;
}

async function listStoredPlanComparisons() {
  const database = await openPlanComparisonDb();
  if (!database) return readPlanComparisonHistoryIndex();

  const items = await new Promise<StoredPlanComparison[]>((resolve, reject) => {
    const transaction = database.transaction(planComparisonStoreName, "readonly");
    const request = transaction.objectStore(planComparisonStoreName).getAll();
    request.onsuccess = () => resolve(request.result as StoredPlanComparison[]);
    request.onerror = () => reject(request.error);
  });
  database.close();

  const metadata = items.map((item) => ({
      id: item.id,
      label: item.label,
      uploadedAt: item.uploadedAt,
      aiFile: item.aiFile,
      planFile: item.planFile,
      weeks: item.weeks,
      recordCount: item.recordCount,
    }));
  const merged = mergePlanComparisonHistory([...metadata, ...readPlanComparisonHistoryIndex()]);
  writePlanComparisonHistoryIndex(merged);
  return merged;
}

async function getStoredPlanComparison(id: string) {
  const database = await openPlanComparisonDb();
  if (!database) return null;

  const item = await new Promise<StoredPlanComparison | undefined>((resolve, reject) => {
    const transaction = database.transaction(planComparisonStoreName, "readonly");
    const request = transaction.objectStore(planComparisonStoreName).get(id);
    request.onsuccess = () => resolve(request.result as StoredPlanComparison | undefined);
    request.onerror = () => reject(request.error);
  });
  database.close();
  return item?.data ?? null;
}

export default function Home() {
  const [data, setData] = useState<AiData | null>(null);
  const [balanceData, setBalanceData] = useState<BalanceData | null>(null);
  const [planningBalanceData, setPlanningBalanceData] = useState<PlanningBalanceData | null>(null);
  const [sheet, setSheet] = useState(allSheets);
  const [factory, setFactory] = useState("");
  const [week, setWeek] = useState("");
  const [metricFilters, setMetricFilters] = useState<string[]>([]);
  const [statusFilters, setStatusFilters] = useState<string[]>([
    "ต่างปานกลาง",
    "ต่างกันมาก",
  ]);
  const [activeTab, setActiveTab] = useState<AppTab>("factoryFeedback");
  const [isUploading, setIsUploading] = useState(false);
  const [aiWorkbookFile, setAiWorkbookFile] = useState<File | null>(null);
  const [balanceWorkbookFile, setBalanceWorkbookFile] = useState<File | null>(null);
  const [planComparisonHistory, setPlanComparisonHistory] = useState<PlanComparisonHistoryItem[]>([]);
  const [activePlanComparisonId, setActivePlanComparisonId] = useState("");
  const [uploadedNames, setUploadedNames] = useState<{ ai?: string; actual?: string; balance?: string }>({});
  const [uploadStatus, setUploadStatus] = useState<{
    tone: "success" | "error";
    message: string;
  } | null>(null);
  const [uploadHistory, setUploadHistory] = useState<UploadHistoryItem[]>(() => {
    if (typeof window === "undefined") return [];
    const saved = window.localStorage.getItem(uploadHistoryKey);
    return saved ? JSON.parse(saved) : [];
  });
  const [transferActuals, setTransferActuals] = useState<TransferActualPayload>(() => {
    if (typeof window === "undefined") return {};
    const saved = window.localStorage.getItem(uploadedActualTransferKey);
    return saved ? JSON.parse(saved) : {};
  });
  const [feedback, setFeedback] = useState<Record<string, Feedback>>(() => {
    if (typeof window === "undefined") return {};
    const saved = window.localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : {};
  });

  useEffect(() => {
    fetch("/ai-highlighted-data.json")
      .then((response) => response.json())
      .then((defaultData: AiData) => {
        const uploaded = window.localStorage.getItem(uploadedDataKey);
        if (!uploaded) {
          setData(defaultData);
          return;
        }

        const uploadedData = JSON.parse(uploaded) as AiData;
        const needsTransferMigration =
          (!uploadedData.transferRecords || uploadedData.transferRecords.length === 0) &&
          uploadedData.sourceFile === defaultData.sourceFile;
        setData(
          needsTransferMigration
            ? {
                ...uploadedData,
                transferRecords: defaultData.transferRecords ?? [],
              }
            : uploadedData,
        );
      });

    fetch("/actual-feedback-data.json")
      .then((response) => (response.ok ? response.json() : {}))
      .then((sampleActual: Record<string, Feedback>) => {
        setFeedback((current) => {
          const merged = { ...sampleActual, ...current };
          return merged;
        });
      })
      .catch(() => undefined);

    fetch("/actual-transfer-data.json")
      .then((response) => (response.ok ? response.json() : {}))
      .then((sampleTransfers: TransferActualPayload) => {
        const uploaded = window.localStorage.getItem(uploadedActualTransferKey);
        const parsed = uploaded ? (JSON.parse(uploaded) as TransferActualPayload) : {};
        setTransferActuals(Object.keys(parsed).length > 0 ? parsed : sampleTransfers);
      })
      .catch(() => undefined);

    fetch("/balance-comparison-data.json")
      .then((response) => (response.ok ? response.json() : null))
      .then(async (comparison: BalanceData | null) => {
        const uploaded = window.localStorage.getItem(uploadedBalanceComparisonKey);
        if (uploaded) {
          const uploadedComparison = JSON.parse(uploaded) as BalanceData;
          const stored = await saveStoredPlanComparison(uploadedComparison);
          setBalanceData(uploadedComparison);
          setActivePlanComparisonId(stored.id);
          setPlanComparisonHistory(await listStoredPlanComparisons());
          return;
        }

        const storedItems = await listStoredPlanComparisons();
        setPlanComparisonHistory(storedItems);
        if (storedItems[0]) {
          const storedComparison = await getStoredPlanComparison(storedItems[0].id);
          if (storedComparison) {
            setBalanceData(storedComparison);
            setActivePlanComparisonId(storedItems[0].id);
            return;
          }
        }

        if (comparison) {
          setBalanceData(comparison);
          setActivePlanComparisonId("");
        }
      })
      .catch(() => undefined);

    fetch("/planning-balance-data.json")
      .then((response) => (response.ok ? response.json() : null))
      .then((planningData: PlanningBalanceData | null) => {
        if (planningData) setPlanningBalanceData(planningData);
      })
      .catch(() => undefined);

  }, []);

  useEffect(() => {
    window.localStorage.setItem(storageKey, JSON.stringify(feedback));
  }, [feedback]);

  useEffect(() => {
    window.localStorage.setItem(uploadedActualTransferKey, JSON.stringify(transferActuals));
  }, [transferActuals]);

  useEffect(() => {
    window.localStorage.setItem(uploadHistoryKey, JSON.stringify(uploadHistory));
  }, [uploadHistory]);

  const records = useMemo(() => data?.records ?? [], [data]);
  const feedbackRecords = useMemo(
    () => records.filter((record) => !isTransferRecord(record)),
    [records],
  );
  const transferRecords = useMemo(
    () => data?.transferRecords ?? [],
    [data],
  );
  const sheets = useMemo(
    () => [allSheets, ...Array.from(new Set(feedbackRecords.map((record) => record.sheet)))],
    [feedbackRecords],
  );
  const factories = useMemo(
    () =>
      Array.from(new Set(feedbackRecords.map((record) => record.factory))).sort((a, b) =>
        a.localeCompare(b, "th"),
      ),
    [feedbackRecords],
  );
  const metrics = useMemo(
    () => [
      ...Array.from(new Set(feedbackRecords.map((record) => record.metric))).sort((a, b) =>
        a.localeCompare(b, "th"),
      ),
    ],
    [feedbackRecords],
  );

  const selectedFactory =
    factory && factories.includes(factory) ? factory : (factories[0] ?? "");
  const weekOptions = useMemo(() => {
    return Array.from(
      new Set(
        feedbackRecords
          .filter((record) => !selectedFactory || record.factory === selectedFactory)
          .flatMap((record) => record.weeks.map((item) => normalizeWeek(item)).filter(Boolean)),
      ),
    ).sort((a, b) => Number(a) - Number(b));
  }, [feedbackRecords, selectedFactory]);
  const selectedWeek = week && weekOptions.includes(week) ? week : (weekOptions[weekOptions.length - 1] ?? "");

  const summaryRows = useMemo(() => {
    return feedbackRecords.filter((record) => {
      const matchesSheet = sheet === allSheets || record.sheet === sheet;
      const matchesFactory = !selectedFactory || record.factory === selectedFactory;
      const matchesMetric =
        metricFilters.length === 0 || metricFilters.includes(record.metric);
      const matchesWeek =
        !selectedWeek || record.weeks.some((item) => normalizeWeek(item) === selectedWeek);

      return matchesSheet && matchesFactory && matchesWeek && matchesMetric;
    });
  }, [feedbackRecords, metricFilters, selectedFactory, selectedWeek, sheet]);

  const filtered = useMemo(() => {
    return summaryRows.filter((record) => {
      const recordStatus = scoreLabel(score(record, feedback[record.id]?.actual ?? ""));
      return statusFilters.length === 0 || statusFilters.includes(recordStatus);
    });
  }, [feedback, statusFilters, summaryRows]);

  const tableRows = filtered;
  const scores = summaryRows
    .map((record) => score(record, feedback[record.id]?.actual ?? ""))
    .filter((value): value is number => value !== null);
  const goodCount = scores.filter((value) => value >= 80).length;
  const warningCount = scores.filter((value) => value >= 60 && value < 80).length;
  const badCount = scores.filter((value) => value < 60).length;
  const closePercent =
    scores.length > 0 ? Math.round((goodCount / scores.length) * 100) : null;
  function updateFeedback(id: string, patch: Partial<Feedback>) {
    setFeedback((current) => ({
      ...current,
      [id]: {
        ...({ actual: "", accuracy: "", comment: "" } satisfies Feedback),
        ...current[id],
        ...patch,
      },
    }));
  }

  function recordUpload(type: UploadHistoryItem["type"], name: string) {
    const uploadedAt = new Intl.DateTimeFormat("th-TH", {
      dateStyle: "medium",
      timeStyle: "short",
    }).format(new Date());
    setUploadHistory((current) => [{ type, name, uploadedAt }, ...current].slice(0, 12));
  }

  async function rebuildBalanceComparison(aiFile: File | null, balanceFile: File, parsedAiData: AiData | null = data) {
    const comparison = aiFile
      ? await parseBalanceComparisonWorkbook(aiFile, balanceFile)
      : parsedAiData
        ? await parseBalanceComparisonFromAiData(parsedAiData, balanceFile)
        : null;
    if (!comparison) {
      throw new Error("กรุณาอัปโหลดไฟล์ผล AI ก่อน แล้วจึงอัปโหลดไฟล์แผน");
    }
    setBalanceData(comparison);
    window.localStorage.setItem(uploadedBalanceComparisonKey, JSON.stringify(comparison));
    const stored = await saveStoredPlanComparison(comparison);
    setActivePlanComparisonId(stored.id);
    setPlanComparisonHistory(await listStoredPlanComparisons());
    return comparison;
  }

  function clearPlanComparison() {
    setBalanceData(null);
    setActivePlanComparisonId("");
    window.localStorage.removeItem(uploadedBalanceComparisonKey);
  }

  async function handlePlanHistoryChange(id: string) {
    setActivePlanComparisonId(id);
    if (!id) return;

    const comparison = await getStoredPlanComparison(id);
    if (!comparison) {
      setUploadStatus({
        tone: "error",
        message: "พบชื่อประวัติ แต่ไม่พบข้อมูลไฟล์ชุดนี้ในเครื่อง กรุณาอัปโหลด AI + แผนชุดนั้นอีกครั้งเพื่อบันทึกข้อมูลย้อนหลัง",
      });
      return;
    }

    setBalanceData(comparison);
    setUploadedNames((current) => ({
      ...current,
      ai: comparison.aiFile,
      balance: comparison.sourceFile,
    }));
    window.localStorage.setItem(uploadedBalanceComparisonKey, JSON.stringify(comparison));
    setUploadStatus({
      tone: "success",
      message: `เปิดประวัติ ${comparison.aiFile} เทียบ ${comparison.sourceFile}`,
    });
  }

  async function handleAiUpload(file: File) {
    const detectedKind = uploadKindFromName(file.name);
    if (detectedKind === "Actual") {
      await handleActualUpload(file);
      return;
    }
    if (detectedKind === "Balance") {
      await handleBalanceUpload(file);
      return;
    }

    setIsUploading(true);
    setUploadStatus(null);

    try {
      const uploadedData = await parseAiWorkbook(file);
      setAiWorkbookFile(file);
      setData(uploadedData);
      setUploadedNames((current) => ({ ...current, ai: file.name }));
      recordUpload("AI", file.name);
      setSheet(allSheets);
      setFactory("");
      window.localStorage.setItem(uploadedDataKey, JSON.stringify(uploadedData));
      const comparison = balanceWorkbookFile
        ? await rebuildBalanceComparison(file, balanceWorkbookFile, uploadedData)
        : null;
      if (!comparison) clearPlanComparison();
      setUploadStatus({
        tone: "success",
        message: comparison
          ? `อัปโหลดสำเร็จ: โหลดผล AI ${uploadedData.records.length.toLocaleString(
              "th-TH",
            )} รายการ และสร้างผลเทียบแผน ${comparison.records.length.toLocaleString("th-TH")} รายการ`
          : `อัปโหลดสำเร็จ: โหลด ${uploadedData.records.length.toLocaleString(
          "th-TH",
        )} รายการจากหัวข้อ feedback ที่ระบบจำไว้ กรุณาอัปโหลดไฟล์แผนอีกครั้งเพื่ออัปเดตแท็บ Feedback โรงงาน / วิเคราะห์ผล / AI vs แผน`,
      });
    } catch (error) {
      setUploadStatus({
        tone: "error",
        message: error instanceof Error ? error.message : "ไม่สามารถอ่านไฟล์นี้ได้",
      });
    } finally {
      setIsUploading(false);
    }
  }

  async function handleActualUpload(file: File) {
    const detectedKind = uploadKindFromName(file.name);
    if (detectedKind === "Balance") {
      await handleBalanceUpload(file);
      return;
    }

    if (!data?.records.length) {
      setUploadStatus({
        tone: "error",
        message: "กรุณาอัปโหลดไฟล์ผล AI ก่อน แล้วจึงอัปโหลดไฟล์ Actual",
      });
      return;
    }

    setIsUploading(true);
    setUploadStatus(null);

    try {
      const [actualFeedback, actualTransferPayload] = await Promise.all([
        parseActualWorkbook(file, data.records),
        parseActualTransferWorkbook(file),
      ]);
      setFeedback((current) => {
        const merged = { ...current };

        for (const [id, actual] of Object.entries(actualFeedback)) {
          merged[id] = {
            ...({ actual: "", accuracy: "", comment: "" } satisfies Feedback),
            ...current[id],
            actual: actual.actual,
          };
        }

        return merged;
      });
      setTransferActuals(actualTransferPayload);
      window.localStorage.setItem(uploadedActualKey, JSON.stringify(actualFeedback));
      window.localStorage.setItem(uploadedActualTransferKey, JSON.stringify(actualTransferPayload));
      setUploadedNames((current) => ({ ...current, actual: file.name }));
      recordUpload("Actual", file.name);
      setUploadStatus({
        tone: "success",
        message: `เติมค่าจริงสำเร็จ: จับคู่ได้ ${Object.keys(actualFeedback).length.toLocaleString(
          "th-TH",
        )} รายการ / ข้อมูลโอน ${Object.keys(actualTransferPayload).length.toLocaleString("th-TH")} กลุ่ม`,
      });
      setActiveTab("feedback");
    } catch (error) {
      setUploadStatus({
        tone: "error",
        message: error instanceof Error ? error.message : "ไม่สามารถอ่านไฟล์ Actual ได้",
      });
    } finally {
      setIsUploading(false);
    }
  }

  async function handleBalanceUpload(file: File) {
    const detectedKind = uploadKindFromName(file.name);
    if (detectedKind === "Actual") {
      await handleActualUpload(file);
      return;
    }

    setIsUploading(true);
    setUploadStatus(null);
    setBalanceWorkbookFile(file);
    setUploadedNames((current) => ({ ...current, balance: file.name }));
    recordUpload("Balance", file.name);

    try {
      const comparison = await rebuildBalanceComparison(aiWorkbookFile, file, data);
      setUploadStatus({
        tone: "success",
        message: `เพิ่มไฟล์แผนแล้ว: ${file.name} และสร้างผลวิเคราะห์ ${comparison.records.length.toLocaleString(
          "th-TH",
        )} รายการ`,
      });
      setActiveTab("analyze");
    } catch (error) {
      setUploadStatus({
        tone: "error",
        message: error instanceof Error ? error.message : "ไม่สามารถอ่านไฟล์แผนได้",
      });
    } finally {
      setIsUploading(false);
    }
  }

  return (
    <main className="app-background min-h-screen text-[#172033]">
      <div className="relative z-10 grid min-h-screen lg:grid-cols-[260px_1fr]">
        <Sidebar activeTab={activeTab} setActiveTab={setActiveTab} />

        <section className="min-w-0">
          <header className="sticky top-0 z-20 border-b border-[#f5b4cf] bg-white/90 backdrop-blur">
            <div className="flex min-h-20 flex-col items-start justify-between gap-3 px-4 py-3 sm:min-h-24 sm:px-8 sm:py-4 xl:flex-row xl:items-center">
              <div className="flex w-full items-center gap-4 xl:w-auto xl:gap-5">
                <button className="grid size-10 shrink-0 place-items-center rounded-md text-slate-600 hover:bg-slate-100 sm:size-11">
                  <Menu size={28} />
                </button>
                <div>
                  <h1 className="text-lg font-bold leading-snug sm:text-2xl">
                    {activeTab === "upload"
                      ? "อัปโหลดผล AI"
                      : activeTab === "planning"
                        ? "วางแผน"
                      : activeTab === "analyze"
                        ? "วิเคราะห์ผล"
                      : activeTab === "balance"
                        ? "เทียบผล AI: แผน"
                      : activeTab === "transfer"
                        ? "เทียบผล AI: การโอนสินค้า"
                        : "บันทึก Feedback: เทียบผล AI กับค่าจริง"}
                  </h1>
                  <p className="mt-1 line-clamp-2 text-sm text-slate-500">
                      {activeTab === "upload"
                        ? "นำเข้าไฟล์ Excel ที่มีโครงสร้างชีตและหัวคอลัมน์แบบเดิม"
                      : activeTab === "planning"
                        ? "ใช้ผล AI เป็นจุดตั้งต้นให้ทีมวางแผนรายโรงงาน"
                      : activeTab === "balance"
                          ? "เทียบผล AI กับไฟล์แผนตามโรงงาน สัปดาห์ และกลุ่มสินค้า"
                        : activeTab === "transfer"
                          ? "ตรวจว่าต้นทาง ปลายทาง และปริมาณที่ AI แนะนำโอนตรงกับ Actual หรือไม่"
                          : "กรอกและตรวจสอบความถูกต้องของผลลัพธ์ AI เทียบกับค่าจริง"}
                  </p>
                </div>
              </div>

              <div className="hidden items-center gap-3 xl:flex">
                <TopButton>
                  31/05/2024
                  <CalendarDays size={17} />
                </TopButton>
                {activeTab === "feedback" ? (
                  <select
                    className="h-11 min-w-48 rounded-md border border-[#dfe6ef] bg-white px-4 text-sm font-medium shadow-sm"
                    value={sheet}
                    onChange={(event) => setSheet(event.target.value)}
                  >
                    {sheets.map((option) => (
                      <option key={option}>{option}</option>
                    ))}
                  </select>
                ) : activeTab === "transfer" ? (
                  <TopButton>ชีท 4-5</TopButton>
                ) : activeTab === "balance" ? (
                  <TopButton>แผน</TopButton>
                ) : null}
                <TopButton>
                  <HelpCircle size={17} />
                  วิธีใช้งาน
                </TopButton>
              </div>
              <MobileTabBar activeTab={activeTab} setActiveTab={setActiveTab} />
            </div>
          </header>

          <div className="space-y-5 px-4 py-5 sm:px-8">
            {activeTab === "planning" ? (
              <PlanningPanel
                key={data ? `${data.sourceFile}|${data.generatedAt}|${data.records.length}` : "no-ai-data"}
                data={data}
                balanceData={balanceData}
                planningData={planningBalanceData}
                feedback={feedback}
                updateFeedback={updateFeedback}
              />
            ) : activeTab === "analyze" ? (
              <>
                <PlanHistorySelector
                  items={planComparisonHistory}
                  value={activePlanComparisonId}
                  onChange={handlePlanHistoryChange}
                />
                <AnalyzeVdpPanel data={balanceData} aiData={data} feedback={feedback} uploadedNames={uploadedNames} />
              </>
            ) : activeTab === "factoryFeedback" ? (
              <>
                <PlanHistorySelector
                  items={planComparisonHistory}
                  value={activePlanComparisonId}
                  onChange={handlePlanHistoryChange}
                />
                <FactoryFeedbackPanel data={balanceData} feedback={feedback} updateFeedback={updateFeedback} />
              </>
            ) : activeTab === "upload" ? (
              <UploadAiPanel
                data={data}
                isUploading={isUploading}
                status={uploadStatus}
                uploadedNames={uploadedNames}
                uploadHistory={uploadHistory}
                onUpload={handleAiUpload}
                onActualUpload={handleActualUpload}
                onBalanceUpload={handleBalanceUpload}
              />
            ) : activeTab === "transfer" ? (
              <TransferComparison
                records={transferRecords}
                actuals={transferActuals}
                feedback={feedback}
                updateFeedback={updateFeedback}
              />
            ) : activeTab === "balance" ? (
              <>
                <PlanHistorySelector
                  items={planComparisonHistory}
                  value={activePlanComparisonId}
                  onChange={handlePlanHistoryChange}
                />
                <BalanceComparison
                  data={balanceData}
                  feedback={feedback}
                  updateFeedback={updateFeedback}
                />
              </>
            ) : (
              <>
                <ContextBar
                  factory={selectedFactory || "โรงงาน A"}
                  factories={factories}
                  onFactoryChange={(nextFactory) => {
                    setFactory(nextFactory);
                    setWeek("");
                  }}
                  week={selectedWeek}
                  weeks={weekOptions}
                  onWeekChange={setWeek}
                  sourceFile={data?.sourceFile ?? "-"}
                />

                <section className="feedback-layout grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_300px]">
                  <ComparisonCard
                    metricFilters={metricFilters}
                    setMetricFilters={setMetricFilters}
                    metrics={metrics}
                    statusFilters={statusFilters}
                    setStatusFilters={setStatusFilters}
                    rows={tableRows}
                    feedback={feedback}
                    updateFeedback={updateFeedback}
                  />
                  <SummaryCard
                    closePercent={closePercent}
                    analyzedCount={scores.length}
                    goodCount={goodCount}
                    warningCount={warningCount}
                    badCount={badCount}
                  />
                </section>

                <CommentCard
                  overall={feedback.__overall?.comment ?? ""}
                  updateFeedback={updateFeedback}
                />
              </>
            )}
          </div>
        </section>
      </div>
    </main>
  );
}

function Sidebar({
  activeTab,
  setActiveTab,
}: {
  activeTab: AppTab;
  setActiveTab: (tab: AppTab) => void;
}) {
  const tabs = navItems.filter((item): item is typeof item & { tab: AppTab } => "tab" in item);

  return (
    <aside className="hidden border-r border-[#e3e8f0] bg-white lg:flex lg:flex-col">
      <div className="flex h-24 items-center gap-3 border-b border-[#edf1f6] px-6">
        <Image
          src="/logo.png"
          alt="PigFactory AI"
          width={56}
          height={56}
          className="size-14 shrink-0 rounded-md object-contain"
        />
        <div>
          <h1 className="text-xl font-bold tracking-tight">PigFactory AI</h1>
          <p className="text-sm text-slate-500">AI Feedback System</p>
        </div>
      </div>

      <nav className="space-y-2 px-3 py-5">
        {tabs.map((item) => {
          const Icon = item.icon;
          const active = item.tab === activeTab;

          return (
            <button
              key={item.label}
              onClick={() => {
                setActiveTab(item.tab);
              }}
              className={`flex h-12 w-full items-center gap-4 rounded-md px-4 text-left text-sm font-medium transition ${
                active ? "bg-[#ffe8f1] text-[#ee3f95]" : "text-slate-600 hover:bg-slate-50"
              }`}
            >
              <Icon size={20} />
              {item.label}
            </button>
          );
        })}
      </nav>

      <div className="mt-auto border-t border-[#edf1f6] p-4">
        <div className="mb-4 flex items-center gap-3 rounded-lg bg-white p-3 shadow-sm ring-1 ring-[#e4e9f1]">
          <div className="grid size-11 place-items-center rounded-full bg-slate-100">
            <User size={22} />
          </div>
          <div>
            <p className="font-semibold">Admin</p>
            <p className="text-xs text-slate-500">Administrator</p>
          </div>
          <ChevronDown className="ml-auto text-slate-400" size={18} />
        </div>
        <button className="flex h-10 w-full items-center gap-2 text-sm font-medium text-[#ef3e8f]">
          <LogOut size={18} />
          ออกจากระบบ
        </button>
      </div>
    </aside>
  );
}

function MobileTabBar({
  activeTab,
  setActiveTab,
}: {
  activeTab: AppTab;
  setActiveTab: (tab: AppTab) => void;
}) {
  const tabs = navItems.filter((item): item is typeof item & { tab: AppTab } => "tab" in item);

  return (
    <div className="flex w-full gap-2 overflow-x-auto pb-1 lg:hidden">
      {tabs.map((item) => {
        const Icon = item.icon;
        const active =
          item.tab === activeTab;

        return (
          <button
            key={item.label}
            className={`flex h-10 shrink-0 items-center gap-2 rounded-md border px-3 text-sm font-bold shadow-sm ${
              active
                ? "border-[#ef3e8f] bg-[#ffe8f1] text-[#ef3e8f]"
                : "border-[#f5b4cf] bg-white/90 text-slate-600"
            }`}
            type="button"
            onClick={() => setActiveTab(item.tab)}
          >
            <Icon size={17} />
            {item.label}
          </button>
        );
      })}
    </div>
  );
}

function TopButton({ children }: { children: ReactNode }) {
  return (
    <button className="flex h-11 items-center gap-3 rounded-md border border-[#dfe6ef] bg-white px-4 text-sm font-medium shadow-sm">
      {children}
    </button>
  );
}

function ContextBar({
  factory,
  factories,
  onFactoryChange,
  week,
  weeks,
  onWeekChange,
  sourceFile,
}: {
  factory: string;
  factories: string[];
  onFactoryChange: (factory: string) => void;
  week: string;
  weeks: string[];
  onWeekChange: (week: string) => void;
  sourceFile: string;
}) {
  return (
    <section className="rounded-xl border border-[#e3e8f0] bg-white p-5 shadow-sm">
      <div className="grid gap-4 lg:grid-cols-[minmax(320px,1.6fr)_0.7fr_0.7fr_1.5fr_1.2fr]">
        <FactorySelectTile
          factories={factories}
          value={factory}
          onChange={onFactoryChange}
        />
        <InfoTile label="วันที่บันทึก" value="31/05/2024" />
        <WeekSelectTile weeks={weeks} value={week} onChange={onWeekChange} />
        <InfoTile label="ช่วงเวลาการทำนาย" value="30/05/2024 20:00 - 31/05/2024 20:00" />
        <InfoTile label="ไฟล์ผล AI" value={sourceFile} />
      </div>
    </section>
  );
}

function FactorySelectTile({
  factories,
  value,
  onChange,
}: {
  factories: string[];
  value: string;
  onChange: (factory: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const filteredFactories = factories.filter((option) =>
    option.toLowerCase().includes(search.trim().toLowerCase()),
  );

  return (
    <div className="relative flex min-w-0 items-center gap-3 border-r border-[#edf1f6] pr-4 last:border-r-0">
      <div className="grid size-11 shrink-0 place-items-center rounded-full bg-slate-100 text-slate-600">
        <Factory size={28} />
      </div>
      <div className="min-w-0 flex-1">
        <p className="text-xs text-slate-500">โรงงาน</p>
        <button
          className="mt-1 flex h-10 w-full items-center justify-between gap-2 rounded-md border border-[#dfe6ef] bg-white px-3 text-left font-semibold outline-none focus:border-[#ef4b98]"
          type="button"
          onClick={() => setOpen((current) => !current)}
        >
          <span className="min-w-0 truncate">{value || "กำลังโหลดโรงเรือน..."}</span>
          <ChevronDown size={16} className="shrink-0 text-slate-400" />
        </button>
        {open ? (
          <div className="absolute left-14 right-4 top-[72px] z-40 rounded-md border border-[#dfe6ef] bg-white p-3 shadow-lg">
            <label className="relative mb-3 block">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
              <input
                className="h-10 w-full rounded-md border border-[#dfe6ef] pl-9 pr-3 text-sm outline-none focus:border-[#ef4b98]"
                placeholder="ค้นหาโรงงาน..."
                value={search}
                onChange={(event) => setSearch(event.target.value)}
              />
            </label>
            <div className="max-h-72 space-y-1 overflow-auto pr-1">
              {filteredFactories.map((option) => (
                <button
                  key={option}
                  className={`w-full rounded-md px-3 py-2 text-left text-sm font-medium hover:bg-[#fff1f7] ${
                    option === value ? "bg-[#ffe8f1] text-[#ef3e8f]" : "text-slate-700"
                  }`}
                  type="button"
                  onClick={() => {
                    onChange(option);
                    setOpen(false);
                    setSearch("");
                  }}
                >
                  {option}
                </button>
              ))}
              {filteredFactories.length === 0 ? (
                <p className="py-3 text-center text-sm text-slate-500">ไม่พบโรงงาน</p>
              ) : null}
            </div>
          </div>
        ) : null}
      </div>
    </div>
  );
}

function WeekSelectTile({
  weeks,
  value,
  onChange,
}: {
  weeks: string[];
  value: string;
  onChange: (week: string) => void;
}) {
  return (
    <div className="flex min-w-0 items-center gap-3 border-r border-[#edf1f6] pr-4 last:border-r-0">
      <div className="min-w-0 flex-1">
        <p className="text-xs text-slate-500">สัปดาห์</p>
        <select
          className="mt-1 h-10 w-full rounded-md border border-[#dfe6ef] bg-white px-3 font-semibold outline-none focus:border-[#ef4b98]"
          value={value}
          onChange={(event) => onChange(event.target.value)}
        >
          {weeks.length === 0 ? <option value="">รอข้อมูลสัปดาห์</option> : null}
          {weeks.map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
      </div>
    </div>
  );
}

function InfoTile({
  icon,
  label,
  value,
}: {
  icon?: ReactNode;
  label: string;
  value: string;
}) {
  return (
    <div className="flex min-w-0 items-center gap-3 border-r border-[#edf1f6] pr-4 last:border-r-0">
      {icon && <div className="grid size-11 shrink-0 place-items-center rounded-full bg-slate-100 text-slate-600">{icon}</div>}
      <div className="min-w-0">
        <p className="text-xs text-slate-500">{label}</p>
        <p className="mt-1 break-words font-semibold">{value}</p>
      </div>
    </div>
  );
}

function ComparisonCard({
  metricFilters,
  setMetricFilters,
  metrics,
  statusFilters,
  setStatusFilters,
  rows,
  feedback,
  updateFeedback,
}: {
  metricFilters: string[];
  setMetricFilters: (metrics: string[]) => void;
  metrics: string[];
  statusFilters: string[];
  setStatusFilters: (statuses: string[]) => void;
  rows: AiRecord[];
  feedback: Record<string, Feedback>;
  updateFeedback: (id: string, patch: Partial<Feedback>) => void;
}) {
  const [savedAt, setSavedAt] = useState<string | null>(null);

  function saveTableFeedback() {
    window.localStorage.setItem(storageKey, JSON.stringify(feedback));
    setSavedAt(
      new Intl.DateTimeFormat("th-TH", {
        hour: "2-digit",
        minute: "2-digit",
      }).format(new Date()),
    );
  }

  return (
    <div className="comparison-card min-w-0 rounded-xl border border-[#e3e8f0] bg-white p-5 shadow-sm">
      <div className="mb-5 flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
        <div className="flex items-center gap-3">
          <span className="grid size-8 place-items-center rounded-full bg-[#ef3e8f] text-sm font-bold text-white">
            1
          </span>
          <div>
            <h3 className="text-lg font-bold">ตรวจสอบและเปรียบเทียบผลลัพธ์</h3>
            <p className="text-sm text-slate-500">ตารางเปรียบเทียบผลลัพธ์ AI กับค่าจริง</p>
          </div>
        </div>
        <div className="flex flex-wrap items-center gap-4 text-xs font-medium">
          <span className="rounded-md bg-[#ffe8f1] px-3 py-1 font-bold text-[#ef3e8f]">
            แสดง {rows.length.toLocaleString("th-TH")} รายการ
          </span>
          <Legend color="bg-emerald-500" label="ดี" />
          <Legend color="bg-orange-500" label="ต่างปานกลาง" />
          <Legend color="bg-red-500" label="ต่างกันมาก" />
        </div>
      </div>

      <div className="mb-4 grid gap-3 lg:grid-cols-[minmax(0,1fr)_260px]">
        <MetricFilterBox
          metrics={metrics}
          values={metricFilters}
          onChange={setMetricFilters}
        />
        <StatusFilterBox values={statusFilters} onChange={setStatusFilters} />
      </div>

        <div className="mobile-table-frame min-w-0 rounded-lg border border-[#dfe6ef]">
          <div className="mobile-table-scroll max-h-[640px] overflow-auto">
          <table className="feedback-table w-full min-w-[1080px] border-collapse text-sm">
            <thead className="bg-[#f8fafc] text-xs font-bold text-slate-600">
              <tr>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-left">ตัวชี้วัด</th>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-right">AI ทำนาย</th>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-right">ค่าจริง (Actual)</th>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-right">ผลต่าง</th>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-right">% ต่างกัน</th>
                <th className="border-r border-[#e3e8f0] px-4 py-3">สถานะ</th>
                <th className="px-4 py-3">ความคิดเห็น (กรอก)</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((record) => {
                const itemFeedback = feedback[record.id] ?? {
                  actual: "",
                  accuracy: "",
                  comment: "",
                };
                const matchScore = score(record, itemFeedback.actual);
                const diff = difference(record, itemFeedback.actual);
                const diffPercent = matchScore === null ? null : 100 - matchScore;
                const commentRequired = matchScore !== null && matchScore < 80;
                const missingRequiredComment =
                  commentRequired && !itemFeedback.comment.trim();

                return (
                  <tr key={record.id} className="border-t border-[#e8edf4]">
                    <td className="border-r border-[#e8edf4] px-4 py-3">
                      <div className="flex items-center gap-3">
                        <div className="grid size-8 shrink-0 place-items-center rounded-md bg-slate-100 text-slate-500">
                          <Gauge size={18} />
                        </div>
                        <div className="min-w-0">
                          <p className="truncate font-semibold leading-tight">{record.metric}</p>
                        </div>
                      </div>
                    </td>
                    <td className="border-r border-[#e8edf4] px-4 py-3 text-right font-mono">
                      {record.kind === "number"
                        ? formatNumber(record.aiValue)
                        : record.examples.join(", ") || "-"}
                    </td>
                    <td className="border-r border-[#e8edf4] px-4 py-3">
                      <input
                        className="ml-auto h-9 w-32 rounded-md bg-white px-3 text-right font-mono outline-none ring-1 ring-[#dfe6ef] focus:ring-[#ef4b98]"
                        value={itemFeedback.actual}
                        onChange={(event) =>
                          updateFeedback(record.id, { actual: event.target.value })
                        }
                      />
                    </td>
                    <td
                      className={`border-r border-[#e8edf4] px-4 py-3 text-right font-mono ${
                        diff !== null && diff > 0 ? "text-red-600" : "text-emerald-600"
                      }`}
                    >
                      {diff === null ? "-" : `${diff > 0 ? "+" : ""}${formatNumber(diff)}`}
                    </td>
                    <td
                      className={`border-r border-[#e8edf4] px-4 py-3 text-right font-mono ${
                        diffPercent !== null && diffPercent > 20
                          ? "text-red-600"
                          : "text-emerald-600"
                      }`}
                    >
                      {diffPercent === null ? "-" : `${diffPercent > 0 ? "+" : ""}${diffPercent}%`}
                    </td>
                    <td className="border-r border-[#e8edf4] px-4 py-3 text-center">
                      <span
                        className={`inline-flex h-8 min-w-24 items-center justify-center rounded-md px-3 text-xs font-bold ${scoreTone(
                          matchScore,
                        )}`}
                      >
                        {scoreLabel(matchScore)}
                      </span>
                    </td>
                    <td className="px-4 py-3">
                      <div className="relative">
                        <input
                          className={`h-9 w-full rounded-md border px-3 pr-9 text-sm outline-none ${
                            missingRequiredComment
                              ? "border-red-400 bg-red-50 focus:border-red-500"
                              : "border-[#dfe6ef] focus:border-[#ef4b98]"
                          }`}
                          placeholder={
                            commentRequired
                              ? "ต้องกรอกความคิดเห็น..."
                              : "กรอกความคิดเห็น..."
                          }
                          value={itemFeedback.comment}
                          onChange={(event) =>
                            updateFeedback(record.id, { comment: event.target.value })
                          }
                        />
                        {commentRequired ? (
                          <span className="absolute right-2 top-1/2 -translate-y-1/2 text-lg font-bold text-red-500">
                            *
                          </span>
                        ) : (
                          <MessageCircle
                            className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400"
                            size={17}
                          />
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
      <div className="mt-4 flex flex-col gap-3 border-t border-[#f5b4cf] pt-4 sm:flex-row sm:items-center sm:justify-between">
        <p className="text-xs text-slate-500">* % ต่างกัน = |ผลต่าง| / AI ทำนาย × 100</p>
        <div className="flex flex-col gap-2 sm:items-end">
          <button
            className="inline-flex h-11 items-center justify-center gap-2 rounded-md bg-[#ef3e8f] px-5 text-sm font-bold text-white shadow-sm hover:bg-[#dc2e81]"
            type="button"
            onClick={saveTableFeedback}
          >
            <Save size={18} />
            บันทึกความเห็น
          </button>
          {savedAt ? (
            <p className="text-xs font-bold text-emerald-600">บันทึกแล้ว {savedAt}</p>
          ) : null}
        </div>
      </div>
    </div>
  );
}

function Legend({ color, label }: { color: string; label: string }) {
  return (
    <span className="flex items-center gap-2">
      <span className={`size-2 rounded-full ${color}`} />
      {label}
    </span>
  );
}

function MetricFilterBox({
  metrics,
  values,
  onChange,
}: {
  metrics: string[];
  values: string[];
  onChange: (values: string[]) => void;
}) {
  const [search, setSearch] = useState("");
  const filteredMetrics = metrics.filter((option) =>
    option.toLowerCase().includes(search.trim().toLowerCase()),
  );

  function toggleMetric(option: string) {
    onChange(
      values.includes(option)
        ? values.filter((value) => value !== option)
        : [...values, option],
    );
  }

  return (
    <details className="relative rounded-md border border-[#dfe6ef] bg-white">
      <summary className="flex h-11 cursor-pointer list-none items-center justify-between gap-2 px-3 text-sm font-medium marker:hidden">
        <span className="min-w-0 truncate">
          ตัวแปร:{" "}
          {values.length === 0
            ? "ทั้งหมด"
            : values.length === 1
              ? values[0]
              : `${values.length} รายการ`}
        </span>
        <ChevronDown size={16} className="shrink-0 text-slate-400" />
      </summary>
      <div className="absolute left-0 top-12 z-30 w-full rounded-md border border-[#dfe6ef] bg-white p-3 shadow-lg">
        <div className="mb-2 flex items-center justify-between gap-2">
          <p className="text-xs font-bold text-slate-600">ตัวแปร</p>
          <button
            className="text-xs font-bold text-[#ef3e8f]"
            type="button"
            onClick={() => onChange([])}
          >
            ทั้งหมด
          </button>
        </div>
        <label className="relative mb-3 block">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
          <input
            className="h-10 w-full rounded-md border border-[#dfe6ef] pl-9 pr-3 text-sm outline-none focus:border-[#ef4b98]"
            placeholder="ค้นหาและเลือกหลายตัวแปร..."
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>
        <FilterGroup
          options={filteredMetrics}
          values={values}
          onToggle={toggleMetric}
          emptyText="ไม่พบตัวแปร"
        />
      </div>
    </details>
  );
}

function StatusFilterBox({
  values,
  onChange,
}: {
  values: string[];
  onChange: (values: string[]) => void;
}) {
  const statuses = ["ดี", "ต่างปานกลาง", "ต่างกันมาก", "รอข้อมูล"];

  function toggleStatus(option: string) {
    onChange(
      values.includes(option)
        ? values.filter((value) => value !== option)
        : [...values, option],
    );
  }

  return (
    <details className="relative rounded-md border border-[#dfe6ef] bg-white">
      <summary className="flex h-11 cursor-pointer list-none items-center justify-between gap-2 px-3 text-sm font-medium marker:hidden">
        <span className="min-w-0 truncate">
          สถานะ:{" "}
          {values.length === 0
            ? "ทั้งหมด"
            : values.length === 1
              ? values[0]
              : `${values.length} รายการ`}
        </span>
        <ChevronDown size={16} className="shrink-0 text-slate-400" />
      </summary>
      <div className="absolute right-0 top-12 z-30 w-full rounded-md border border-[#dfe6ef] bg-white p-3 shadow-lg">
        <div className="mb-2 flex items-center justify-between gap-2">
          <p className="text-xs font-bold text-slate-600">สถานะ</p>
          <button
            className="text-xs font-bold text-[#ef3e8f]"
            type="button"
            onClick={() => onChange([])}
          >
            ทั้งหมด
          </button>
        </div>
        <FilterGroup
          options={statuses}
          values={values}
          onToggle={toggleStatus}
          emptyText="ไม่พบสถานะ"
        />
      </div>
    </details>
  );
}

function FilterGroup({
  options,
  values,
  onToggle,
  emptyText,
}: {
  options: string[];
  values: string[];
  onToggle: (value: string) => void;
  emptyText: string;
}) {
  return (
    <div className="max-h-64 overflow-auto pr-1">
      <div className="space-y-2">
        {options.map((option) => (
          <label key={option} className="flex items-start gap-2 text-sm">
            <input
              type="checkbox"
              className="mt-1 size-4 accent-[#ef3e8f]"
              checked={values.includes(option)}
              onChange={() => onToggle(option)}
            />
            <span className="leading-5">{option}</span>
          </label>
        ))}
        {options.length === 0 && (
          <p className="py-2 text-center text-sm text-slate-500">{emptyText}</p>
        )}
      </div>
    </div>
  );
}

function SummaryCard({
  closePercent,
  analyzedCount,
  goodCount,
  warningCount,
  badCount,
}: {
  closePercent: number | null;
  analyzedCount: number;
  goodCount: number;
  warningCount: number;
  badCount: number;
}) {
  const percent = closePercent ?? 0;
  const range =
    closePercent === null
      ? { color: "#94a3b8", soft: "#f1f5f9", label: "รอข้อมูล" }
      : closePercent >= 80
        ? { color: "#39b87f", soft: "#ecfdf5", label: "ดี" }
        : closePercent >= 60
        ? { color: "#f59e0b", soft: "#fff7ed", label: "ต่างปานกลาง" }
          : { color: "#e11d48", soft: "#fff1f2", label: "ต่างกันมาก" };

  return (
    <div className="summary-card min-w-0 space-y-5">
      <div className="rounded-xl border border-[#e3e8f0] bg-white p-5 shadow-sm">
        <h3 className="text-lg font-bold">สรุปภาพรวมรอบการผลิตนี้</h3>
        <div className="mt-6 flex justify-center">
          <div
            className="relative grid size-40 place-items-center rounded-full"
            style={{
              background: `conic-gradient(${range.color} 0 ${percent}%, #e8edf4 ${percent}% 100%)`,
            }}
          >
            <div className="grid size-28 place-items-center rounded-full bg-white shadow-inner">
              <div className="text-center">
                <p className="text-3xl font-bold">{closePercent === null ? "-" : `${closePercent}%`}</p>
                <p className="text-sm font-bold" style={{ color: range.color }}>
                  {range.label}
                </p>
              </div>
            </div>
          </div>
        </div>
        <div
          className="mt-4 rounded-md border px-3 py-2 text-center text-sm font-bold"
          style={{ borderColor: range.color, backgroundColor: range.soft, color: range.color }}
        >
          ใกล้เคียง {goodCount.toLocaleString("th-TH")} / {analyzedCount.toLocaleString("th-TH")} ตัวชี้วัด
        </div>
        <div className="mt-6 space-y-3">
          <SummaryPill tone="good" label="ดี (ต่างกันน้อย)" value={`${goodCount.toLocaleString("th-TH")} ตัวชี้วัด`} />
          <SummaryPill tone="warn" label="ต่างปานกลาง" value={`${warningCount.toLocaleString("th-TH")} ตัวชี้วัด`} />
          <SummaryPill tone="bad" label="ต่างกันมาก" value={`${badCount.toLocaleString("th-TH")} ตัวชี้วัด`} />
        </div>
      </div>

      <div className="rounded-xl border border-blue-200 bg-blue-50 p-5 text-sm text-blue-800 shadow-sm">
        <div className="mb-3 flex items-center gap-2 font-bold">
          <Target size={18} />
          คำแนะนำ
        </div>
        <ul className="space-y-2 pl-4">
          <li>หากค่าจริงต่างจาก AI มาก โปรดใส่ความคิดเห็นเพื่อช่วยพัฒนาโมเดล</li>
          <li>ความคิดเห็นควรครอบคลุมปัจจัยที่มีผลต่อความคลาดเคลื่อน</li>
        </ul>
      </div>
    </div>
  );
}

function SummaryPill({
  tone,
  label,
  value,
}: {
  tone: "good" | "warn" | "bad";
  label: string;
  value: string;
}) {
  const classes = {
    good: "border-emerald-200 bg-emerald-50 text-emerald-700",
    warn: "border-orange-200 bg-orange-50 text-orange-700",
    bad: "border-rose-200 bg-rose-50 text-rose-700",
  };

  return (
    <div className={`flex items-center justify-between rounded-md border px-4 py-3 text-sm font-bold ${classes[tone]}`}>
      <span>{label}</span>
      <span>{value}</span>
    </div>
  );
}

function CommentCard({
  overall,
  updateFeedback,
}: {
  overall: string;
  updateFeedback: (id: string, patch: Partial<Feedback>) => void;
}) {
  return (
    <section className="rounded-xl border border-[#e3e8f0] bg-white p-5 shadow-sm">
      <div className="mb-5 flex items-center gap-3">
        <span className="grid size-8 place-items-center rounded-full bg-[#31415e] text-sm font-bold text-white">
          2
        </span>
        <h3 className="text-lg font-bold">กรอกความคิดเห็นและบันทึก</h3>
      </div>
      <div className="grid gap-5 xl:grid-cols-[1fr_1fr_300px]">
        <label className="block text-sm font-bold">
          ความคิดเห็นโดยรวมต่อผลลัพธ์ AI ในรอบการผลิตนี้
          <textarea
            className="mt-3 min-h-36 w-full resize-y rounded-md border border-[#dfe6ef] px-4 py-3 font-normal outline-none focus:border-[#ef4b98]"
            placeholder="เช่น สภาพอากาศเปลี่ยนแปลง, มีการปรับสูตรอาหาร, อุปกรณ์บางตัวมีปัญหา ฯลฯ"
            value={overall}
            onChange={(event) => updateFeedback("__overall", { comment: event.target.value })}
          />
          <span className="mt-1 block text-right text-xs font-normal text-slate-400">
            {overall.length} / 500
          </span>
        </label>

        <div className="rounded-lg border border-[#dfe6ef] p-4">
          <p className="mb-3 text-sm font-bold">ปัจจัยที่อาจส่งผลต่อความคลาดเคลื่อน</p>
          <div className="grid gap-3 text-sm sm:grid-cols-2">
            {[
              "สภาพอากาศเปลี่ยนแปลง",
              "การจัดการภายในโรงเรือน",
              "การปรับสูตรอาหาร",
              "ความหนาแน่นของสัตว์",
              "อื่นๆ (โปรดระบุ)",
            ].map((label) => (
              <label key={label} className="flex items-center gap-2">
                <input type="checkbox" className="size-4 accent-[#ef3e8f]" />
                {label}
              </label>
            ))}
          </div>
          <input
            className="mt-4 h-10 w-full rounded-md border border-[#dfe6ef] px-3 text-sm outline-none focus:border-[#ef4b98]"
            placeholder="ระบุเพิ่มเติม..."
          />
        </div>

        <div className="rounded-lg border border-dashed border-[#ff9ac3] bg-[#fff7fb] p-5 text-center">
          <Paperclip className="mx-auto text-slate-500" size={30} />
          <p className="mt-3 text-sm font-bold">แนบไฟล์เพิ่มเติม (ถ้ามี)</p>
          <p className="mt-2 text-xs text-slate-500">ลากไฟล์มาวางที่นี่ หรือ</p>
          <button className="mt-4 inline-flex h-9 items-center gap-2 rounded-md bg-[#ef3e8f] px-4 text-sm font-bold text-white">
            <Upload size={16} />
            เลือกไฟล์
          </button>
        </div>
      </div>
      <div className="mt-6 flex justify-end gap-3 border-t border-[#edf1f6] pt-5">
        <button className="h-12 rounded-md border border-[#dfe6ef] bg-white px-8 font-bold text-slate-700">
          ยกเลิก
        </button>
        <button className="flex h-12 items-center gap-2 rounded-md bg-[#ef3e8f] px-8 font-bold text-white shadow-sm">
          <Save size={18} />
          บันทึก Feedback
        </button>
      </div>
    </section>
  );
}

function PlanningPanel({
  data,
  balanceData,
  planningData,
  feedback,
  updateFeedback,
}: {
  data: AiData | null;
  balanceData: BalanceData | null;
  planningData: PlanningBalanceData | null;
  feedback: Record<string, Feedback>;
  updateFeedback: (id: string, patch: Partial<Feedback>) => void;
}) {
  const templateRows = useMemo(() => planningData?.records ?? [], [planningData]);
  const aiPlanRows = useMemo(() => dedupeBalanceRows(balanceData?.records ?? []), [balanceData]);
  const aiRecords = useMemo(() => (data?.records ?? []).filter((record) => !isTransferRecord(record)), [data]);
  const [weekFilters, setWeekFilters] = useState<string[]>([]);
  const [factoryFilter, setFactoryFilter] = useState("");
  const [productFilters, setProductFilters] = useState<string[]>([]);
  const [statusFilters, setStatusFilters] = useState<string[]>([]);
  const [savedCard, setSavedCard] = useState<{ id: string; time: string } | null>(null);

  const weeks = useMemo(
    () =>
      Array.from(
        new Set(
          [
            ...aiRecords.flatMap((record) => record.weeks),
            ...aiPlanRows.map((row) => row.week),
            ...templateRows.map((row) => row.week),
          ]
            .map((item) => normalizeWeek(item))
            .filter(Boolean),
        ),
      )
        .sort((a, b) => Number(b) - Number(a)),
    [aiPlanRows, aiRecords, templateRows],
  );
  const activeWeekFilters =
    weekFilters.length > 0 && weekFilters.every((item) => weeks.includes(item))
      ? weekFilters
      : weeks[0]
        ? [weeks[0]]
        : [];
  const factories = useMemo(
    () =>
      Array.from(
        new Set(
          [
            ...aiPlanRows.map((record) => record.factory),
            ...templateRows.map((record) => record.factory),
            ...aiRecords.map((record) => record.factory),
          ].filter(Boolean),
        ),
      ).sort((a, b) => a.localeCompare(b, "th")),
    [aiPlanRows, aiRecords, templateRows],
  );
  const selectedFactory = factoryFilter && factories.includes(factoryFilter) ? factoryFilter : (factories[0] ?? "");
  const products = useMemo(
    () =>
      Array.from(
        new Set(
          [
            ...templateRows.map((record) => record.product),
            ...aiPlanRows.map((record) => record.productGroup).filter(Boolean),
            ...(aiRecords.length > 0 ? ["รวมจากไฟล์ผล AI"] : []),
          ].filter(Boolean),
        ),
      ).sort((a, b) => a.localeCompare(b, "th")),
    [aiPlanRows, aiRecords, templateRows],
  );
  const selectedTemplateRows = templateRows.filter(
    (record) =>
      (!selectedFactory || record.factory === selectedFactory) &&
      (activeWeekFilters.length === 0 || activeWeekFilters.includes(normalizeWeek(record.week))),
  );
  const selectedAiRows = aiPlanRows.filter(
    (record) =>
      record.tableType === "product-group" &&
      (!selectedFactory || record.factory === selectedFactory) &&
      (activeWeekFilters.length === 0 || activeWeekFilters.includes(normalizeWeek(record.week))),
  );
  const rawAiRowsForFactory = aiRecords.filter((record) => !selectedFactory || record.factory === selectedFactory);
  const weekMatchedRawAiRows = rawAiRowsForFactory.filter(
    (record) =>
      activeWeekFilters.length === 0 ||
      record.weeks.length === 0 ||
      record.weeks.some((week) => activeWeekFilters.includes(normalizeWeek(week))),
  );
  const selectedRawAiRows = weekMatchedRawAiRows.length > 0 ? weekMatchedRawAiRows : rawAiRowsForFactory;
  const planningRows = buildUnifiedPlanningRows(selectedTemplateRows, selectedAiRows, selectedRawAiRows).filter((row) => {
    const matchesProduct = productFilters.length === 0 || productFilters.includes(row.product);
    const status = unifiedPlanningStatus(row.aiShortageSurplus);
    const matchesStatus = statusFilters.length === 0 || statusFilters.includes(status);
    return matchesProduct && matchesStatus;
  });
  const totalAiProduction = planningRows.reduce((sum, row) => sum + (row.aiProduction ?? 0), 0);
  const totalAiShortage = planningRows.reduce((sum, row) => sum + Math.min(row.aiShortageSurplus ?? 0, 0), 0);
  const totalAiSurplus = planningRows.reduce((sum, row) => sum + Math.max(row.aiShortageSurplus ?? 0, 0), 0);

  function savePlanCard(id: string) {
    window.localStorage.setItem(storageKey, JSON.stringify(feedback));
    setSavedCard({
      id,
      time: new Intl.DateTimeFormat("th-TH", {
        hour: "2-digit",
        minute: "2-digit",
      }).format(new Date()),
    });
  }

  return (
    <section className="space-y-5">
      <div className="rounded-xl border border-[#f5b4cf] bg-white/95 p-5 shadow-sm">
        <div className="mb-4 flex flex-col gap-2 lg:flex-row lg:items-start lg:justify-between">
          <div>
            <p className="text-sm font-bold text-[#ef3e8f]">AI Planning Workspace</p>
            <h2 className="mt-1 text-xl font-bold">วางแผนจากผล AI รายโรงงาน</h2>
            <p className="mt-1 text-sm text-slate-500">
              ใช้ค่า AI เป็นจุดตั้งต้นให้ทีมวางแผนโรงงานเลือกแนวทาง ปรับตัวเลข และบันทึกเหตุผล
            </p>
          </div>
          <span className="rounded-md bg-[#ffe8f1] px-3 py-1 text-sm font-bold text-[#ef3e8f]">
            {selectedFactory || "ยังไม่พบโรงงาน"}
          </span>
        </div>
        <div className="grid gap-3 xl:grid-cols-4">
          <TransferFilterBox label="สัปดาห์" options={weeks} values={activeWeekFilters} onChange={setWeekFilters} />
          <label className="block">
            <span className="mb-1 block text-xs font-bold text-slate-500">โรงงาน</span>
            <select
              className="h-11 w-full rounded-md border border-[#dfe6ef] bg-white px-3 text-sm font-bold text-slate-700 outline-none focus:border-[#ef3e8f]"
              value={selectedFactory}
              onChange={(event) => setFactoryFilter(event.target.value)}
            >
              {factories.map((factoryName) => (
                <option key={factoryName} value={factoryName}>
                  {factoryName}
                </option>
              ))}
            </select>
          </label>
          <TransferFilterBox label="ชิ้นส่วน/กลุ่มสินค้า" options={products} values={productFilters} onChange={setProductFilters} />
          <TransferFilterBox
            label="สถานะ"
            options={["ขาด", "เหลือ", "สมดุล", "รอข้อมูล"]}
            values={statusFilters}
            onChange={setStatusFilters}
          />
        </div>
      </div>

      <div className="grid gap-5">
        {selectedFactory ? (() => {
          const feedbackId = `planning|${selectedFactory}|${activeWeekFilters.join(",") || "all"}`;
          const currentFeedback = feedback[feedbackId] ?? { actual: "", accuracy: "", comment: "" };

          return (
            <div key={selectedFactory} className="rounded-xl border border-[#e3e8f0] bg-white/95 p-5 shadow-sm">
              <div className="mb-4 flex flex-col gap-2 lg:flex-row lg:items-start lg:justify-between">
                <div>
                  <h3 className="text-lg font-bold">{selectedFactory}</h3>
                  <p className="text-sm text-slate-500">
                    สัปดาห์ {activeWeekFilters.join(", ") || "-"} · {planningRows.length.toLocaleString("th-TH")} แถวใน template กลาง
                  </p>
                </div>
                <span className="rounded-md bg-slate-100 px-3 py-1 text-sm font-bold text-slate-600">
                  Template กลางทุกโรงงาน
                </span>
              </div>

              <div className="mb-4 grid gap-3 md:grid-cols-3">
                <PlanningSummaryTile label="AI ผลิตรวม" value={`${formatCompact(totalAiProduction)} kg`} />
                <PlanningSummaryTile label="AI ของขาด" value={`${formatCompact(Math.abs(totalAiShortage))} kg`} tone="red" />
                <PlanningSummaryTile label="AI ของเหลือ" value={`${formatCompact(totalAiSurplus)} kg`} tone="green" />
              </div>

              <div className="mobile-table-frame rounded-lg border border-[#dfe6ef]">
                <div className="mobile-table-scroll max-h-[620px] overflow-auto">
                  <table className="w-full min-w-[1480px] border-collapse text-sm">
                    <thead className="bg-[#f8fafc] text-xs font-bold text-slate-600">
                      <tr>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-left">ชิ้นส่วน/กลุ่มสินค้า</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-right">%Yield</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-right">AI ผลิต</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-right">AI ยกมา</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-right">AI รับโอน</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-right">AI โอนออก</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-right">AI Total Supply</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-right">AI FC</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-right">AI QT</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-right">AI ขาด/เหลือ</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-center">สถานะ</th>
                        <th className="border-r border-[#e3e8f0] px-3 py-3 text-left">แผนที่จะใช้</th>
                        <th className="px-3 py-3 text-left">หมายเหตุ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {planningRows.map((row) => {
                        const rowFeedbackId = `planning-row|${selectedFactory}|${activeWeekFilters.join(",")}|${row.product}`;
                        const rowFeedback = feedback[rowFeedbackId] ?? { actual: "", accuracy: "", comment: "" };
                        const status = unifiedPlanningStatus(row.aiShortageSurplus);
                        return (
                          <tr key={rowFeedbackId} className="border-t border-[#e8edf4]">
                            <td className="border-r border-[#e8edf4] px-3 py-3 font-medium">
                              <p>{row.product}</p>
                              {row.sap ? <p className="mt-1 text-xs text-slate-400">SAP {row.sap}</p> : null}
                            </td>
                            <td className="border-r border-[#e8edf4] px-3 py-3 text-right font-mono">{formatPlanningPercent(row.yieldFg)}</td>
                            <td className="border-r border-[#e8edf4] px-3 py-3 text-right font-mono">{formatNumber(row.aiProduction)}</td>
                            <td className="border-r border-[#e8edf4] px-3 py-3 text-right font-mono">{formatNumber(row.aiStock)}</td>
                            <td className="border-r border-[#e8edf4] px-3 py-3 text-right font-mono">{formatNumber(row.aiTransferIn)}</td>
                            <td className="border-r border-[#e8edf4] px-3 py-3 text-right font-mono">{formatNumber(row.aiTransferOut)}</td>
                            <td className="border-r border-[#e8edf4] px-3 py-3 text-right font-mono">{formatNumber(row.aiTotalSupply)}</td>
                            <td className="border-r border-[#e8edf4] px-3 py-3 text-right font-mono">{formatNumber(row.aiFcTotal)}</td>
                            <td className="border-r border-[#e8edf4] px-3 py-3 text-right font-mono">{formatNumber(row.aiQtTotal)}</td>
                            <td className={`border-r border-[#e8edf4] px-3 py-3 text-right font-mono ${row.aiShortageSurplus !== null && row.aiShortageSurplus < 0 ? "text-red-600" : "text-emerald-600"}`}>
                              {formatNumber(row.aiShortageSurplus)}
                            </td>
                            <td className="border-r border-[#e8edf4] px-3 py-3 text-center">
                              <span className={`inline-flex h-8 min-w-20 items-center justify-center rounded-md px-2 text-xs font-bold ${planningStatusTone(status)}`}>
                                {status}
                              </span>
                            </td>
                            <td className="border-r border-[#e8edf4] px-3 py-3">
                              <input
                                className="h-9 w-44 rounded-md border border-[#dfe6ef] px-3 text-sm outline-none focus:border-[#ef3e8f]"
                                placeholder="ใส่แผนที่จะใช้"
                                value={rowFeedback.actual}
                                onChange={(event) => updateFeedback(rowFeedbackId, { actual: event.target.value })}
                              />
                            </td>
                            <td className="px-3 py-3">
                              <input
                                className="h-9 w-60 rounded-md border border-[#dfe6ef] px-3 text-sm outline-none focus:border-[#ef3e8f]"
                                placeholder="เหตุผล/ข้อจำกัด"
                                value={rowFeedback.comment}
                                onChange={(event) => updateFeedback(rowFeedbackId, { comment: event.target.value })}
                              />
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="mt-5 grid gap-4 xl:grid-cols-[0.8fr_1fr_1fr]">
                <label className="block text-sm font-bold">
                  แนวทางที่จะใช้
                  <select
                    className="mt-2 h-11 w-full rounded-md border border-[#dfe6ef] bg-white px-3 text-sm font-semibold outline-none focus:border-[#ef3e8f]"
                    value={currentFeedback.accuracy}
                    onChange={(event) => updateFeedback(feedbackId, { accuracy: event.target.value })}
                  >
                    <option value="">เลือกแนวทาง</option>
                    <option value="ใช้ตาม AI">ใช้ตาม AI</option>
                    <option value="ปรับจาก AI">ปรับจาก AI</option>
                    <option value="ใช้แผนเดิม">ใช้แผนเดิม</option>
                    <option value="รอตรวจสอบเพิ่ม">รอตรวจสอบเพิ่ม</option>
                  </select>
                </label>
                <label className="block text-sm font-bold">
                  ตัวเลข/เป้าหมายที่วางแผน
                  <input
                    className="mt-2 h-11 w-full rounded-md border border-[#dfe6ef] px-3 text-sm font-normal outline-none focus:border-[#ef3e8f]"
                    placeholder="เช่น ใช้ Total Supply ตาม AI หรือปรับเป็น..."
                    value={currentFeedback.actual}
                    onChange={(event) => updateFeedback(feedbackId, { actual: event.target.value })}
                  />
                </label>
                <label className="block text-sm font-bold">
                  หมายเหตุแผน
                  <input
                    className="mt-2 h-11 w-full rounded-md border border-[#dfe6ef] px-3 text-sm font-normal outline-none focus:border-[#ef3e8f]"
                    placeholder="เหตุผล/ข้อจำกัด/สิ่งที่ต้องติดตาม"
                    value={currentFeedback.comment}
                    onChange={(event) => updateFeedback(feedbackId, { comment: event.target.value })}
                  />
                </label>
              </div>

              <div className="mt-4 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-end">
                {savedCard?.id === feedbackId ? (
                  <span className="text-xs font-bold text-emerald-600">บันทึกแล้ว {savedCard.time}</span>
                ) : null}
                <button
                  type="button"
                  className="inline-flex h-10 items-center justify-center gap-2 rounded-md bg-[#ef3e8f] px-4 text-sm font-bold text-white shadow-sm hover:bg-[#dc2e81]"
                  onClick={() => savePlanCard(feedbackId)}
                >
                  <Save size={16} />
                  บันทึกแผนโรงงานนี้
                </button>
              </div>
            </div>
          );
        })() : null}
        {!selectedFactory || planningRows.length === 0 ? (
          <div className="rounded-xl border border-dashed border-[#f5b4cf] bg-white/90 p-8 text-center text-slate-500">
            ยังไม่พบข้อมูลผล AI ตามตัวกรอง
          </div>
        ) : null}
      </div>
    </section>
  );
}

function buildUnifiedPlanningRows(
  templateRows: PlanningBalanceRow[],
  aiRows: BalanceRecord[],
  rawAiRows: AiRecord[],
): UnifiedPlanningRow[] {
  const byProduct = new Map<string, UnifiedPlanningRow>();

  function ensure(product: string) {
    const key = product || "ไม่ระบุกลุ่มสินค้า";
    const current =
      byProduct.get(key) ??
      ({
        product: key,
        sap: "",
        yieldFg: null,
        aiProduction: null,
        aiStock: null,
        aiTransferIn: null,
        aiTransferOut: null,
        aiTotalSupply: null,
        aiFcTotal: null,
        aiQtTotal: null,
        aiShortageSurplus: null,
      } satisfies UnifiedPlanningRow);
    byProduct.set(key, current);
    return current;
  }

  function addMetricValue(current: UnifiedPlanningRow, metricText: string, value: number | null) {
    if (value === null) return;
    const metric = metricText.toLowerCase();
    const setNumber = (field: UnifiedPlanningNumberField) => {
      current[field] = (current[field] ?? 0) + value;
    };

    if (metric.includes("production") || metric.includes("ผลิต")) setNumber("aiProduction");
    else if (metric.includes("stock") || metric.includes("ยกมา")) setNumber("aiStock");
    else if (metric.includes("transfer in") || metric.includes("รับโอน")) setNumber("aiTransferIn");
    else if (metric.includes("transfer out") || metric.includes("โอนออก")) setNumber("aiTransferOut");
    else if (metric.includes("total supply")) setNumber("aiTotalSupply");
    else if (metric.includes("fc total") || metric === "fc" || metric.includes("คาดการณ์ยอดขาย")) {
      setNumber("aiFcTotal");
    } else if (metric.includes("qt total") || metric.includes("quota")) {
      setNumber("aiQtTotal");
    } else if (metric.includes("ขาด") || metric.includes("เหลือ") || metric.includes("shortage")) {
      setNumber("aiShortageSurplus");
    }
  }

  for (const row of templateRows) {
    const current = ensure(row.product);
    current.sap = current.sap || row.sap;
    current.yieldFg = current.yieldFg ?? row.yieldFg;
  }

  for (const row of aiRows) {
    const current = ensure(row.productGroup);
    addMetricValue(current, row.metric, row.aiValue);
  }

  if (rawAiRows.length > 0) {
    const current = ensure("รวมจากไฟล์ผล AI");
    for (const row of rawAiRows) {
      addMetricValue(current, row.metric, row.aiValue);
    }
  }

  return [...byProduct.values()]
    .filter((row) =>
      [
        row.yieldFg,
        row.aiProduction,
        row.aiStock,
        row.aiTransferIn,
        row.aiTransferOut,
        row.aiTotalSupply,
        row.aiFcTotal,
        row.aiQtTotal,
        row.aiShortageSurplus,
      ].some((value) => value !== null),
    )
    .sort((a, b) => {
      if (a.product === "รวมจากไฟล์ผล AI") return -1;
      if (b.product === "รวมจากไฟล์ผล AI") return 1;
      return a.product.localeCompare(b.product, "th");
    });
}

function unifiedPlanningStatus(value: number | null) {
  if (value === null) return "รอข้อมูล";
  if (value < 0) return "ขาด";
  if (value > 0) return "เหลือ";
  return "สมดุล";
}

function planningStatusTone(status: string) {
  if (status === "ขาด") return "bg-rose-50 text-rose-700";
  if (status === "เหลือ") return "bg-emerald-50 text-emerald-700";
  if (status === "สมดุล") return "bg-blue-50 text-blue-700";
  return "bg-slate-100 text-slate-500";
}

function formatPlanningPercent(value: number | null) {
  if (value === null) return "-";
  const percent = Math.abs(value) <= 1 ? value * 100 : value;
  return `${numberFormatter.format(percent)}%`;
}

function PlanningSummaryTile({
  label,
  value,
  tone = "slate",
}: {
  label: string;
  value: string;
  tone?: "slate" | "red" | "green";
}) {
  const toneClass = {
    slate: "border-slate-200 bg-slate-50 text-slate-700",
    red: "border-rose-200 bg-rose-50 text-rose-700",
    green: "border-emerald-200 bg-emerald-50 text-emerald-700",
  }[tone];

  return (
    <div className={`rounded-lg border p-4 ${toneClass}`}>
      <p className="text-sm font-bold">{label}</p>
      <p className="mt-2 text-2xl font-bold">{value}</p>
    </div>
  );
}

function PlanDataEmptyState() {
  return (
    <section className="rounded-xl border border-dashed border-[#f5b4cf] bg-white/90 p-8 text-center shadow-sm">
      <div className="mx-auto grid size-14 place-items-center rounded-full bg-[#ffe8f1] text-[#ef3e8f]">
        <Database size={28} />
      </div>
      <h2 className="mt-4 text-xl font-bold">ยังไม่มีข้อมูล AI เทียบแผนชุดล่าสุด</h2>
      <p className="mx-auto mt-2 max-w-2xl text-sm text-slate-500">
        กรุณาอัปโหลดไฟล์ผล AI และไฟล์แผนที่มีคำว่า Balance ในชื่อไฟล์ให้ครบคู่ ระบบจะล้างผลเทียบเก่าและสร้างข้อมูลใหม่ให้แท็บ Feedback โรงงาน, วิเคราะห์ผล และ AI vs แผน
      </p>
    </section>
  );
}

function PlanHistorySelector({
  items,
  value,
  onChange,
}: {
  items: PlanComparisonHistoryItem[];
  value: string;
  onChange: (id: string) => void;
}) {
  const active = items.find((item) => item.id === value);

  return (
    <section className="rounded-xl border border-[#f5b4cf] bg-white/95 p-4 shadow-sm">
      <div className="grid gap-3 lg:grid-cols-[1fr_320px] lg:items-center">
        <div>
          <p className="text-sm font-bold text-[#ef3e8f]">ชุดข้อมูล AI vs แผน</p>
          <h2 className="mt-1 text-lg font-bold">
            {active ? active.label : "ยังไม่ได้เลือกประวัติ"}
          </h2>
          <p className="mt-1 text-sm text-slate-500">
            {active
              ? `${active.recordCount.toLocaleString("th-TH")} รายการ · อัปโหลด ${new Intl.DateTimeFormat("th-TH", {
                  dateStyle: "medium",
                  timeStyle: "short",
                }).format(new Date(active.uploadedAt))}`
              : "เมื่ออัปโหลด AI + ไฟล์แผน ระบบจะเก็บชุดข้อมูลไว้ให้เลือกย้อนหลัง"}
            {items.length > 0 ? ` · มีประวัติ ${items.length.toLocaleString("th-TH")} ชุด` : ""}
          </p>
        </div>
        <select
          className="h-11 w-full rounded-md border border-[#dfe6ef] bg-white px-4 text-sm font-bold text-slate-700 shadow-sm outline-none focus:border-[#ef3e8f]"
          value={value}
          onChange={(event) => onChange(event.target.value)}
        >
          <option value="">เลือกประวัติไฟล์ที่เคยอัปโหลด</option>
          {items.map((item) => (
            <option key={item.id} value={item.id}>
              {item.label}
            </option>
          ))}
        </select>
      </div>
    </section>
  );
}

function FactoryFeedbackPanel({
  data,
  feedback,
  updateFeedback,
}: {
  data: BalanceData | null;
  feedback: Record<string, Feedback>;
  updateFeedback: (id: string, patch: Partial<Feedback>) => void;
}) {
  const rows = useMemo(() => dedupeBalanceRows(data?.records ?? []), [data]);
  const [weekFilters, setWeekFilters] = useState<string[]>([]);
  const [factoryFilters, setFactoryFilters] = useState<string[]>([]);
  const [savedAt, setSavedAt] = useState<string | null>(null);
  const [savedCard, setSavedCard] = useState<{ id: string; time: string } | null>(null);
  const weeks = Array.from(new Set(rows.map((row) => normalizeWeek(row.week)).filter(Boolean))).sort((a, b) => Number(b) - Number(a));
  const activeWeekFilters =
    weekFilters.length > 0 && weekFilters.every((week) => weeks.includes(week))
      ? weekFilters
      : weeks[0]
        ? [weeks[0]]
        : [];
  const factories = Array.from(new Set(rows.map((row) => row.factory).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b, "th"),
  );
  const visibleFactories = factoryFilters.length > 0 ? factoryFilters : factories;
  const productRows = rows.filter(
    (row) =>
      row.tableType === "product-group" &&
      (activeWeekFilters.length === 0 || activeWeekFilters.includes(normalizeWeek(row.week))) &&
      visibleFactories.includes(row.factory),
  );
  const factoryCards = visibleFactories
    .map((factoryName) => {
      const factoryRows = productRows.filter((row) => row.factory === factoryName);
      const summary = summarizeVdpRows(factoryRows, factoryName);
      const issues = summarizeVdpBy(factoryRows, "productGroup").slice(0, 5);
      return { factoryName, summary, issues, rows: factoryRows };
    })
    .filter((item) => item.rows.length > 0);
  const reasonOptions = [
    "Demand เปลี่ยนจากแผน",
    "Stock จริงไม่ตรงกับแผน",
    "ข้อจำกัดกำลังผลิต",
    "ข้อจำกัดการขนส่ง/โอน",
    "มีคำสั่งขายเร่งด่วน",
    "ข้อมูลตั้งต้นของ AI ไม่ครบ",
  ];

  if (rows.length === 0) {
    return <PlanDataEmptyState />;
  }

  function toggleReason(id: string, reason: string) {
    const current = feedback[id]?.accuracy ? feedback[id].accuracy.split("|").filter(Boolean) : [];
    updateFeedback(id, {
      accuracy: current.includes(reason)
        ? current.filter((item) => item !== reason).join("|")
        : [...current, reason].join("|"),
    });
  }

  function saveFactoryFeedback() {
    window.localStorage.setItem(storageKey, JSON.stringify(feedback));
    setSavedAt(
      new Intl.DateTimeFormat("th-TH", {
        hour: "2-digit",
        minute: "2-digit",
      }).format(new Date()),
    );
  }

  function saveFactoryCard(id: string) {
    window.localStorage.setItem(storageKey, JSON.stringify(feedback));
    setSavedCard({
      id,
      time: new Intl.DateTimeFormat("th-TH", {
        hour: "2-digit",
        minute: "2-digit",
      }).format(new Date()),
    });
  }

  return (
    <section className="space-y-5">
      <div className="rounded-xl border border-[#f5b4cf] bg-white/95 p-5 shadow-sm">
        <div className="mb-4">
          <h2 className="text-xl font-bold">Feedback โรงงาน</h2>
          <p className="mt-1 text-sm text-slate-500">
            สรุปเฉพาะประเด็นสำคัญต่อโรงงาน/สัปดาห์ ให้กรอกเหตุผลรวมโดยไม่ต้องตอบทุกกลุ่มชิ้นส่วน
          </p>
        </div>
        <div className="grid gap-3 lg:grid-cols-2">
          <TransferFilterBox label="สัปดาห์" options={weeks} values={activeWeekFilters} onChange={setWeekFilters} />
          <TransferFilterBox label="โรงงาน" options={factories} values={factoryFilters} onChange={setFactoryFilters} />
        </div>
      </div>

      <div className="grid gap-5">
        {factoryCards.map(({ factoryName, summary, issues }) => {
          const feedbackId = `factory-feedback|${factoryName}|${activeWeekFilters.join(",") || "all"}`;
          const currentFeedback = feedback[feedbackId] ?? { actual: "", accuracy: "", comment: "" };
          const selectedReasons = currentFeedback.accuracy.split("|").filter(Boolean);
          const vdpDiff = summary.aiVdp - summary.balanceVdp;
          const shortageDiff = summary.aiShortage - summary.balanceShortage;
          const surplusDiff = summary.aiSurplus - summary.balanceSurplus;

          return (
            <div key={factoryName} className="rounded-xl border border-[#e3e8f0] bg-white/95 p-5 shadow-sm">
              <div className="mb-4 flex flex-col gap-2 lg:flex-row lg:items-start lg:justify-between">
                <div>
                  <h3 className="text-lg font-bold">{factoryName}</h3>
                  <p className="text-sm text-slate-500">สัปดาห์ {activeWeekFilters.join(", ") || "-"}</p>
                </div>
                <span className="rounded-md bg-[#ffe8f1] px-3 py-1 text-sm font-bold text-[#ef3e8f]">
                  Top {issues.length} กลุ่มที่ต่างจาก AI มากสุด
                </span>
              </div>

              <div className="mb-4 grid gap-3 md:grid-cols-3">
                <FactorySignal
                  label="%VDP"
                  aiValue={`${summary.aiVdp.toFixed(1)}%`}
                  planValue={`${summary.balanceVdp.toFixed(1)}%`}
                  verdict={vdpDiff >= 0 ? `AI ดีกว่าแผน ${formatNumber(Math.abs(vdpDiff))}%` : `แผนดีกว่า AI ${formatNumber(Math.abs(vdpDiff))}%`}
                  detail="ดูจาก %VDP: ค่าสูงกว่าดีกว่า"
                  good={vdpDiff >= 0}
                />
                <FactorySignal
                  label="ของขาด"
                  aiValue={`${formatCompact(summary.aiShortage)} ตัน`}
                  planValue={`${formatCompact(summary.balanceShortage)} ตัน`}
                  verdict={
                    shortageDiff <= 0
                      ? `AI ขาดน้อยกว่าแผน ${formatCompact(Math.abs(shortageDiff))} ตัน`
                      : `แผนขาดน้อยกว่า AI ${formatCompact(Math.abs(shortageDiff))} ตัน`
                  }
                  detail="ดูจากของขาด: ค่าน้อยกว่าดีกว่า"
                  good={shortageDiff <= 0}
                />
                <FactorySignal
                  label="ของเหลือ"
                  aiValue={`${formatCompact(summary.aiSurplus)} ตัน`}
                  planValue={`${formatCompact(summary.balanceSurplus)} ตัน`}
                  verdict={
                    surplusDiff <= 0
                      ? `AI เหลือน้อยกว่าแผน ${formatCompact(Math.abs(surplusDiff))} ตัน`
                      : `แผนเหลือน้อยกว่า AI ${formatCompact(Math.abs(surplusDiff))} ตัน`
                  }
                  detail="ดูจากของเหลือ: ค่าน้อยกว่าดีกว่า"
                  good={surplusDiff <= 0}
                />
              </div>

              <div className="grid gap-5 xl:grid-cols-[1fr_1fr]">
                <div className="rounded-lg border border-[#edf1f6] p-4">
                  <p className="mb-3 font-bold">กลุ่มที่โรงงานใช้ผลต่างจาก AI มากสุด</p>
                  <div className="space-y-2">
                    {issues.map((issue, index) => (
                      <div key={issue.name} className="rounded-md bg-[#f8fafc] px-3 py-2 text-sm">
                        <div className="flex items-center justify-between gap-3">
                          <span className="font-bold">{index + 1}. {issue.name}</span>
                          <span className="font-mono text-xs text-slate-500">%VDP {issue.aiVdp.toFixed(1)}%</span>
                        </div>
                        <p className="mt-1 text-xs text-slate-500">
                          {describeIssueDifference(issue)}
                        </p>
                      </div>
                    ))}
                    {issues.length === 0 ? <p className="text-sm text-slate-500">ไม่พบประเด็นสำคัญ</p> : null}
                  </div>
                </div>

                <div className="rounded-lg border border-[#edf1f6] p-4">
                  <p className="mb-3 font-bold">เหตุผลที่โรงงานใช้ผลต่างจาก AI</p>
                  <div className="grid gap-2 sm:grid-cols-2">
                    {reasonOptions.map((reason) => (
                      <label key={reason} className="flex items-center gap-2 rounded-md border border-[#edf1f6] px-3 py-2 text-sm">
                        <input
                          type="checkbox"
                          className="size-4 accent-[#ef3e8f]"
                          checked={selectedReasons.includes(reason)}
                          onChange={() => toggleReason(feedbackId, reason)}
                        />
                        {reason}
                      </label>
                    ))}
                  </div>
                  <textarea
                    className="mt-4 min-h-28 w-full resize-y rounded-md border border-[#dfe6ef] px-4 py-3 text-sm outline-none focus:border-[#ef3e8f]"
                    placeholder="เขียนเหตุผลรวม เช่น ทำไมโรงงานใช้แผน/ผลจริงที่ต่างจาก AI ในรอบนี้..."
                    value={currentFeedback.comment}
                    onChange={(event) => updateFeedback(feedbackId, { comment: event.target.value })}
                  />
                  <div className="mt-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-end">
                    {savedCard?.id === feedbackId ? (
                      <span className="text-xs font-bold text-emerald-600">บันทึกแล้ว {savedCard.time}</span>
                    ) : null}
                    <button
                      type="button"
                      className="inline-flex h-10 items-center justify-center gap-2 rounded-md bg-[#ef3e8f] px-4 text-sm font-bold text-white shadow-sm hover:bg-[#dc2e81]"
                      onClick={() => saveFactoryCard(feedbackId)}
                    >
                      <Save size={16} />
                      บันทึก Feedback การ์ดนี้
                    </button>
                  </div>
                </div>
              </div>
            </div>
          );
        })}
        {factoryCards.length === 0 ? (
          <div className="rounded-xl border border-dashed border-[#f5b4cf] bg-white/80 p-8 text-center text-slate-500">
            ไม่พบข้อมูลตามตัวกรอง
          </div>
        ) : null}
      </div>

      <div className="flex flex-col items-end gap-2">
        <button
          type="button"
          className="inline-flex h-11 items-center gap-2 rounded-md bg-[#ef3e8f] px-5 text-sm font-bold text-white shadow-sm hover:bg-[#dc2e81]"
          onClick={saveFactoryFeedback}
        >
          <Save size={18} />
          บันทึก Feedback โรงงาน
        </button>
        {savedAt ? <p className="text-xs font-bold text-emerald-600">บันทึกแล้ว {savedAt}</p> : null}
      </div>
    </section>
  );
}

function FactorySignal({
  label,
  aiValue,
  planValue,
  verdict,
  detail,
  good,
}: {
  label: string;
  aiValue: string;
  planValue: string;
  verdict: string;
  detail: string;
  good: boolean;
}) {
  return (
    <div className={`rounded-lg border p-4 ${good ? "border-emerald-200 bg-emerald-50" : "border-rose-200 bg-rose-50"}`}>
      <div className="flex items-start justify-between gap-3">
        <p className={`text-sm font-bold ${good ? "text-emerald-700" : "text-rose-700"}`}>{label}</p>
        <span className={`rounded-md px-2 py-1 text-xs font-bold ${good ? "bg-emerald-600 text-white" : "bg-rose-600 text-white"}`}>
          {good ? "AI ดีกว่า" : "แผนดีกว่า"}
        </span>
      </div>
      <p className="mt-2 text-lg font-bold">{verdict}</p>
      <div className="mt-3 grid gap-2 text-xs font-semibold text-slate-600">
        <div className="flex justify-between gap-3 rounded-md bg-white/70 px-3 py-2">
          <span>AI</span>
          <span className="font-mono">{aiValue}</span>
        </div>
        <div className="flex justify-between gap-3 rounded-md bg-white/70 px-3 py-2">
          <span>แผน/ผลที่โรงงานใช้</span>
          <span className="font-mono">{planValue}</span>
        </div>
      </div>
      <p className="mt-2 text-xs font-semibold text-slate-500">{detail}</p>
    </div>
  );
}

function describeIssueDifference(issue: VdpSummary) {
  const vdpDiff = issue.aiVdp - issue.balanceVdp;
  const shortageDiff = issue.aiShortage - issue.balanceShortage;
  const surplusDiff = issue.aiSurplus - issue.balanceSurplus;
  const parts = [
    vdpDiff >= 0
      ? `AI %VDP สูงกว่าแผน ${Math.abs(vdpDiff).toFixed(1)}%`
      : `แผน %VDP สูงกว่า AI ${Math.abs(vdpDiff).toFixed(1)}%`,
    shortageDiff <= 0
      ? `AI ขาดน้อยกว่าแผน ${formatCompact(Math.abs(shortageDiff))} ตัน`
      : `แผนขาดน้อยกว่า AI ${formatCompact(Math.abs(shortageDiff))} ตัน`,
    surplusDiff <= 0
      ? `AI เหลือน้อยกว่าแผน ${formatCompact(Math.abs(surplusDiff))} ตัน`
      : `แผนเหลือน้อยกว่า AI ${formatCompact(Math.abs(surplusDiff))} ตัน`,
  ];

  return parts.join(" · ");
}

function AnalyzeVdpPanel({
  data,
  aiData,
  feedback,
  uploadedNames,
}: {
  data: BalanceData | null;
  aiData: AiData | null;
  feedback: Record<string, Feedback>;
  uploadedNames: { ai?: string; actual?: string; balance?: string };
}) {
  const [compareMode, setCompareMode] = useState<"plan" | "actual">("plan");
  const [weekFilter, setWeekFilter] = useState<string[]>([]);
  const [factoryFilter, setFactoryFilter] = useState<string[]>([]);
  const planRows = useMemo(() => dedupeBalanceRows(data?.records ?? []), [data]);
  const actualRows = useMemo(() => buildActualAnalyzeRows(aiData, feedback), [aiData, feedback]);
  const rows = compareMode === "plan" ? planRows : actualRows;
  const weeks = Array.from(new Set(rows.map((row) => normalizeWeek(row.week)).filter(Boolean))).sort((a, b) => Number(b) - Number(a));
  const activeWeekFilters =
    weekFilter.length > 0 && weekFilter.every((week) => weeks.includes(week))
      ? weekFilter
      : weeks[0]
        ? [weeks[0]]
        : [];
  const factories = Array.from(new Set(rows.map((row) => row.factory).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b, "th"),
  );
  const filteredRows = rows.filter(
    (row) =>
      (activeWeekFilters.length === 0 || activeWeekFilters.includes(normalizeWeek(row.week))) &&
      (factoryFilter.length === 0 || factoryFilter.includes(row.factory)),
  );
  const productRows = filteredRows.filter((row) => row.tableType === "product-group");
  const directAiVdpByFactory = getAiVdpByFactory(aiData, activeWeekFilters, factoryFilter);
  const factorySummaries = applyDirectAiVdpToFactorySummaries(
    summarizeVdpBy(productRows, "factory"),
    directAiVdpByFactory,
  );
  const productSummaries = summarizeVdpBy(productRows, "productGroup");
  const directAiVdp = averageValues([...directAiVdpByFactory.values()]);
  const total = applyDirectAiVdp(summarizeVdpRows(productRows, "รวมทั้งหมด"), directAiVdp);
  const topProducts = productSummaries.slice(0, 10);
  const vdpDiff = total.aiVdp - total.balanceVdp;
  const shortageDiff = total.aiShortage - total.balanceShortage;
  const surplusDiff = total.aiSurplus - total.balanceSurplus;
  const supplyDiff = total.aiSupply - total.balanceSupply;
  const aiProfit = sumAiProfit(aiData, activeWeekFilters, factoryFilter);
  const dimensionSummaries = [
    {
      label: "%VDP",
      better: vdpDiff >= 0 ? "AI" : "แผน",
      rule: "ค่าสูงกว่าดีกว่า",
      ai: `${total.aiVdp.toFixed(1)}%`,
      balance: `${total.balanceVdp.toFixed(1)}%`,
      diff: `${formatSigned(vdpDiff)}%`,
      tone: vdpDiff >= 0 ? "green" : "red",
    },
    {
      label: "ของขาด",
      better: shortageDiff <= 0 ? "AI" : "แผน",
      rule: "ค่าน้อยกว่าดีกว่า",
      ai: `${formatCompact(total.aiShortage)} ตัน`,
      balance: `${formatCompact(total.balanceShortage)} ตัน`,
      diff: `${formatSigned(shortageDiff)} ตัน`,
      tone: shortageDiff <= 0 ? "green" : "red",
    },
    {
      label: "ของเหลือ",
      better: surplusDiff <= 0 ? "AI" : "แผน",
      rule: "ค่าน้อยกว่าดีกว่า",
      ai: `${formatCompact(total.aiSurplus)} ตัน`,
      balance: `${formatCompact(total.balanceSurplus)} ตัน`,
      diff: `${formatSigned(surplusDiff)} ตัน`,
      tone: surplusDiff <= 0 ? "green" : "orange",
    },
    {
      label: "Total Supply",
      better: Math.abs(supplyDiff) <= Math.abs(total.balanceSupply * 0.02) ? "ใกล้เคียงกัน" : supplyDiff >= 0 ? "AI มากกว่า" : "แผนมากกว่า",
      rule: "ใช้ดูปริมาณ supply รวม",
      ai: `${formatCompact(total.aiSupply)} ตัน`,
      balance: `${formatCompact(total.balanceSupply)} ตัน`,
      diff: `${formatSigned(supplyDiff)} ตัน`,
      tone: Math.abs(supplyDiff) <= Math.abs(total.balanceSupply * 0.02) ? "green" : "orange",
    },
    {
      label: "กำไร",
      better: aiProfit.hasData ? (aiProfit.value >= 0 ? "AI มีกำไร" : "AI ขาดทุน") : "ยังไม่พบข้อมูล",
      rule: "อ่านจาก กำไร/ขาดทุน รวม (Baht) ในไฟล์ AI",
      ai: aiProfit.hasData ? `${formatCompact(aiProfit.value)} บาท` : "ไม่พบในไฟล์ AI/สัปดาห์ที่เลือก",
      balance: "ยังไม่มีคอลัมน์กำไรในไฟล์แผน",
      diff: aiProfit.hasData ? formatSigned(aiProfit.value) : "-",
      tone: aiProfit.hasData && aiProfit.value >= 0 ? "green" : "red",
    },
  ] as const;

  return (
    <section className="analyze-page space-y-5">
      <div className="grid gap-3 rounded-xl border border-[#f5b4cf] bg-white/95 p-4 shadow-sm lg:grid-cols-[240px_1fr_1fr]">
        <div className="grid grid-cols-2 rounded-lg border border-[#ffd1e3] bg-[#fff7fb] p-1">
          <button
            type="button"
            className={`h-10 rounded-md text-sm font-bold ${compareMode === "plan" ? "bg-[#ef3e8f] text-white" : "text-slate-600"}`}
            onClick={() => setCompareMode("plan")}
          >
            AI vs แผน
          </button>
          <button
            type="button"
            className={`h-10 rounded-md text-sm font-bold ${compareMode === "actual" ? "bg-[#ef3e8f] text-white" : "text-slate-600"}`}
            onClick={() => setCompareMode("actual")}
          >
            AI vs Actual
          </button>
        </div>
        <TransferFilterBox label="สัปดาห์" options={weeks} values={activeWeekFilters} onChange={setWeekFilter} />
        <TransferFilterBox label="โรงงาน" options={factories} values={factoryFilter} onChange={setFactoryFilter} />
      </div>

      {compareMode === "plan" && planRows.length === 0 ? <PlanDataEmptyState /> : null}

      {compareMode === "plan" && planRows.length === 0 ? null : <div className="rounded-xl border border-[#f5b4cf] bg-white/95 p-5 shadow-sm">
        <div className="mb-4">
          <p className="text-sm font-bold text-[#ef3e8f]">สรุปแยกตามมิติ</p>
          <h3 className="mt-1 text-2xl font-bold">ดูว่า AI หรือแผนดีกว่าในแต่ละด้าน</h3>
          <p className="mt-2 max-w-4xl text-sm text-slate-600">
            หน้านี้ไม่รวมคะแนนเป็นสูตรเดียว เพื่อให้เห็นชัดว่าฝั่งไหนชนะด้าน %VDP, ของขาด, ของเหลือ, Total Supply และกำไร
          </p>
        </div>
        <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-5">
          {dimensionSummaries.map((item) => (
            <DimensionSummaryCard key={item.label} item={item} />
          ))}
        </div>
      </div>}

      {compareMode === "plan" && planRows.length === 0 ? null : <div className="grid gap-5 xl:grid-cols-[1.05fr_0.95fr]">
        <AnalyzeCard title="%VDP เปรียบเทียบรายโรงงาน">
          <div className="analyze-chart-scroll w-full max-w-full overflow-x-auto overflow-y-hidden pb-3">
            <div className="flex h-[340px] min-w-max items-end gap-3 border-b border-l border-[#dfe6ef] px-4 pb-8">
              {factorySummaries.map((item) => (
                <VdpFactoryBars key={item.name} item={item} />
              ))}
            </div>
          </div>
          <div className="mt-3 flex justify-center gap-5 text-xs font-bold text-slate-500">
            <span className="inline-flex items-center gap-2"><span className="size-3 rounded-sm bg-blue-500" /> AI</span>
            <span className="inline-flex items-center gap-2"><span className="size-3 rounded-sm bg-emerald-500" /> แผน</span>
          </div>
        </AnalyzeCard>

        <AnalyzeCard title="เปรียบเทียบ Supply / ของขาด / ของเหลือ">
          <div className="space-y-5">
            <SupplyCompareBars
              label="Total Supply"
              aiValue={total.aiSupply}
              balanceValue={total.balanceSupply}
              unit="ตัน"
              betterText="ใช้ดูปริมาณรวม ไม่ตัดสินดี/แย่ทันที"
            />
            <SupplyCompareBars
              label="ของขาด"
              aiValue={total.aiShortage}
              balanceValue={total.balanceShortage}
              unit="ตัน"
              betterText="ค่าน้อยกว่าดีกว่า"
              lowerIsBetter
            />
            <SupplyCompareBars
              label="ของเหลือ"
              aiValue={total.aiSurplus}
              balanceValue={total.balanceSurplus}
              unit="ตัน"
              betterText="ค่าน้อยกว่าดีกว่า"
              lowerIsBetter
            />
          </div>
        </AnalyzeCard>
      </div>}

      {compareMode === "plan" && planRows.length === 0 ? null : <div className="grid gap-5 xl:grid-cols-[0.95fr_1.05fr_0.9fr]">
        <AnalyzeCard title="เทียบ KPI สำคัญ">
          <div className="mobile-table-frame rounded-lg border border-[#e8edf4]">
            <div className="mobile-table-scroll overflow-auto">
              <table className="w-full min-w-[680px] border-collapse text-sm">
                <thead className="bg-[#f8fafc] text-xs text-slate-500">
                  <tr>
                    <th className="px-3 py-2 text-left">KPI</th>
                    <th className="px-3 py-2 text-right">AI</th>
                    <th className="px-3 py-2 text-right">แผน</th>
                    <th className="px-3 py-2 text-right">Difference</th>
                    <th className="px-3 py-2 text-right">%Diff</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    ["%VDP", total.aiVdp, total.balanceVdp, "%"],
                    ["Total Supply", total.aiSupply, total.balanceSupply, "ตัน"],
                    ["ของขาด (Shortage)", total.aiShortage, total.balanceShortage, "ตัน"],
                    ["ของเหลือ (Surplus)", total.aiSurplus, total.balanceSurplus, "ตัน"],
                  ].map(([label, ai, balance, unit]) => {
                    const aiValue = Number(ai);
                    const balanceValue = Number(balance);
                    const diff = aiValue - balanceValue;
                    const percent = balanceValue ? (diff / Math.abs(balanceValue)) * 100 : 0;
                    return (
                      <tr key={String(label)} className="border-t border-[#edf1f6]">
                        <td className="px-3 py-2 font-bold">{label}</td>
                        <td className="px-3 py-2 text-right font-mono">{formatNumber(aiValue)} {unit}</td>
                        <td className="px-3 py-2 text-right font-mono">{formatNumber(balanceValue)} {unit}</td>
                        <td className={`px-3 py-2 text-right font-mono ${diff >= 0 ? "text-emerald-600" : "text-red-600"}`}>{formatSigned(diff)}</td>
                        <td className={`px-3 py-2 text-right font-mono ${diff >= 0 ? "text-emerald-600" : "text-red-600"}`}>{formatSigned(percent)}%</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </AnalyzeCard>

        <AnalyzeCard title="Top 10 สินค้า/ชิ้นส่วน ที่คลาดเคลื่อนมากที่สุด">
          <div className="mobile-table-frame rounded-lg border border-[#e8edf4]">
            <div className="mobile-table-scroll overflow-auto">
              <table className="w-full min-w-[780px] border-collapse text-sm">
                <thead className="bg-[#f8fafc] text-xs text-slate-500">
                  <tr>
                    <th className="px-3 py-2 text-left">สินค้า/ชิ้นส่วน</th>
                    <th className="px-3 py-2 text-right">%VDP AI</th>
                    <th className="px-3 py-2 text-right">%VDP แผน</th>
                    <th className="px-3 py-2 text-right">ของขาดต่าง</th>
                    <th className="px-3 py-2 text-right">ของเหลือต่าง</th>
                  </tr>
                </thead>
                <tbody>
                  {topProducts.map((item, index) => (
                    <tr key={item.name} className="border-t border-[#edf1f6]">
                      <td className="px-3 py-2 font-medium">{index + 1}. {item.name}</td>
                      <td className="px-3 py-2 text-right font-mono">{item.aiVdp.toFixed(1)}%</td>
                      <td className="px-3 py-2 text-right font-mono">{item.balanceVdp.toFixed(1)}%</td>
                      <td className="px-3 py-2 text-right font-mono text-red-600">{formatSigned(item.aiShortage - item.balanceShortage)}</td>
                      <td className="px-3 py-2 text-right font-mono text-orange-600">{formatSigned(item.aiSurplus - item.balanceSurplus)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </AnalyzeCard>

        <AnalyzeCard title="Impact จากการใช้ AI">
          <div className="space-y-3">
            <Insight tone={vdpDiff >= 0 ? "green" : "red"} icon={<TrendingUp size={18} />} text={`${vdpDiff >= 0 ? "เพิ่ม" : "ลด"} %VDP ${formatSigned(vdpDiff)}%`} />
            <Insight tone="red" icon={<AlertCircle size={18} />} text={`ของขาดต่างจากแผน ${formatSigned(shortageDiff)} ตัน`} />
            <Insight tone="yellow" icon={<Sparkles size={18} />} text={`ของเหลือต่างจากแผน ${formatSigned(surplusDiff)} ตัน`} />
            <Insight tone="green" icon={<CheckCircle2 size={18} />} text={`ไฟล์ AI: ${uploadedNames.ai ?? data?.aiFile ?? "-"}`} />
            <Insight tone="yellow" icon={<Database size={18} />} text={`ไฟล์แผน: ${uploadedNames.balance ?? data?.sourceFile ?? "-"}`} />
          </div>
        </AnalyzeCard>
      </div>}
    </section>
  );
}

function AnalyzePanel({
  data,
  uploadedNames,
}: {
  data: BalanceData | null;
  uploadedNames: { ai?: string; actual?: string; balance?: string };
}) {
  const rows = useMemo(() => dedupeBalanceRows(data?.records ?? []), [data]);
  const productRows = rows.filter((row) => row.tableType === "product-group");
  const factories = Array.from(new Set(rows.map((row) => row.factory).filter(Boolean)));
  const weeks = Array.from(new Set(rows.map((row) => normalizeWeek(row.week)).filter(Boolean))).sort((a, b) => Number(b) - Number(a));
  const weekLabel = weeks[0] ? `สัปดาห์ที่ ${weeks[0]}` : "สัปดาห์ล่าสุด";
  const comparableRows = rows.filter((row) => row.aiValue !== null && row.balanceValue !== null);
  const planTotal = sumBalanceMetric(productRows, "ผลิต", "balanceValue");
  const aiTotal = sumBalanceMetric(productRows, "ผลิต", "aiValue");
  const diffTotal = aiTotal - planTotal;
  const diffPercent = planTotal ? (diffTotal / Math.abs(planTotal)) * 100 : 0;
  const avgAccuracy = comparableRows.length
    ? comparableRows.reduce((sum, row) => sum + analyzeScore(row.aiValue, row.balanceValue), 0) / comparableRows.length
    : 0;
  const majorDiffs = comparableRows.filter((row) => analyzeScore(row.aiValue, row.balanceValue) < 70).length;
  const factorySummaries = summarizeAnalyzeBy(productRows, "factory");
  const topProducts = summarizeAnalyzeBy(productRows, "productGroup").slice(0, 10);
  const trend = [0.45, 1.12, 2.35, Number(diffTotal.toFixed(2))];

  return (
    <section className="analyze-page space-y-5">
      <div className="relative overflow-hidden rounded-xl border border-[#dbe8f5] bg-white shadow-sm">
        <div className="absolute inset-0 bg-[url('/background.png')] bg-cover bg-center opacity-55" />
        <div className="absolute inset-0 bg-gradient-to-r from-white via-white/80 to-white/20" />
        <div className="relative flex min-h-40 flex-col justify-between gap-5 p-5 lg:flex-row lg:items-start">
          <div>
            <h2 className="text-2xl font-bold text-[#172033]">AI vs แผน Overview</h2>
            <p className="mt-2 max-w-2xl text-sm font-medium text-slate-600">
              ภาพรวมการเปรียบเทียบผลลัพธ์ AI กับแผนของโรงงาน เพื่อดูผลกระทบเชิง Business
            </p>
          </div>
          <div className="flex flex-wrap gap-3">
            <AnalyzePill>{weekLabel} <CalendarDays size={16} /></AnalyzePill>
            <AnalyzePill>โรงงานทั้งหมด {factorySummaries.length || factories.length || "-"} แห่ง <ChevronDown size={16} /></AnalyzePill>
            <AnalyzePill><Download size={16} /> ดาวน์โหลดรายงาน</AnalyzePill>
          </div>
        </div>
      </div>

      <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-5">
        <AnalyzeKpi icon={<Database size={28} />} label="มูลค่าการผลิตตามแผน (Plan)" value={formatCompact(planTotal)} detail="100% ของแผนทั้งหมด" tone="blue" />
        <AnalyzeKpi icon={<Target size={28} />} label="มูลค่าผลลัพธ์จาก AI" value={formatCompact(aiTotal)} detail={`${formatSigned(diffTotal)} (${formatSigned(diffPercent)}%)`} tone="green" />
        <AnalyzeKpi icon={<TrendingUp size={28} />} label="มูลค่าความแตกต่าง" value={formatSigned(diffTotal)} detail={`${formatSigned(diffPercent)}% เทียบกับแผน`} tone={diffTotal >= 0 ? "orange" : "rose"} />
        <AnalyzeKpi icon={<Gauge size={28} />} label="ประสิทธิภาพการเทียบผล" value={`${avgAccuracy.toFixed(1)}%`} detail="AI เทียบกับแผน" tone="teal" />
        <AnalyzeKpi icon={<AlertCircle size={28} />} label="รายการต่างกันมาก" value={`${majorDiffs.toLocaleString("th-TH")}`} detail="ต้องตรวจสอบเชิง Business" tone="purple" />
      </div>

      <div className="grid gap-5 xl:grid-cols-[1.05fr_0.95fr]">
        <AnalyzeCard title="เปรียบเทียบมูลค่ารวมรายโรงงาน">
          <div className="max-h-[520px] space-y-3 overflow-auto pr-2">
            {factorySummaries.map((item) => (
              <AnalyzeBarRow key={item.name} item={item} />
            ))}
            {factorySummaries.length === 0 ? (
              <div className="rounded-lg border border-dashed border-[#dfe6ef] p-6 text-center text-sm text-slate-500">
                ไม่พบข้อมูลรายโรงงาน
              </div>
            ) : null}
          </div>
        </AnalyzeCard>

        <AnalyzeCard title="ความแตกต่างแยกตามสินค้า/ชิ้นส่วน (Top 10)">
          <div className="mobile-table-frame rounded-lg border border-[#e8edf4]">
            <div className="mobile-table-scroll overflow-auto">
              <table className="w-full min-w-[760px] border-collapse text-sm">
                <thead className="bg-[#f8fafc] text-xs text-slate-500">
                  <tr>
                    <th className="px-3 py-2 text-left">สินค้า/ชิ้นส่วน</th>
                    <th className="px-3 py-2 text-right">Plan</th>
                    <th className="px-3 py-2 text-right">AI</th>
                    <th className="px-3 py-2 text-right">ต่างกัน</th>
                    <th className="px-3 py-2 text-right">ต่างกัน (%)</th>
                  </tr>
                </thead>
                <tbody>
                  {topProducts.map((item, index) => (
                    <tr key={item.name} className="border-t border-[#edf1f6]">
                      <td className="px-3 py-2 font-medium">{index + 1}. {item.name}</td>
                      <td className="px-3 py-2 text-right font-mono">{formatNumber(item.plan)}</td>
                      <td className="px-3 py-2 text-right font-mono">{formatNumber(item.ai)}</td>
                      <td className={`px-3 py-2 text-right font-mono ${item.diff >= 0 ? "text-emerald-600" : "text-red-600"}`}>
                        {formatSigned(item.diff)}
                      </td>
                      <td className={`px-3 py-2 text-right font-mono ${item.diff >= 0 ? "text-emerald-600" : "text-red-600"}`}>
                        {formatSigned(item.percent)}%
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </AnalyzeCard>
      </div>

      <div className="grid gap-5 xl:grid-cols-[0.85fr_0.85fr_1fr_1.15fr]">
        <AnalyzeCard title="Impact ทางการเงิน">
          <div className="space-y-4 text-sm">
            <AnalyzeMoney label="เพิ่มขึ้นจากแผน" value={formatSigned(Math.max(diffTotal, 0))} />
            <AnalyzeMoney label="คิดเป็น" value={`${formatSigned(diffPercent)}%`} />
            <AnalyzeMoney label="หากทำได้ต่อเนื่องทั้งปี" value={formatSigned(diffTotal * 52)} />
          </div>
        </AnalyzeCard>

        <AnalyzeCard title="ความแม่นยำของ AI">
          <div className="flex items-center gap-5">
            <div className="grid size-32 place-items-center rounded-full bg-[conic-gradient(#3b82f6_var(--score),#e8eef8_0)]" style={{ "--score": `${avgAccuracy}%` } as CSSProperties}>
              <div className="grid size-20 place-items-center rounded-full bg-white text-xl font-bold">{avgAccuracy.toFixed(1)}%</div>
            </div>
            <div className="flex-1 space-y-2 text-sm">
              <MetricLine label="R²" value={(avgAccuracy / 100).toFixed(3)} />
              <MetricLine label="MAPE" value={`${(100 - avgAccuracy).toFixed(2)}%`} />
              <MetricLine label="MAE" value={formatCompact(Math.abs(diffTotal) / Math.max(factories.length, 1))} />
            </div>
          </div>
        </AnalyzeCard>

        <AnalyzeCard title="สรุปประเด็นสำคัญ">
          <div className="space-y-3 text-sm">
            <Insight tone="green" icon={<CheckCircle2 size={18} />} text={`AI ให้ผลรวม ${diffTotal >= 0 ? "สูงกว่า" : "ต่ำกว่า"}แผน ${formatSigned(diffTotal)}`} />
            <Insight tone="yellow" icon={<Sparkles size={18} />} text={`สินค้า/ชิ้นส่วนที่ต่างมากสุด: ${topProducts[0]?.name ?? "-"}`} />
            <Insight tone="red" icon={<AlertCircle size={18} />} text={`${majorDiffs.toLocaleString("th-TH")} รายการควรตรวจสอบก่อนใช้งานจริง`} />
          </div>
        </AnalyzeCard>

        <AnalyzeCard title="แนวโน้ม 4 สัปดาห์ล่าสุด">
          <div className="flex h-44 items-end gap-4 border-b border-l border-[#dfe6ef] px-4 pb-4">
            {trend.map((value, index) => (
              <div key={index} className="flex flex-1 flex-col items-center gap-2">
                <span className="text-xs font-bold text-slate-600">{formatSigned(value)}</span>
                <div className="w-full rounded-t-md bg-emerald-400" style={{ height: `${Math.max(18, Math.min(130, Math.abs(value) * 24))}px` }} />
                <span className="text-xs text-slate-500">W{index + 18}</span>
              </div>
            ))}
          </div>
        </AnalyzeCard>
      </div>

      <p className="text-right text-xs text-slate-500">
        AI: {uploadedNames.ai ?? data?.aiFile ?? "-"} | แผน: {uploadedNames.balance ?? data?.sourceFile ?? "-"}
      </p>
    </section>
  );
}

void AnalyzePanel;

type AnalyzeSummary = {
  name: string;
  plan: number;
  ai: number;
  diff: number;
  percent: number;
};

type VdpSummary = {
  name: string;
  aiSupply: number;
  balanceSupply: number;
  aiShortage: number;
  balanceShortage: number;
  aiSurplus: number;
  balanceSurplus: number;
  aiVdp: number;
  balanceVdp: number;
  severity: number;
};

function buildActualAnalyzeRows(aiData: AiData | null, feedback: Record<string, Feedback>): BalanceRecord[] {
  if (!aiData) return [];

  return aiData.records
    .filter((record) => record.kind === "number" && feedback[record.id]?.actual)
    .flatMap((record) => {
      const actualValue = numericValue(feedback[record.id]?.actual ?? "");
      if (actualValue === null) return [];
      const week = record.weeks[record.weeks.length - 1] ?? "";
      const metric = normalizeAnalyzeMetric(record.metric);
      if (!metric) return [];

      return [
        {
          id: `actual-analyze|${record.id}`,
          tableType: "product-group" as const,
          sourceSheet: record.sheet,
          factory: record.factory,
          week,
          productGroup: record.metric,
          metric,
          aiMetric: record.metric,
          aiValue: record.aiValue,
          balanceValue: actualValue,
        },
      ];
    });
}

function normalizeAnalyzeMetric(metric: string) {
  const lower = metric.toLowerCase();
  if (metric.includes("ขาด") || metric.includes("เหลือ")) return "สินค้าที่ขาด/เหลือจากการบาล้าน";
  if (metric.includes("ผลิต")) return "ผลิต";
  if (lower.includes("total supply")) return "Total Supply";
  if (lower.includes("ของขาด") || lower.includes("ของเหลือ") || lower.includes("shortage")) {
    return "สินค้าที่ขาด/เหลือจากการบาล้าน";
  }
  if (lower.includes("production") || metric.includes("ผลิต")) return "ผลิต";
  return "";
}

function summarizeVdpBy(rows: BalanceRecord[], key: "factory" | "productGroup") {
  const grouped = new Map<string, BalanceRecord[]>();
  rows.forEach((row) => {
    const name = row[key] || "-";
    const current = grouped.get(name) ?? [];
    current.push(row);
    grouped.set(name, current);
  });

  return Array.from(grouped.entries())
    .map(([name, groupRows]) => summarizeVdpRows(groupRows, name))
    .sort((a, b) => b.severity - a.severity);
}

function summarizeVdpRows(rows: BalanceRecord[], name: string): VdpSummary {
  const summary = {
    name,
    aiSupply: 0,
    balanceSupply: 0,
    aiShortage: 0,
    balanceShortage: 0,
    aiSurplus: 0,
    balanceSurplus: 0,
  };

  rows.forEach((row) => {
    const metric = row.metric.toLowerCase();
    if (metric.includes("total supply")) {
      summary.aiSupply += row.aiValue ?? 0;
      summary.balanceSupply += row.balanceValue ?? 0;
    }

    if (metric.includes("ขาด") || metric.includes("เหลือ") || metric.includes("shortage")) {
      addShortageSurplus(summary, row.aiValue ?? 0, "ai");
      addShortageSurplus(summary, row.balanceValue ?? 0, "balance");
    }
  });

  const aiVdp = computeVdp(summary.aiSupply, summary.aiShortage);
  const balanceVdp = computeVdp(summary.balanceSupply, summary.balanceShortage);

  return {
    ...summary,
    aiVdp,
    balanceVdp,
    severity:
      Math.abs(aiVdp - balanceVdp) +
      Math.abs(summary.aiShortage - summary.balanceShortage) / 100 +
      Math.abs(summary.aiSurplus - summary.balanceSurplus) / 100,
  };
}

function getAiVdpByFactory(aiData: AiData | null, weeks: string[], factories: string[]) {
  const grouped = new Map<string, number[]>();
  if (!aiData) return new Map<string, number>();

  aiData.records.forEach((record) => {
    const metric = record.metric.toLowerCase();
    const isVdp = metric.includes("vdp");
    const matchesWeek = weeks.length === 0 || record.weeks.some((week) => weeks.includes(normalizeWeek(week)));
    const matchesFactory = factories.length === 0 || factories.includes(record.factory);
    const value = record.average ?? record.aiValue;
    if (record.kind !== "number" || !isVdp || !matchesWeek || !matchesFactory || value === null) return;

    const current = grouped.get(record.factory) ?? [];
    current.push(normalizePercentValue(value));
    grouped.set(record.factory, current);
  });

  return new Map([...grouped.entries()].map(([factory, values]) => [factory, averageValues(values) ?? 0]));
}

function applyDirectAiVdpToFactorySummaries(summaries: VdpSummary[], directVdp: Map<string, number>) {
  const merged = new Map(summaries.map((summary) => [summary.name, applyDirectAiVdp(summary, directVdp.get(summary.name) ?? null)]));

  directVdp.forEach((vdp, factory) => {
    if (merged.has(factory)) return;
    merged.set(factory, {
      name: factory,
      aiSupply: 0,
      balanceSupply: 0,
      aiShortage: 0,
      balanceShortage: 0,
      aiSurplus: 0,
      balanceSurplus: 0,
      aiVdp: vdp,
      balanceVdp: 0,
      severity: Math.abs(vdp),
    });
  });

  return [...merged.values()].sort((a, b) => b.severity - a.severity);
}

function applyDirectAiVdp(summary: VdpSummary, directVdp: number | null) {
  if (directVdp === null) return summary;
  return {
    ...summary,
    aiVdp: directVdp,
    severity:
      Math.abs(directVdp - summary.balanceVdp) +
      Math.abs(summary.aiShortage - summary.balanceShortage) / 100 +
      Math.abs(summary.aiSurplus - summary.balanceSurplus) / 100,
  };
}

function normalizePercentValue(value: number) {
  return Math.abs(value) <= 1 ? value * 100 : value;
}

function averageValues(values: number[]) {
  if (values.length === 0) return null;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function addShortageSurplus(
  summary: {
    aiShortage: number;
    balanceShortage: number;
    aiSurplus: number;
    balanceSurplus: number;
  },
  value: number,
  target: "ai" | "balance",
) {
  const shortageKey = target === "ai" ? "aiShortage" : "balanceShortage";
  const surplusKey = target === "ai" ? "aiSurplus" : "balanceSurplus";

  if (value < 0) {
    summary[shortageKey] += Math.abs(value);
  } else {
    summary[surplusKey] += value;
  }
}

function computeVdp(supply: number, shortage: number) {
  if (supply <= 0) return 0;
  return Math.max(0, Math.min(100, ((supply - shortage) / supply) * 100));
}

function sumAiProfit(aiData: AiData | null, weeks: string[], factories: string[]) {
  if (!aiData) return { value: 0, hasData: false };

  const matchingFactoryRecords = aiData.records.filter((record) => {
    const isTotalProfit = isTotalProfitMetric(record.metric);
    const matchesFactory = factories.length === 0 || factories.includes(record.factory);
    return record.kind === "number" && isTotalProfit && matchesFactory;
  });
  const weekMatchedRecords = matchingFactoryRecords.filter(
    (record) => weeks.length === 0 || record.weeks.some((week) => weeks.includes(normalizeWeek(week))),
  );
  const records = weekMatchedRecords.length > 0 ? weekMatchedRecords : matchingFactoryRecords;

  return {
    value: records.reduce((sum, record) => sum + (record.aiValue ?? 0), 0),
    hasData: records.length > 0,
  };
}

function isTotalProfitMetric(metric: string) {
  const normalized = metric.toLowerCase().replace(/\s+/g, "");
  return (
    (normalized.includes("กำไร") && normalized.includes("ขาดทุน") && normalized.includes("รวม")) ||
    normalized.includes("profittotal") ||
    normalized.includes("totalprofit")
  );
}

function sumBalanceMetric(rows: BalanceRecord[], metric: string, field: "aiValue" | "balanceValue") {
  return rows
    .filter((row) => row.metric.toLowerCase().includes(metric.toLowerCase()))
    .reduce((sum, row) => sum + (row[field] ?? 0), 0);
}

function summarizeAnalyzeBy(rows: BalanceRecord[], key: "factory" | "productGroup"): AnalyzeSummary[] {
  const summaries = new Map<string, { plan: number; ai: number }>();
  rows
    .filter((row) => row.metric === "ผลิต")
    .forEach((row) => {
      const name = row[key] || "-";
      const current = summaries.get(name) ?? { plan: 0, ai: 0 };
      current.plan += row.balanceValue ?? 0;
      current.ai += row.aiValue ?? 0;
      summaries.set(name, current);
    });

  return Array.from(summaries.entries())
    .map(([name, value]) => {
      const diff = value.ai - value.plan;
      return {
        name,
        plan: value.plan,
        ai: value.ai,
        diff,
        percent: value.plan ? (diff / Math.abs(value.plan)) * 100 : 0,
      };
    })
    .sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
}

function analyzeScore(aiValue: number | null, balanceValue: number | null) {
  if (aiValue === null || balanceValue === null) return 0;
  const denominator = Math.max(Math.abs(balanceValue), 1);
  const errorRate = Math.abs(aiValue - balanceValue) / denominator;
  return Math.max(0, Math.round((1 - errorRate) * 100));
}

function formatCompact(value: number) {
  const abs = Math.abs(value);
  if (abs >= 1_000_000) return `${numberFormatter.format(value / 1_000_000)} ล้าน`;
  if (abs >= 1_000) return `${numberFormatter.format(value / 1_000)} พัน`;
  return numberFormatter.format(value);
}

function formatSigned(value: number) {
  const formatted = numberFormatter.format(value);
  return value > 0 ? `+${formatted}` : formatted;
}

function AnalyzePill({ children }: { children: ReactNode }) {
  return (
    <button className="inline-flex h-11 items-center gap-2 rounded-md border border-[#dfe6ef] bg-white/90 px-4 text-sm font-bold text-slate-700 shadow-sm">
      {children}
    </button>
  );
}

function AnalyzeKpi({
  icon,
  label,
  value,
  detail,
  tone,
}: {
  icon: ReactNode;
  label: string;
  value: string;
  detail: string;
  tone: "blue" | "green" | "orange" | "rose" | "teal" | "purple";
}) {
  const tones = {
    blue: "bg-blue-50 text-blue-600",
    green: "bg-emerald-50 text-emerald-600",
    orange: "bg-orange-50 text-orange-600",
    rose: "bg-rose-50 text-rose-600",
    teal: "bg-teal-50 text-teal-600",
    purple: "bg-violet-50 text-violet-600",
  };

  return (
    <div className="rounded-xl border border-[#e3e8f0] bg-white/90 p-5 shadow-sm">
      <div className="flex items-center gap-4">
        <div className={`grid size-14 place-items-center rounded-lg ${tones[tone]}`}>{icon}</div>
        <div>
          <p className="text-xs font-bold text-slate-500">{label}</p>
          <p className="mt-1 text-2xl font-bold">{value}</p>
          <p className="mt-1 text-xs font-semibold text-slate-500">{detail}</p>
        </div>
      </div>
    </div>
  );
}

function DimensionSummaryCard({
  item,
}: {
  item: {
    label: string;
    better: string;
    rule: string;
    ai: string;
    balance: string;
    diff: string;
    tone: "green" | "red" | "orange";
  };
}) {
  const toneClass = {
    green: "border-emerald-200 bg-emerald-50 text-emerald-700",
    red: "border-rose-200 bg-rose-50 text-rose-700",
    orange: "border-orange-200 bg-orange-50 text-orange-700",
  }[item.tone];

  return (
    <div className={`rounded-xl border p-4 ${toneClass}`}>
      <div className="flex items-start justify-between gap-3">
        <div>
          <p className="text-sm font-bold">{item.label}</p>
          <p className="mt-1 text-xs font-semibold opacity-80">{item.rule}</p>
        </div>
        <span className="rounded-md bg-white/70 px-2 py-1 text-xs font-bold">{item.better}</span>
      </div>
      <div className="mt-4 space-y-2 text-xs font-semibold">
        <div className="flex justify-between gap-3">
          <span>AI</span>
          <span className="font-mono">{item.ai}</span>
        </div>
        <div className="flex justify-between gap-3">
          <span>แผน</span>
          <span className="font-mono">{item.balance}</span>
        </div>
      </div>
      <div className="mt-4 rounded-lg bg-white/70 px-3 py-2 text-right font-mono text-lg font-bold">
        {item.diff}
      </div>
    </div>
  );
}

function ProfitIndexCard({
  label,
  value,
  vdp,
  lossRate,
  active,
}: {
  label: string;
  value: number;
  vdp: number;
  lossRate: number;
  active: boolean;
}) {
  return (
    <div
      className={`rounded-xl border p-4 shadow-sm ${
        active ? "border-emerald-200 bg-emerald-50" : "border-[#e3e8f0] bg-white/90"
      }`}
    >
      <div className="flex items-center justify-between gap-3">
        <p className="font-bold">{label}</p>
        {active ? <span className="rounded-md bg-emerald-600 px-2 py-1 text-xs font-bold text-white">ดีกว่า</span> : null}
      </div>
      <p className="mt-2 text-2xl font-bold">{value.toFixed(1)}%</p>
      <p className="mt-1 text-xs font-semibold text-slate-500">
        %VDP {vdp.toFixed(1)}% - อัตราของขาด/เหลือ {lossRate.toFixed(1)}%
      </p>
    </div>
  );
}

void ProfitIndexCard;

function AnalyzeCard({ title, children }: { title: string; children: ReactNode }) {
  return (
    <div className="min-w-0 overflow-hidden rounded-xl border border-[#e3e8f0] bg-white/90 p-5 shadow-sm">
      <h3 className="mb-4 text-lg font-bold">{title}</h3>
      {children}
    </div>
  );
}

function AnalyzeBarRow({ item }: { item: AnalyzeSummary }) {
  const max = Math.max(Math.abs(item.plan), Math.abs(item.ai), 1);
  return (
    <div className="grid gap-2 rounded-lg border border-[#edf1f6] p-3">
      <div className="flex items-center justify-between gap-3 text-sm">
        <span className="font-bold">{item.name}</span>
        <span className={item.diff >= 0 ? "font-mono text-emerald-600" : "font-mono text-red-600"}>
          {formatSigned(item.diff)}
        </span>
      </div>
      <div className="grid gap-1">
        <div className="h-3 rounded-full bg-slate-100">
          <div className="h-3 rounded-full bg-slate-400" style={{ width: `${Math.min(100, Math.abs(item.plan / max) * 100)}%` }} />
        </div>
        <div className="h-3 rounded-full bg-blue-50">
          <div className="h-3 rounded-full bg-blue-500" style={{ width: `${Math.min(100, Math.abs(item.ai / max) * 100)}%` }} />
        </div>
      </div>
    </div>
  );
}

function VdpFactoryBars({ item }: { item: VdpSummary }) {
  return (
    <div className="flex w-24 shrink-0 flex-col items-center gap-2">
      <div className="flex h-56 items-end gap-2">
        <div className="relative w-8 rounded-t-md bg-blue-500" style={{ height: `${Math.max(8, item.aiVdp * 2)}px` }}>
          <span className="absolute -top-6 left-1/2 -translate-x-1/2 text-xs font-bold text-blue-700">{item.aiVdp.toFixed(1)}%</span>
        </div>
        <div className="relative w-8 rounded-t-md bg-emerald-500" style={{ height: `${Math.max(8, item.balanceVdp * 2)}px` }}>
          <span className="absolute -top-6 left-1/2 -translate-x-1/2 text-xs font-bold text-emerald-700">{item.balanceVdp.toFixed(1)}%</span>
        </div>
      </div>
      <span className="line-clamp-2 min-h-10 text-center text-xs font-bold text-slate-600">{item.name}</span>
    </div>
  );
}

function SupplyCompareBars({
  label,
  aiValue,
  balanceValue,
  unit,
  betterText,
  lowerIsBetter = false,
}: {
  label: string;
  aiValue: number;
  balanceValue: number;
  unit: string;
  betterText: string;
  lowerIsBetter?: boolean;
}) {
  const max = Math.max(Math.abs(aiValue), Math.abs(balanceValue), 1);
  const diff = aiValue - balanceValue;
  const aiBetter = lowerIsBetter ? aiValue <= balanceValue : aiValue >= balanceValue;

  return (
    <div className="rounded-lg border border-[#edf1f6] p-4">
      <div className="mb-3 flex flex-col gap-1 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <p className="font-bold">{label}</p>
          <p className="text-xs font-semibold text-slate-500">{betterText}</p>
        </div>
        <span className={`text-sm font-bold ${diff >= 0 ? "text-emerald-600" : "text-red-600"}`}>
          ต่างกัน {formatSigned(diff)} {unit}
        </span>
      </div>
      <div className="grid gap-3">
        <CompareBar
          label="AI"
          value={aiValue}
          max={max}
          unit={unit}
          color="bg-blue-500"
          active={aiBetter && lowerIsBetter}
        />
        <CompareBar
          label="แผน"
          value={balanceValue}
          max={max}
          unit={unit}
          color="bg-emerald-500"
          active={!aiBetter && lowerIsBetter}
        />
      </div>
    </div>
  );
}

function CompareBar({
  label,
  value,
  max,
  unit,
  color,
  active,
}: {
  label: string;
  value: number;
  max: number;
  unit: string;
  color: string;
  active: boolean;
}) {
  return (
    <div className="grid gap-1">
      <div className="flex items-center justify-between gap-3 text-xs font-bold text-slate-600">
        <span>{label}</span>
        <span className="font-mono">
          {formatCompact(value)} {unit}
          {active ? " · ดีกว่า" : ""}
        </span>
      </div>
      <div className="h-4 overflow-hidden rounded-full bg-slate-100">
        <div className={`h-full rounded-full ${color}`} style={{ width: `${Math.max(3, Math.min(100, (Math.abs(value) / max) * 100))}%` }} />
      </div>
    </div>
  );
}

void WaterfallBar;

function WaterfallBar({ label, value, tone }: { label: string; value: number; tone: "slate" | "green" | "red" | "orange" | "blue" }) {
  const colors = {
    slate: "bg-slate-400 text-slate-700",
    green: "bg-emerald-500 text-emerald-700",
    red: "bg-red-500 text-red-700",
    orange: "bg-orange-400 text-orange-700",
    blue: "bg-blue-500 text-blue-700",
  };
  const height = Math.max(24, Math.min(230, Math.abs(value) / 120));

  return (
    <div className="flex min-w-20 flex-1 flex-col items-center gap-2">
      <span className={`text-xs font-bold ${colors[tone].split(" ")[1]}`}>{formatSigned(value)}</span>
      <div className={`w-full max-w-16 rounded-t-md ${colors[tone].split(" ")[0]}`} style={{ height: `${height}px` }} />
      <span className="line-clamp-2 min-h-9 text-center text-xs font-semibold text-slate-500">{label}</span>
    </div>
  );
}

function AnalyzeMoney({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <p className="font-bold text-slate-600">{label}</p>
      <p className="mt-1 text-2xl font-bold text-emerald-600">{value}</p>
    </div>
  );
}

function MetricLine({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex items-center justify-between border-b border-[#edf1f6] pb-2">
      <span className="font-bold text-slate-500">{label}</span>
      <span className="font-mono font-bold">{value}</span>
    </div>
  );
}

function Insight({ tone, icon, text }: { tone: "green" | "yellow" | "red"; icon: ReactNode; text: string }) {
  const tones = {
    green: "border-emerald-100 bg-emerald-50 text-emerald-700",
    yellow: "border-amber-100 bg-amber-50 text-amber-700",
    red: "border-rose-100 bg-rose-50 text-rose-700",
  };

  return (
    <div className={`flex items-start gap-3 rounded-lg border p-3 font-semibold ${tones[tone]}`}>
      <span className="mt-0.5 shrink-0">{icon}</span>
      <span>{text}</span>
    </div>
  );
}


function BalanceComparison({
  data,
  feedback,
  updateFeedback,
}: {
  data: BalanceData | null;
  feedback: Record<string, Feedback>;
  updateFeedback: (id: string, patch: Partial<Feedback>) => void;
}) {
  const rows = useMemo(() => data?.records ?? [], [data]);
  const [factoryFilters, setFactoryFilters] = useState<string[]>([]);
  const [weekFilters, setWeekFilters] = useState<string[]>([]);
  const [groupFilters, setGroupFilters] = useState<string[]>([]);
  const [factoryMetricFilters, setFactoryMetricFilters] = useState<string[]>([]);
  const [productMetricFilters, setProductMetricFilters] = useState<string[]>([]);
  const [factoryStatusFilters, setFactoryStatusFilters] = useState<string[]>(["ต่างปานกลาง", "ต่างกันมาก"]);
  const [productStatusFilters, setProductStatusFilters] = useState<string[]>(["ต่างปานกลาง", "ต่างกันมาก"]);
  const [savedAt, setSavedAt] = useState<string | null>(null);

  const factories = useMemo(
    () => Array.from(new Set(rows.map((row) => row.factory))).sort((a, b) => a.localeCompare(b, "th")),
    [rows],
  );
  const weeks = useMemo(
    () => Array.from(new Set(rows.map((row) => normalizeWeek(row.week)).filter(Boolean))).sort((a, b) => Number(b) - Number(a)),
    [rows],
  );
  const groups = useMemo(
    () =>
      Array.from(new Set(rows.map((row) => row.productGroup).filter(Boolean))).sort((a, b) =>
        a.localeCompare(b, "th"),
      ),
    [rows],
  );
  const factoryMetrics = useMemo(
    () =>
      Array.from(new Set(rows.filter((row) => row.tableType === "factory").map((row) => row.metric))).sort((a, b) =>
        a.localeCompare(b, "th"),
      ),
    [rows],
  );
  const productMetrics = useMemo(
    () =>
      Array.from(
        new Set(rows.filter((row) => row.tableType === "product-group").map((row) => row.metric)),
      ).sort((a, b) => a.localeCompare(b, "th")),
    [rows],
  );
  const activeWeekFilters =
    weekFilters.length > 0 && weekFilters.every((week) => weeks.includes(week))
      ? weekFilters
      : weeks[0]
        ? [weeks[0]]
        : [];

  function balanceScore(aiValue: number | null, balanceValue: number | null) {
    if (aiValue === null || balanceValue === null) return null;
    const denominator = Math.max(Math.abs(balanceValue), 1);
    const errorRate = Math.abs(aiValue - balanceValue) / denominator;
    return Math.max(0, Math.round((1 - errorRate) * 100));
  }

  const commonFilteredRows = rows.filter(
    (row) =>
      (factoryFilters.length === 0 || factoryFilters.includes(row.factory)) &&
      (activeWeekFilters.length === 0 || activeWeekFilters.includes(normalizeWeek(row.week))),
  );
  const factoryRows = commonFilteredRows.filter(
    (row) => {
      const status = scoreLabel(balanceScore(row.aiValue, row.balanceValue));
      return (
        row.tableType === "factory" &&
        (factoryMetricFilters.length === 0 || factoryMetricFilters.includes(row.metric)) &&
        (factoryStatusFilters.length === 0 || factoryStatusFilters.includes(status))
      );
    },
  );
  const productRows = dedupeBalanceRows(
    commonFilteredRows.filter((row) => {
      const status = scoreLabel(balanceScore(row.aiValue, row.balanceValue));
      return (
        row.tableType === "product-group" &&
        (groupFilters.length === 0 || groupFilters.includes(row.productGroup)) &&
        (productMetricFilters.length === 0 || productMetricFilters.includes(row.metric)) &&
        (productStatusFilters.length === 0 || productStatusFilters.includes(status))
      );
    }),
  );

  function saveBalanceFeedback() {
    window.localStorage.setItem(storageKey, JSON.stringify(feedback));
    setSavedAt(
      new Intl.DateTimeFormat("th-TH", {
        hour: "2-digit",
        minute: "2-digit",
      }).format(new Date()),
    );
  }

  if (rows.length === 0) {
    return <PlanDataEmptyState />;
  }

  return (
    <section className="comparison-card min-w-0 rounded-xl border border-[#e3e8f0] bg-white p-5 shadow-sm">
      <div className="mb-5 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h2 className="text-xl font-bold">เทียบผล AI กับแผน</h2>
          <p className="mt-1 text-sm text-slate-500">
            เทียบจากไฟล์ {data?.sourceFile ?? "แผน"} กับ {data?.aiFile ?? "AI"} เฉพาะคอลัมน์ที่จับคู่ได้
          </p>
        </div>
        <span className="rounded-md bg-[#ffe8f1] px-3 py-1 text-sm font-bold text-[#ef3e8f]">
          {(factoryRows.length + productRows.length).toLocaleString("th-TH")} รายการ
        </span>
      </div>

      <div className="mb-4 grid gap-3 xl:grid-cols-2">
        <TransferFilterBox label="สัปดาห์" options={weeks} values={activeWeekFilters} onChange={setWeekFilters} />
        <TransferFilterBox label="โรงงาน" options={factories} values={factoryFilters} onChange={setFactoryFilters} />
      </div>

      <div className="space-y-5">
        <BalanceComparisonTable
          title="เทียบรายโรงงาน"
          description="ใช้ชีท 1. ปริมาณตัดแต่ง จากไฟล์ AI เทียบกับคอลัมน์จำนวนตัดแต่งในแผนตามรายชื่อโรงงาน"
          rows={factoryRows}
          feedback={feedback}
          updateFeedback={updateFeedback}
          showProductGroup={false}
          filters={
            <>
              <TransferFilterBox
                label="ตัวแปร"
                options={factoryMetrics}
                values={factoryMetricFilters}
                onChange={setFactoryMetricFilters}
              />
              <TransferFilterBox
                label="สถานะ"
                options={["ดี", "ต่างปานกลาง", "ต่างกันมาก", "รอข้อมูล"]}
                values={factoryStatusFilters}
                onChange={setFactoryStatusFilters}
              />
            </>
          }
        />
        <BalanceComparisonTable
          title="เทียบรายกลุ่มชิ้นส่วน"
          description="ใช้ชีท 2. ปริมาณ Supply และ 3. FC,QT จากไฟล์ AI เทียบตาม ProductForPlan19 / กลุ่มชิ้นส่วน"
          rows={productRows}
          feedback={feedback}
          updateFeedback={updateFeedback}
          showProductGroup
          filters={
            <>
              <TransferFilterBox
                label="กลุ่มชิ้นส่วน"
                options={groups}
                values={groupFilters}
                onChange={setGroupFilters}
              />
              <TransferFilterBox
                label="ตัวแปร"
                options={productMetrics}
                values={productMetricFilters}
                onChange={setProductMetricFilters}
              />
              <TransferFilterBox
                label="สถานะ"
                options={["ดี", "ต่างปานกลาง", "ต่างกันมาก", "รอข้อมูล"]}
                values={productStatusFilters}
                onChange={setProductStatusFilters}
              />
            </>
          }
        />
      </div>

      <div className="mt-4 flex flex-col gap-2 border-t border-[#f5b4cf] pt-4 sm:flex-row sm:items-center sm:justify-end">
        <button
          className="inline-flex h-11 items-center justify-center gap-2 rounded-md bg-[#ef3e8f] px-5 text-sm font-bold text-white shadow-sm hover:bg-[#dc2e81]"
          type="button"
          onClick={saveBalanceFeedback}
        >
          <Save size={18} />
          บันทึกความเห็น
        </button>
        {savedAt ? <p className="text-xs font-bold text-emerald-600">บันทึกแล้ว {savedAt}</p> : null}
      </div>
    </section>
  );
}

function BalanceComparisonTable({
  title,
  description,
  rows,
  feedback,
  updateFeedback,
  showProductGroup,
  filters,
}: {
  title: string;
  description: string;
  rows: BalanceRecord[];
  feedback: Record<string, Feedback>;
  updateFeedback: (id: string, patch: Partial<Feedback>) => void;
  showProductGroup: boolean;
  filters?: ReactNode;
}) {
  function balanceScore(aiValue: number | null, balanceValue: number | null) {
    if (aiValue === null || balanceValue === null) return null;
    const denominator = Math.max(Math.abs(balanceValue), 1);
    const errorRate = Math.abs(aiValue - balanceValue) / denominator;
    return Math.max(0, Math.round((1 - errorRate) * 100));
  }

  return (
    <div className="rounded-xl border border-[#f5b4cf] bg-white/85 p-4">
      <div className="mb-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h3 className="text-lg font-bold">{title}</h3>
          <p className="text-sm text-slate-500">{description}</p>
        </div>
        <span className="rounded-md bg-[#ffe8f1] px-3 py-1 text-sm font-bold text-[#ef3e8f]">
          {rows.length.toLocaleString("th-TH")} รายการ
        </span>
      </div>

      {filters ? <div className="mb-3 grid gap-3 md:grid-cols-2 xl:grid-cols-3">{filters}</div> : null}

      <div className="mobile-table-frame min-w-0 rounded-lg border border-[#dfe6ef]">
        <div className="mobile-table-scroll max-h-[560px] overflow-auto">
          <table className="balance-table w-full min-w-[1280px] border-collapse text-sm">
            <thead className="bg-[#f8fafc] text-xs font-bold text-slate-600">
              <tr>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-left">โรงงาน</th>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-center">สัปดาห์</th>
                {showProductGroup ? (
                  <th className="border-r border-[#e3e8f0] px-4 py-3 text-left">กลุ่มชิ้นส่วน</th>
                ) : null}
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-left">ตัวชี้วัดแผน</th>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-right">AI</th>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-right">แผน</th>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-right">ต่างกัน</th>
                <th className="border-r border-[#e3e8f0] px-4 py-3 text-center">สถานะ</th>
                <th className="px-4 py-3 text-left">ความคิดเห็น</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => {
                const feedbackId = `balance|${row.id}`;
                const itemFeedback = feedback[feedbackId] ?? { actual: "", accuracy: "", comment: "" };
                const matchScore = balanceScore(row.aiValue, row.balanceValue);
                const diff = row.aiValue === null || row.balanceValue === null ? null : row.balanceValue - row.aiValue;
                const commentRequired = matchScore !== null && matchScore < 80;
                const missingRequiredComment = commentRequired && !itemFeedback.comment.trim();

                return (
                  <tr key={row.id} className="border-t border-[#e8edf4]">
                    <td className="border-r border-[#e8edf4] px-4 py-3 font-medium">{row.factory}</td>
                    <td className="border-r border-[#e8edf4] px-4 py-3 text-center font-mono">{row.week}</td>
                    {showProductGroup ? (
                      <td className="border-r border-[#e8edf4] px-4 py-3 font-medium">{row.productGroup}</td>
                    ) : null}
                    <td className="border-r border-[#e8edf4] px-4 py-3">{formatBalanceMetric(row.metric)}</td>
                    <td className="border-r border-[#e8edf4] px-4 py-3 text-right font-mono">
                      {formatNumber(row.aiValue)}
                    </td>
                    <td className="border-r border-[#e8edf4] px-4 py-3 text-right font-mono">
                      {formatNumber(row.balanceValue)}
                    </td>
                    <td
                      className={`border-r border-[#e8edf4] px-4 py-3 text-right font-mono ${
                        diff !== null && diff > 0 ? "text-red-600" : "text-emerald-600"
                      }`}
                    >
                      {diff === null ? "-" : `${diff > 0 ? "+" : ""}${formatNumber(diff)}`}
                    </td>
                    <td className="border-r border-[#e8edf4] px-4 py-3 text-center">
                      <span
                        className={`inline-flex h-8 min-w-24 items-center justify-center rounded-md px-3 text-xs font-bold ${scoreTone(
                          matchScore,
                        )}`}
                      >
                        {scoreLabel(matchScore)}
                      </span>
                    </td>
                    <td className="px-4 py-3">
                      <div className="relative">
                        <input
                          className={`h-9 w-full rounded-md border px-3 pr-9 text-sm outline-none ${
                            missingRequiredComment
                              ? "border-red-400 bg-red-50 focus:border-red-500"
                              : "border-[#dfe6ef] focus:border-[#ef3e8f]"
                          }`}
                          placeholder="กรอกความคิดเห็น..."
                          value={itemFeedback.comment}
                          onChange={(event) => updateFeedback(feedbackId, { comment: event.target.value })}
                        />
                        {commentRequired ? (
                          <span className="absolute right-2 top-1/2 -translate-y-1/2 text-lg font-bold text-red-500">
                            *
                          </span>
                        ) : null}
                      </div>
                    </td>
                  </tr>
                );
              })}
              {rows.length === 0 ? (
                <tr>
                  <td className="px-4 py-8 text-center text-sm text-slate-500" colSpan={showProductGroup ? 9 : 8}>
                    ไม่พบข้อมูลตามตัวกรอง
                  </td>
                </tr>
              ) : null}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function dedupeBalanceRows(rows: BalanceRecord[]) {
  const preferredRows = new Map<string, BalanceRecord>();

  rows.forEach((row) => {
    const key = [row.tableType, row.factory, row.week, row.productGroup, row.metric].join("|");
    const current = preferredRows.get(key);

    if (!current || balanceRowPriority(row) > balanceRowPriority(current)) {
      preferredRows.set(key, row);
    }
  });

  return Array.from(preferredRows.values());
}

function formatBalanceMetric(metric: string) {
  return /\((kg|head|%|ตัว\/สัปดาห์)\)/i.test(metric) || metric.includes("%") ? metric : `${metric} (Kg)`;
}

function balanceRowPriority(row: BalanceRecord) {
  const metric = row.metric.toLowerCase();
  const isFcQtMetric =
    metric.includes("fc") ||
    metric.includes("qt") ||
    metric.includes("ขาด") ||
    metric.includes("เหลือ");

  if (isFcQtMetric && row.sourceSheet.includes("3.")) return 3;
  if (!isFcQtMetric && row.sourceSheet.includes("2.")) return 2;
  return 1;
}

function TransferComparison({
  records,
  actuals,
  feedback,
  updateFeedback,
}: {
  records: TransferRecord[];
  actuals: TransferActualPayload;
  feedback: Record<string, Feedback>;
  updateFeedback: (id: string, patch: Partial<Feedback>) => void;
}) {
  const [sourceFilters, setSourceFilters] = useState<string[]>([]);
  const [destinationFilters, setDestinationFilters] = useState<string[]>([]);
  const [groupFilters, setGroupFilters] = useState<string[]>([]);
  const [typeFilters, setTypeFilters] = useState<string[]>([]);
  const [weekFilters, setWeekFilters] = useState<string[]>([]);
  const [transferStatusFilters, setTransferStatusFilters] = useState<string[]>([
    "ต่างปานกลาง",
    "ต่างกันมาก",
  ]);
  const rowsWithActual = useMemo(
    () =>
      records.map((record) => {
        const actual = actuals[transferKey(record.source, record.destination, record.productGroup)];
        return {
          ...record,
          actualTransfer: actual?.actualTransfer ?? 0,
          actualProductTypes: actual?.productTypes ?? [],
        };
      }),
    [actuals, records],
  );
  const sources = useMemo(
    () => Array.from(new Set(rowsWithActual.map((record) => record.source))).sort((a, b) => a.localeCompare(b, "th")),
    [rowsWithActual],
  );
  const destinations = useMemo(
    () =>
      Array.from(new Set(rowsWithActual.map((record) => record.destination))).sort((a, b) =>
        a.localeCompare(b, "th"),
      ),
    [rowsWithActual],
  );
  const productGroups = useMemo(
    () =>
      Array.from(new Set(rowsWithActual.map((record) => record.productGroup))).sort((a, b) =>
        a.localeCompare(b, "th"),
      ),
    [rowsWithActual],
  );
  const productTypes = useMemo(
    () =>
      Array.from(
        new Set(rowsWithActual.flatMap((record) => [record.productType, ...record.actualProductTypes])),
      ).sort((a, b) => a.localeCompare(b, "th")),
    [rowsWithActual],
  );
  const weeks = useMemo(
    () =>
      Array.from(
        new Set(rowsWithActual.flatMap((record) => record.weeks.map((item) => normalizeWeek(item)).filter(Boolean))),
      ).sort(
        (a, b) => Number(b) - Number(a),
      ),
    [rowsWithActual],
  );
  const activeWeekFilters =
    weekFilters.length > 0 && weekFilters.every((week) => weeks.includes(week))
      ? weekFilters
      : weeks[0]
        ? [weeks[0]]
        : [];
  const filteredRows = rowsWithActual.filter((record) => {
    const matchesSource = sourceFilters.length === 0 || sourceFilters.includes(record.source);
    const matchesDestination =
      destinationFilters.length === 0 || destinationFilters.includes(record.destination);
    const matchesGroup = groupFilters.length === 0 || groupFilters.includes(record.productGroup);
    const matchesType =
      typeFilters.length === 0 ||
      typeFilters.includes(record.productType) ||
      record.actualProductTypes.some((type) => typeFilters.includes(type));
    const matchesWeek =
      activeWeekFilters.length === 0 ||
      record.weeks.some((week) => activeWeekFilters.includes(normalizeWeek(week)));
    const recordStatus = scoreLabel(
      transferScore(record.aiTransfer, record.actualTransfer),
    );
    const matchesStatus =
      transferStatusFilters.length === 0 || transferStatusFilters.includes(recordStatus);
    return matchesSource && matchesDestination && matchesGroup && matchesType && matchesWeek && matchesStatus;
  });

  function transferScore(aiValue: number | null, actualValue: number | null) {
    if (aiValue === null || actualValue === null) return null;
    const denominator = Math.max(Math.abs(actualValue), 1);
    const errorRate = Math.abs(aiValue - actualValue) / denominator;
    return Math.max(0, Math.round((1 - errorRate) * 100));
  }

  return (
    <section className="space-y-5">
      <div className="rounded-xl border border-[#e3e8f0] bg-white p-5 shadow-sm">
        <div className="mb-5 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <h2 className="text-xl font-bold">เทียบผลการโอนจาก AI กับ Actual</h2>
            <p className="mt-1 text-sm text-slate-500">
              ใช้ผล AI จากชีต 4-5 และเทียบค่าจริงจาก Actual ชีต 4. โอน
            </p>
          </div>
          <span className="rounded-md bg-[#ffe8f1] px-3 py-1 text-sm font-bold text-[#ef3e8f]">
            {filteredRows.length.toLocaleString("th-TH")} รายการ
          </span>
        </div>

        <div className="mb-4 grid gap-3 xl:grid-cols-6">
          <TransferFilterBox
            label="สัปดาห์"
            options={weeks}
            values={activeWeekFilters}
            onChange={setWeekFilters}
          />
          <TransferFilterBox label="ต้นทาง" options={sources} values={sourceFilters} onChange={setSourceFilters} />
          <TransferFilterBox
            label="ปลายทาง"
            options={destinations}
            values={destinationFilters}
            onChange={setDestinationFilters}
          />
          <TransferFilterBox
            label="กลุ่มสินค้า"
            options={productGroups}
            values={groupFilters}
            onChange={setGroupFilters}
          />
          <TransferFilterBox
            label="ชนิดสินค้า"
            options={productTypes}
            values={typeFilters}
            onChange={setTypeFilters}
          />
          <TransferFilterBox
            label="สถานะ"
            options={["ดี", "ต่างปานกลาง", "ต่างกันมาก", "รอข้อมูล"]}
            values={transferStatusFilters}
            onChange={setTransferStatusFilters}
          />
        </div>

        <div className="mobile-table-frame rounded-lg border border-[#dfe6ef]">
          <div className="mobile-table-scroll max-h-[680px] overflow-auto">
            <table className="transfer-table w-full min-w-[1080px] border-collapse text-sm">
              <colgroup>
                <col />
                <col />
                <col />
                <col className="w-48" />
                <col />
                <col />
                <col />
                <col />
                <col />
              </colgroup>
              <thead className="bg-[#f8fafc] text-xs font-bold text-slate-600">
                <tr>
                  <th className="border-r border-[#e3e8f0] px-4 py-3 text-left">ต้นทาง</th>
                  <th className="border-r border-[#e3e8f0] px-4 py-3 text-left">ปลายทาง</th>
                  <th className="border-r border-[#e3e8f0] px-4 py-3 text-left">กลุ่มสินค้า</th>
                  <th className="border-r border-[#e3e8f0] px-4 py-3 text-left">ชนิดสินค้า</th>
                  <th className="border-r border-[#e3e8f0] px-4 py-3 text-right">AI แนะนำโอน (kg)</th>
                  <th className="border-r border-[#e3e8f0] px-4 py-3 text-right">Actual โอน (kg)</th>
                  <th className="border-r border-[#e3e8f0] px-4 py-3 text-right">ต่างกัน (kg)</th>
                  <th className="border-r border-[#e3e8f0] px-4 py-3 text-center">สถานะ</th>
                  <th className="px-4 py-3">ความคิดเห็น</th>
                </tr>
              </thead>
              <tbody>
                {filteredRows.map((record) => {
                  const feedbackId = `transfer|${record.id}`;
                  const itemFeedback = feedback[feedbackId] ?? {
                    actual: "",
                    accuracy: "",
                    comment: "",
                  };
                  const actualValue = record.actualTransfer;
                  const matchScore = transferScore(record.aiTransfer, actualValue);
                  const diff =
                    record.aiTransfer === null || actualValue === null
                      ? null
                      : actualValue - record.aiTransfer;
                  const commentRequired = matchScore !== null && matchScore < 80;
                  const missingRequiredComment =
                    commentRequired && !itemFeedback.comment.trim();

                  return (
                    <tr key={record.id} className="border-t border-[#e8edf4]">
                      <td className="border-r border-[#e8edf4] px-4 py-3 font-medium">
                        {record.source}
                      </td>
                      <td className="border-r border-[#e8edf4] px-4 py-3 font-medium">
                        {record.destination}
                      </td>
                      <td className="border-r border-[#e8edf4] px-4 py-3 font-medium">
                        {record.productGroup}
                      </td>
                      <td className="max-w-48 border-r border-[#e8edf4] px-3 py-3">
                        <p className="truncate font-medium">{record.productType}</p>
                        {record.actualProductTypes.length > 0 ? (
                          <p className="mt-1 line-clamp-1 text-xs text-slate-500">
                            Actual: {record.actualProductTypes.join(", ")}
                          </p>
                        ) : null}
                      </td>
                      <td className="border-r border-[#e8edf4] px-4 py-3 text-right font-mono">
                        {formatNumber(record.aiTransfer)}
                      </td>
                      <td className="border-r border-[#e8edf4] px-4 py-3 text-right font-mono">
                        {formatNumber(actualValue)}
                      </td>
                      <td
                        className={`border-r border-[#e8edf4] px-4 py-3 text-right font-mono ${
                          diff !== null && diff > 0 ? "text-red-600" : "text-emerald-600"
                        }`}
                      >
                        {diff === null ? "-" : `${diff > 0 ? "+" : ""}${formatNumber(diff)}`}
                      </td>
                      <td className="border-r border-[#e8edf4] px-4 py-3 text-center">
                        <span
                          className={`inline-flex h-8 min-w-24 items-center justify-center rounded-md px-3 text-xs font-bold ${scoreTone(
                            matchScore,
                          )}`}
                        >
                          {scoreLabel(matchScore)}
                        </span>
                      </td>
                      <td className="px-4 py-3">
                        <div className="relative">
                          <input
                            className={`h-9 w-full rounded-md border px-3 pr-9 text-sm outline-none ${
                              missingRequiredComment
                                ? "border-red-400 bg-red-50 focus:border-red-500"
                                : "border-[#dfe6ef] focus:border-[#ef4b98]"
                            }`}
                            placeholder={
                              commentRequired
                                ? "ต้องกรอกความคิดเห็น..."
                                : "กรอกความคิดเห็น..."
                            }
                            value={itemFeedback.comment}
                            onChange={(event) =>
                              updateFeedback(feedbackId, {
                                comment: event.target.value,
                              })
                            }
                          />
                          {commentRequired && (
                            <span className="absolute right-2 top-1/2 -translate-y-1/2 text-lg font-bold text-red-500">
                              *
                            </span>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </section>
  );
}

function TransferFilterBox({
  label,
  options,
  values,
  onChange,
}: {
  label: string;
  options: string[];
  values: string[];
  onChange: (values: string[]) => void;
}) {
  const [search, setSearch] = useState("");
  const filteredOptions = options.filter((option) =>
    option.toLowerCase().includes(search.trim().toLowerCase()),
  );

  function toggle(option: string) {
    onChange(
      values.includes(option)
        ? values.filter((value) => value !== option)
        : [...values, option],
    );
  }

  return (
    <details className="relative rounded-md border border-[#dfe6ef] bg-white">
      <summary className="flex h-11 cursor-pointer list-none items-center justify-between gap-2 px-3 text-sm font-medium marker:hidden">
        <span className="min-w-0 truncate">
          {label}:{" "}
          {values.length === 0
            ? "ทั้งหมด"
            : values.length === 1
              ? values[0]
              : `${values.length} รายการ`}
        </span>
        <ChevronDown size={16} className="shrink-0 text-slate-400" />
      </summary>
      <div className="absolute left-0 top-12 z-30 w-full rounded-md border border-[#dfe6ef] bg-white p-3 shadow-lg">
        <div className="mb-2 flex items-center justify-between gap-2">
          <p className="text-xs font-bold text-slate-600">{label}</p>
          <button className="text-xs font-bold text-[#ef3e8f]" type="button" onClick={() => onChange([])}>
            ทั้งหมด
          </button>
        </div>
        <label className="relative mb-3 block">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
          <input
            className="h-10 w-full rounded-md border border-[#dfe6ef] pl-9 pr-3 text-sm outline-none focus:border-[#ef4b98]"
            placeholder={`ค้นหา${label}...`}
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>
        <FilterGroup
          options={filteredOptions}
          values={values}
          onToggle={toggle}
          emptyText={`ไม่พบ${label}`}
        />
      </div>
    </details>
  );
}

function UploadAiPanel({
  data,
  isUploading,
  status,
  uploadedNames,
  uploadHistory,
  onUpload,
  onActualUpload,
  onBalanceUpload,
}: {
  data: AiData | null;
  isUploading: boolean;
  status: { tone: "success" | "error"; message: string } | null;
  uploadedNames: { ai?: string; actual?: string; balance?: string };
  uploadHistory: UploadHistoryItem[];
  onUpload: (file: File) => Promise<void>;
  onActualUpload: (file: File) => Promise<void>;
  onBalanceUpload: (file: File) => Promise<void>;
}) {
  const sheets = Array.from(new Set(data?.records.map((record) => record.sheet) ?? []));
  const factories = Array.from(new Set(data?.records.map((record) => record.factory) ?? []));
  const metrics = Array.from(new Set(data?.records.map((record) => record.metric) ?? []));

  return (
    <section className="grid gap-5 xl:grid-cols-[1.25fr_0.75fr]">
      <Panel title="นำเข้าไฟล์เพื่อเปรียบเทียบ">
        <div className="grid gap-5 xl:grid-cols-3">
          <UploadBox
            title="1. ไฟล์ผลลัพธ์ AI"
            description="ไฟล์ sigmas หรือไฟล์ผล AI ไม่จำเป็นต้องมีสีเหลืองแล้ว แต่หัวคอลัมน์ต้องตรงกับหัวข้อ feedback ที่ระบบจำไว้"
            buttonLabel={isUploading ? "กำลังอ่านไฟล์..." : "เลือกไฟล์ผล AI"}
            fileName={uploadedNames.ai ?? data?.sourceFile}
            disabled={isUploading}
            onUpload={onUpload}
          />
          <UploadBox
            title="2. ไฟล์ค่าจริง Actual"
            description="ไฟล์ Actual.xlsx ต้องอัปโหลดช่องนี้ ระบบจะจับคู่กับผล AI แล้วเติมค่า Actual ลงตารางอัตโนมัติ"
            buttonLabel={isUploading ? "กำลังอ่านไฟล์..." : "เลือกไฟล์ Actual"}
            fileName={uploadedNames.actual}
            disabled={isUploading || !data?.records.length}
            onUpload={onActualUpload}
          />
          <UploadBox
            title="3. ไฟล์แผน"
            description="อัปโหลดไฟล์แผน สำหรับใช้เทียบผลเชิง Business กับผลลัพธ์ AI และใช้เป็นฐานของแท็บวิเคราะห์ผล"
            buttonLabel={isUploading ? "กำลังอ่านไฟล์..." : "เลือกไฟล์แผน"}
            fileName={uploadedNames.balance}
            disabled={isUploading}
            onUpload={onBalanceUpload}
          />
        </div>

        <div className="mt-5 rounded-lg border border-blue-100 bg-blue-50 p-4 text-sm leading-6 text-blue-800">
          <p className="font-bold">ลำดับการใช้งาน</p>
          <p className="mt-1">
            ระบบจำหัวข้อจากคอลัมน์สีเหลืองชุดแรกไว้แล้ว ไฟล์ใหม่ไม่ต้องไฮไลต์สีเหลืองอีก
            ให้อัปโหลดไฟล์ผล AI ก่อน จากนั้นอัปโหลดไฟล์ Actual ระบบจะใช้ชื่อชีต โรงงาน และ mapping
            ของคอลัมน์ เช่น Production (kg) ↔ ProductionWeight, Quota (kg) ↔ Quota เพื่อเติมค่าจริง
            ถ้าเลือกไฟล์ Actual ผิดช่อง ระบบจะพยายามส่งต่อไปประมวลผลแบบ Actual ให้อัตโนมัติ
          </p>
        </div>

        {status && (
          <div
            className={`mt-5 flex items-center gap-3 rounded-md border p-4 text-left text-sm ${
              status.tone === "success"
                ? "border-emerald-200 bg-emerald-50 text-emerald-800"
                : "border-rose-200 bg-rose-50 text-rose-800"
            }`}
          >
            {status.tone === "success" ? (
              <CheckCircle2 className="shrink-0" size={20} />
            ) : (
              <AlertCircle className="shrink-0" size={20} />
            )}
            {status.message}
          </div>
        )}
      </Panel>

      <div className="space-y-5">
        <Panel title="ข้อมูลชุดที่ใช้อยู่">
          <InfoRow icon={<Database size={22} />} label="ไฟล์ปัจจุบัน" value={data?.sourceFile ?? "-"} />
          <InfoRow
            icon={<Target size={22} />}
            label="จำนวนรายการ"
            value={`${(data?.records.length ?? 0).toLocaleString("th-TH")} รายการ`}
          />
          <InfoRow icon={<Factory size={22} />} label="จำนวนโรงงาน" value={`${factories.length}`} />
          <InfoRow icon={<BarChart3 size={22} />} label="จำนวนชีต" value={`${sheets.length}`} />
        </Panel>
        <Panel title="ประวัติการอัปโหลดไฟล์">
          <div className="max-h-[280px] space-y-2 overflow-auto pr-1">
            {uploadHistory.map((item, index) => (
              <div
                key={`${item.type}-${item.name}-${index}`}
                className="rounded-lg border border-[#f5b4cf] bg-white/80 px-3 py-2 text-sm"
              >
                <div className="flex items-center justify-between gap-3">
                  <span className="rounded-md bg-[#ffe8f1] px-2 py-1 text-xs font-bold text-[#ef3e8f]">
                    {item.type}
                  </span>
                  <span className="text-xs text-slate-500">{item.uploadedAt}</span>
                </div>
                <p className="mt-2 truncate font-bold text-slate-700">{item.name}</p>
              </div>
            ))}
            {uploadHistory.length === 0 ? (
              <div className="rounded-lg border border-dashed border-[#dfe6ef] p-4 text-center text-sm text-slate-500">
                ยังไม่มีประวัติการอัปโหลด
              </div>
            ) : null}
          </div>
        </Panel>
        <Panel title="หัวข้อ Feedback ที่ระบบจำไว้">
          <div className="max-h-[360px] space-y-2 overflow-auto pr-1">
            {metrics.slice(0, 18).map((metric) => (
              <div
                key={metric}
                className="rounded-md border border-[#f4dd98] bg-[#fff8da] px-3 py-2 text-sm font-medium text-slate-700"
              >
                {metric}
              </div>
            ))}
          </div>
        </Panel>
      </div>
    </section>
  );
}

function UploadBox({
  title,
  description,
  buttonLabel,
  fileName,
  disabled,
  onUpload,
}: {
  title: string;
  description: string;
  buttonLabel: string;
  fileName?: string;
  disabled: boolean;
  onUpload: (file: File) => Promise<void>;
}) {
  return (
    <div className="rounded-lg border border-dashed border-[#ff9ac3] bg-[#fff7fb] p-6 text-center">
      <div className="mx-auto grid size-14 place-items-center rounded-full bg-white text-[#ef3e8f] shadow-sm">
        <Upload size={26} />
      </div>
      <h3 className="mt-4 text-xl font-bold">{title}</h3>
      <p className="mx-auto mt-3 max-w-md text-sm leading-6 text-slate-600">{description}</p>
      <label
        className={`mx-auto mt-5 flex h-11 w-full max-w-xs items-center justify-center gap-2 rounded-md px-5 text-sm font-bold shadow-sm ${
          disabled
            ? "cursor-not-allowed bg-slate-200 text-slate-500"
            : "cursor-pointer bg-[#ef3e8f] text-white hover:bg-[#dc2e81]"
        }`}
      >
        <Upload size={18} />
        {buttonLabel}
        <input
          type="file"
          accept=".xlsx,.xlsm"
          className="sr-only"
          disabled={disabled}
          onChange={(event) => {
            const file = event.target.files?.[0];
            if (file) void onUpload(file);
            event.currentTarget.value = "";
          }}
        />
      </label>
      <p className="mt-3 min-h-5 truncate text-xs font-medium text-slate-500">
        {fileName ? `ไฟล์: ${fileName}` : "ยังไม่ได้เลือกไฟล์"}
      </p>
    </div>
  );
}

function Panel({ title, children }: { title: string; children: ReactNode }) {
  return (
    <div className="rounded-xl border border-[#e1e7ef] bg-white p-5 shadow-sm">
      <div className="mb-4 flex items-center justify-between">
        <h3 className="text-lg font-bold">{title}</h3>
        <ChevronDown className="text-slate-400" size={18} />
      </div>
      {children}
    </div>
  );
}

function InfoRow({
  icon,
  label,
  value,
}: {
  icon: ReactNode;
  label: string;
  value: string;
}) {
  return (
    <div className="flex gap-3 border-b border-slate-100 py-3 last:border-b-0">
      <div className="text-slate-500">{icon}</div>
      <div className="min-w-0">
        <p className="text-sm text-slate-500">{label}</p>
        <p className="truncate font-medium">{value}</p>
      </div>
    </div>
  );
}
