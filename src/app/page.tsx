"use client";

import {
  AlertCircle,
  BarChart3,
  Brain,
  CalendarDays,
  CheckCircle2,
  ChevronDown,
  Database,
  Edit3,
  Factory,
  Gauge,
  HelpCircle,
  Home as HomeIcon,
  LogOut,
  Menu,
  MessageCircle,
  Paperclip,
  Save,
  Search,
  Settings,
  Target,
  Upload,
  User,
  Users,
} from "lucide-react";
import { type ReactNode, useEffect, useMemo, useState } from "react";

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

type AiData = {
  generatedAt: string;
  sourceFile: string;
  records: AiRecord[];
};

type Feedback = {
  actual: string;
  accuracy: string;
  comment: string;
};

type AppTab = "feedback" | "upload";

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

const allSheets = "ทั้งหมด";
const allFactories = "ทุกโรงเรือน";
const storageKey = "ai-feedback-review-v1";
const uploadedDataKey = "ai-feedback-uploaded-ai-data-v2";
const uploadedActualKey = "ai-feedback-uploaded-actual-feedback-v2";
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

const navItems = [
  { label: "Dashboard", icon: HomeIcon, tab: "feedback" },
  { label: "บันทึก Feedback", icon: Edit3, tab: "feedback" },
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
  if (value >= 60) return "ควรปรับปรุง";
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

async function parseAiWorkbook(file: File): Promise<AiData> {
  const ExcelJS = await import("exceljs");
  const workbook = new ExcelJS.Workbook() as unknown as WorkbookLike;
  await workbook.xlsx.load(await file.arrayBuffer());

  const records: AiRecord[] = [];

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

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;

      const factory = factoryFromRow(row, headers);
      const week = weekCol ? cellText(row.getCell(weekCol).value).trim() : "";

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

  return {
    generatedAt: new Date().toISOString(),
    sourceFile: file.name,
    records,
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

export default function Home() {
  const [data, setData] = useState<AiData | null>(null);
  const [sheet, setSheet] = useState(allSheets);
  const [factory, setFactory] = useState(allFactories);
  const [query, setQuery] = useState("");
  const [activeTab, setActiveTab] = useState<AppTab>("feedback");
  const [isUploading, setIsUploading] = useState(false);
  const [uploadedNames, setUploadedNames] = useState<{ ai?: string; actual?: string }>({});
  const [uploadStatus, setUploadStatus] = useState<{
    tone: "success" | "error";
    message: string;
  } | null>(null);
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
        setData(uploaded ? JSON.parse(uploaded) : defaultData);
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
  }, []);

  useEffect(() => {
    window.localStorage.setItem(storageKey, JSON.stringify(feedback));
  }, [feedback]);

  const records = useMemo(() => data?.records ?? [], [data]);
  const sheets = useMemo(
    () => [allSheets, ...Array.from(new Set(records.map((record) => record.sheet)))],
    [records],
  );
  const factories = useMemo(
    () => [
      allFactories,
      ...Array.from(new Set(records.map((record) => record.factory))).sort((a, b) =>
        a.localeCompare(b, "th"),
      ),
    ],
    [records],
  );

  const filtered = useMemo(() => {
    const normalizedQuery = query.trim().toLowerCase();
    return records.filter((record) => {
      const matchesSheet = sheet === allSheets || record.sheet === sheet;
      const matchesFactory = factory === allFactories || record.factory === factory;
      const matchesQuery =
        !normalizedQuery ||
        `${record.metric} ${record.factory} ${record.sheet}`
          .toLowerCase()
          .includes(normalizedQuery);

      return matchesSheet && matchesFactory && matchesQuery;
    });
  }, [factory, query, records, sheet]);

  const tableRows = filtered;
  const scores = tableRows
    .map((record) => score(record, feedback[record.id]?.actual ?? ""))
    .filter((value): value is number => value !== null);
  const averageScore =
    scores.length > 0
      ? Math.round(scores.reduce((total, value) => total + value, 0) / scores.length)
      : null;
  const goodCount = scores.filter((value) => value >= 80).length;
  const warningCount = scores.filter((value) => value >= 60 && value < 80).length;
  const badCount = scores.filter((value) => value < 60).length;

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

  async function handleAiUpload(file: File) {
    if (file.name.toLowerCase().includes("actual")) {
      await handleActualUpload(file);
      return;
    }

    setIsUploading(true);
    setUploadStatus(null);

    try {
      const uploadedData = await parseAiWorkbook(file);
      setData(uploadedData);
      setUploadedNames((current) => ({ ...current, ai: file.name }));
      setSheet(allSheets);
      setFactory(allFactories);
      setQuery("");
      window.localStorage.setItem(uploadedDataKey, JSON.stringify(uploadedData));
      setUploadStatus({
        tone: "success",
        message: `อัปโหลดสำเร็จ: โหลด ${uploadedData.records.length.toLocaleString(
          "th-TH",
        )} รายการจากหัวข้อ feedback ที่ระบบจำไว้`,
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
      const actualFeedback = await parseActualWorkbook(file, data.records);
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
      window.localStorage.setItem(uploadedActualKey, JSON.stringify(actualFeedback));
      setUploadedNames((current) => ({ ...current, actual: file.name }));
      setUploadStatus({
        tone: "success",
        message: `เติมค่าจริงสำเร็จ: จับคู่ได้ ${Object.keys(actualFeedback).length.toLocaleString(
          "th-TH",
        )} รายการ`,
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

  return (
    <main className="min-h-screen bg-[#f6f8fb] text-[#172033]">
      <div className="grid min-h-screen lg:grid-cols-[260px_1fr]">
        <Sidebar activeTab={activeTab} setActiveTab={setActiveTab} />

        <section className="min-w-0">
          <header className="sticky top-0 z-20 border-b border-[#e2e8f0] bg-white/95 backdrop-blur">
            <div className="flex min-h-24 items-center justify-between gap-4 px-5 py-4 sm:px-8">
              <div className="flex items-center gap-5">
                <button className="grid size-11 place-items-center rounded-md text-slate-600 hover:bg-slate-100">
                  <Menu size={28} />
                </button>
                <div>
                  <h1 className="text-2xl font-bold">
                    {activeTab === "upload"
                      ? "อัปโหลดผล AI"
                      : "บันทึก Feedback: เทียบผล AI กับค่าจริง"}
                  </h1>
                  <p className="mt-1 text-sm text-slate-500">
                    {activeTab === "upload"
                      ? "นำเข้าไฟล์ Excel ที่มีโครงสร้างชีตและหัวคอลัมน์แบบเดิม"
                      : "กรอกและตรวจสอบความถูกต้องของผลลัพธ์ AI เทียบกับค่าจริง"}
                  </p>
                </div>
              </div>

              <div className="hidden items-center gap-3 xl:flex">
                <TopButton>
                  31/05/2024
                  <CalendarDays size={17} />
                </TopButton>
                <select
                  className="h-11 min-w-48 rounded-md border border-[#dfe6ef] bg-white px-4 text-sm font-medium shadow-sm"
                  value={sheet}
                  onChange={(event) => setSheet(event.target.value)}
                >
                  {sheets.map((option) => (
                    <option key={option}>{option}</option>
                  ))}
                </select>
                <select
                  className="h-11 min-w-56 rounded-md border border-[#dfe6ef] bg-white px-4 text-sm font-medium shadow-sm"
                  value={factory}
                  onChange={(event) => setFactory(event.target.value)}
                >
                  {factories.map((option) => (
                    <option key={option}>{option}</option>
                  ))}
                </select>
                <TopButton>
                  <HelpCircle size={17} />
                  วิธีใช้งาน
                </TopButton>
              </div>
            </div>
          </header>

          <div className="space-y-5 px-5 py-5 sm:px-8">
            {activeTab === "upload" ? (
              <UploadAiPanel
                data={data}
                isUploading={isUploading}
                status={uploadStatus}
                uploadedNames={uploadedNames}
                onUpload={handleAiUpload}
                onActualUpload={handleActualUpload}
              />
            ) : (
              <>
                <ContextBar
                  factory={factory === allFactories ? "โรงงาน A" : factory}
                  sourceFile={data?.sourceFile ?? "-"}
                />

                <section className="grid gap-5 xl:grid-cols-[1fr_300px]">
                  <ComparisonCard
                    query={query}
                    setQuery={setQuery}
                    rows={tableRows}
                    feedback={feedback}
                    updateFeedback={updateFeedback}
                  />
                  <SummaryCard
                    score={averageScore}
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
  return (
    <aside className="hidden border-r border-[#e3e8f0] bg-white lg:flex lg:flex-col">
      <div className="flex h-24 items-center gap-3 border-b border-[#edf1f6] px-6">
        <div className="relative grid size-12 place-items-center rounded-full bg-[#ffe4ef] text-sm font-black text-[#ef3e8f]">
          <span className="absolute -left-1 top-1 size-4 rounded-full bg-[#ffc6dc]" />
          <span className="absolute -right-1 top-1 size-4 rounded-full bg-[#ffc6dc]" />
          PF
        </div>
        <div>
          <h1 className="text-xl font-bold tracking-tight">PigFactory AI</h1>
          <p className="text-sm text-slate-500">AI Feedback System</p>
        </div>
      </div>

      <nav className="space-y-2 px-3 py-5">
        {navItems.map((item) => {
          const Icon = item.icon;
          const active =
            activeTab === "upload"
              ? "tab" in item && item.tab === "upload"
              : item.label === "บันทึก Feedback";

          return (
            <button
              key={item.label}
              onClick={() => {
                if ("tab" in item) setActiveTab(item.tab);
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

function TopButton({ children }: { children: ReactNode }) {
  return (
    <button className="flex h-11 items-center gap-3 rounded-md border border-[#dfe6ef] bg-white px-4 text-sm font-medium shadow-sm">
      {children}
    </button>
  );
}

function ContextBar({ factory, sourceFile }: { factory: string; sourceFile: string }) {
  return (
    <section className="rounded-xl border border-[#e3e8f0] bg-white p-5 shadow-sm">
      <div className="grid gap-4 lg:grid-cols-[220px_1fr_1fr_1.7fr_1.2fr]">
        <InfoTile icon={<Factory size={28} />} label="โรงงาน" value={factory} />
        <InfoTile label="วันที่บันทึก" value="31/05/2024" />
        <InfoTile label="รอบการผลิต" value="รอบที่ 25/2024" />
        <InfoTile label="ช่วงเวลาการทำนาย" value="30/05/2024 20:00 - 31/05/2024 20:00" />
        <InfoTile label="ไฟล์ผล AI" value={sourceFile} />
      </div>
    </section>
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
        <p className="mt-1 truncate font-semibold">{value}</p>
      </div>
    </div>
  );
}

function ComparisonCard({
  query,
  setQuery,
  rows,
  feedback,
  updateFeedback,
}: {
  query: string;
  setQuery: (query: string) => void;
  rows: AiRecord[];
  feedback: Record<string, Feedback>;
  updateFeedback: (id: string, patch: Partial<Feedback>) => void;
}) {
  return (
    <div className="rounded-xl border border-[#e3e8f0] bg-white p-5 shadow-sm">
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
          <Legend color="bg-orange-500" label="ควรปรับปรุง" />
          <Legend color="bg-red-500" label="ต่างกันมาก" />
        </div>
      </div>

      <label className="relative mb-4 block max-w-md">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
        <input
          className="h-11 w-full rounded-md border border-[#dfe6ef] pl-10 pr-3 text-sm outline-none focus:border-[#ef4b98]"
          placeholder="ค้นหาตัวชี้วัด"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
        />
      </label>

      <div className="overflow-hidden rounded-lg border border-[#dfe6ef]">
        <div className="max-h-[640px] overflow-auto">
          <table className="w-full min-w-[1080px] border-collapse text-sm">
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

                return (
                  <tr key={record.id} className="border-t border-[#e8edf4]">
                    <td className="border-r border-[#e8edf4] px-4 py-3">
                      <div className="flex items-center gap-3">
                        <div className="grid size-8 shrink-0 place-items-center rounded-md bg-slate-100 text-slate-500">
                          <Gauge size={18} />
                        </div>
                        <div className="min-w-0">
                          <p className="truncate font-semibold">{record.metric}</p>
                          <p className="truncate text-xs text-slate-500">{record.factory}</p>
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
                          className="h-9 w-full rounded-md border border-[#dfe6ef] px-3 pr-9 text-sm outline-none focus:border-[#ef4b98]"
                          placeholder="กรอกความคิดเห็น..."
                          value={itemFeedback.comment}
                          onChange={(event) =>
                            updateFeedback(record.id, { comment: event.target.value })
                          }
                        />
                        <MessageCircle
                          className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400"
                          size={17}
                        />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
      <p className="mt-3 text-xs text-slate-500">* % ต่างกัน = |ผลต่าง| / AI ทำนาย × 100</p>
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

function SummaryCard({
  score: averageScore,
  goodCount,
  warningCount,
  badCount,
}: {
  score: number | null;
  goodCount: number;
  warningCount: number;
  badCount: number;
}) {
  return (
    <div className="space-y-5">
      <div className="rounded-xl border border-[#e3e8f0] bg-white p-5 shadow-sm">
        <h3 className="text-lg font-bold">สรุปภาพรวมรอบการผลิตนี้</h3>
        <div className="mt-6 flex justify-center">
          <div className="relative grid size-40 place-items-center rounded-full bg-[conic-gradient(#39b87f_0_72%,#e8edf4_72%_100%)]">
            <div className="grid size-28 place-items-center rounded-full bg-white shadow-inner">
              <div className="text-center">
                <p className="text-3xl font-bold">
                  {averageScore === null ? "0.88" : `0.${averageScore}`}
                </p>
                <p className="text-sm font-bold text-emerald-600">ดี</p>
              </div>
            </div>
          </div>
        </div>
        <div className="mt-6 space-y-3">
          <SummaryPill tone="good" label="ดี (ต่างกันน้อย)" value={`${goodCount || 5} ตัวชี้วัด`} />
          <SummaryPill tone="warn" label="ควรปรับปรุง" value={`${warningCount || 1} ตัวชี้วัด`} />
          <SummaryPill tone="bad" label="ต่างกันมาก" value={`${badCount || 2} ตัวชี้วัด`} />
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

function UploadAiPanel({
  data,
  isUploading,
  status,
  uploadedNames,
  onUpload,
  onActualUpload,
}: {
  data: AiData | null;
  isUploading: boolean;
  status: { tone: "success" | "error"; message: string } | null;
  uploadedNames: { ai?: string; actual?: string };
  onUpload: (file: File) => Promise<void>;
  onActualUpload: (file: File) => Promise<void>;
}) {
  const sheets = Array.from(new Set(data?.records.map((record) => record.sheet) ?? []));
  const factories = Array.from(new Set(data?.records.map((record) => record.factory) ?? []));
  const metrics = Array.from(new Set(data?.records.map((record) => record.metric) ?? []));

  return (
    <section className="grid gap-5 xl:grid-cols-[1.25fr_0.75fr]">
      <Panel title="นำเข้าไฟล์เพื่อเปรียบเทียบ">
        <div className="grid gap-5 lg:grid-cols-2">
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
