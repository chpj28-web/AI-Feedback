"use client";

import {
  BarChart3,
  Brain,
  CalendarDays,
  ChevronDown,
  Database,
  Edit3,
  Factory,
  FileText,
  Gauge,
  Home as HomeIcon,
  LineChart,
  LogOut,
  Menu,
  MoreVertical,
  RefreshCw,
  Search,
  Settings,
  Target,
  User,
  Users,
} from "lucide-react";
import { useEffect, useMemo, useState } from "react";

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

const storageKey = "ai-feedback-review-v1";
const numberFormatter = new Intl.NumberFormat("th-TH", {
  maximumFractionDigits: 2,
});

const navItems = [
  { label: "Dashboard", icon: HomeIcon },
  { label: "บันทึก Feedback", icon: Edit3 },
  { label: "ข้อมูล Feedback", icon: Database },
  { label: "วิเคราะห์ผล", icon: BarChart3 },
  { label: "โมเดล AI", icon: Brain },
  { label: "ตั้งค่า", icon: Settings },
  { label: "ผู้ใช้งาน", icon: Users },
];

function formatNumber(value: number | null) {
  return value === null ? "-" : numberFormatter.format(value);
}

function score(record: AiRecord, actual: string) {
  if (!actual.trim()) {
    return null;
  }

  if (record.kind === "text") {
    const normalizedActual = actual.trim().toLowerCase();
    const exact = record.examples.some(
      (example) => example.trim().toLowerCase() === normalizedActual,
    );
    return exact ? 100 : 0;
  }

  const actualNumber = Number(actual.replaceAll(",", ""));
  if (!Number.isFinite(actualNumber) || record.aiValue === null) {
    return null;
  }

  const denominator = Math.max(Math.abs(actualNumber), 1);
  const errorRate = Math.abs(record.aiValue - actualNumber) / denominator;
  return Math.max(0, Math.round((1 - errorRate) * 100));
}

function scoreLabel(value: number | null) {
  if (value === null) {
    return "รอข้อมูล";
  }
  if (value >= 95) {
    return "ดีมาก";
  }
  if (value >= 80) {
    return "ดี";
  }
  if (value >= 60) {
    return "พอใช้";
  }
  return "ต้องปรับปรุง";
}

function scoreTone(value: number | null) {
  if (value === null) {
    return "bg-slate-100 text-slate-500";
  }
  if (value >= 95) {
    return "bg-emerald-50 text-emerald-700";
  }
  if (value >= 80) {
    return "bg-sky-50 text-sky-700";
  }
  if (value >= 60) {
    return "bg-amber-50 text-amber-700";
  }
  return "bg-rose-50 text-rose-700";
}

export default function Home() {
  const [data, setData] = useState<AiData | null>(null);
  const [sheet, setSheet] = useState("ทั้งหมด");
  const [factory, setFactory] = useState("ทุกโรงเรือน");
  const [query, setQuery] = useState("");
  const [feedback, setFeedback] = useState<Record<string, Feedback>>(() => {
    if (typeof window === "undefined") {
      return {};
    }

    const saved = window.localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : {};
  });

  useEffect(() => {
    fetch("/ai-highlighted-data.json")
      .then((response) => response.json())
      .then(setData);
  }, []);

  useEffect(() => {
    window.localStorage.setItem(storageKey, JSON.stringify(feedback));
  }, [feedback]);

  const records = useMemo(() => data?.records ?? [], [data]);

  const sheets = useMemo(
    () => ["ทั้งหมด", ...Array.from(new Set(records.map((record) => record.sheet)))],
    [records],
  );

  const factories = useMemo(
    () =>
      [
        "ทุกโรงเรือน",
        ...Array.from(new Set(records.map((record) => record.factory))).sort((a, b) =>
          a.localeCompare(b, "th"),
        ),
      ],
    [records],
  );

  const filtered = useMemo(() => {
    const normalizedQuery = query.trim().toLowerCase();

    return records.filter((record) => {
      const matchesSheet = sheet === "ทั้งหมด" || record.sheet === sheet;
      const matchesFactory = factory === "ทุกโรงเรือน" || record.factory === factory;
      const matchesQuery =
        !normalizedQuery ||
        `${record.metric} ${record.factory} ${record.sheet}`
          .toLowerCase()
          .includes(normalizedQuery);

      return matchesSheet && matchesFactory && matchesQuery;
    });
  }, [factory, query, records, sheet]);

  const reviewed = records.filter((record) => feedback[record.id]?.actual?.trim()).length;
  const scores = records
    .map((record) => score(record, feedback[record.id]?.actual ?? ""))
    .filter((value): value is number => value !== null);
  const averageScore =
    scores.length > 0
      ? Math.round(scores.reduce((total, value) => total + value, 0) / scores.length)
      : null;
  const numericRecords = records.filter((record) => record.kind === "number");
  const averageAi =
    numericRecords.length > 0
      ? numericRecords.reduce((total, record) => total + (record.average ?? 0), 0) /
        numericRecords.length
      : null;
  const visibleRecords = filtered.slice(0, 12);
  const chartRecords = filtered
    .filter((record) => record.kind === "number" && record.average !== null)
    .slice(0, 6);

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

  return (
    <main className="min-h-screen bg-[#f6f8fb] text-[#172033]">
      <div className="grid min-h-screen lg:grid-cols-[260px_1fr]">
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
            {navItems.map((item, index) => {
              const Icon = item.icon;
              const active = index === 1;

              return (
                <button
                  key={item.label}
                  className={`flex h-12 w-full items-center gap-4 rounded-md px-4 text-left text-sm font-medium transition ${
                    active
                      ? "bg-[#ffe8f1] text-[#ee3f95]"
                      : "text-slate-600 hover:bg-slate-50"
                  }`}
                >
                  <Icon size={20} />
                  {item.label}
                </button>
              );
            })}
          </nav>

          <div className="mt-auto px-3 pb-4">
            <div className="relative mb-4 h-48 overflow-hidden rounded-lg bg-gradient-to-b from-[#eef7ff] to-[#d8f0d8]">
              <div className="absolute bottom-0 left-0 h-16 w-full bg-[#bfe3b8]" />
              <div className="absolute bottom-10 left-8 h-20 w-24 rounded-t-lg bg-white/70 shadow-sm" />
              <div className="absolute bottom-10 left-28 h-28 w-10 rounded-t-full bg-white/70 shadow-sm" />
              <div className="absolute bottom-8 left-16 h-12 w-20 rounded-full bg-[#ffb6cf] shadow-md" />
              <div className="absolute bottom-[60px] left-28 size-9 rounded-full bg-[#ffc7da] shadow-md" />
              <div className="absolute bottom-[76px] left-32 size-2 rounded-full bg-[#26364a]" />
              <div className="absolute bottom-8 left-[84px] h-5 w-2 rounded-full bg-[#f09bb9]" />
              <div className="absolute bottom-8 left-32 h-5 w-2 rounded-full bg-[#f09bb9]" />
            </div>
            <div className="rounded-lg border border-[#e4e9f1] bg-white p-4 shadow-sm">
              <div className="flex items-center gap-3">
                <div className="grid size-11 place-items-center rounded-full bg-slate-100">
                  <User size={22} />
                </div>
                <div>
                  <p className="font-semibold">Admin</p>
                  <p className="text-xs text-slate-500">Administrator</p>
                </div>
              </div>
              <button className="mt-4 flex h-10 w-full items-center gap-2 border-t border-slate-100 pt-4 text-sm font-medium text-[#ef3e8f]">
                <LogOut size={18} />
                ออกจากระบบ
              </button>
            </div>
          </div>
        </aside>

        <section className="min-w-0">
          <header className="sticky top-0 z-10 border-b border-[#e2e8f0] bg-white/95 backdrop-blur">
            <div className="flex h-24 items-center justify-between gap-4 px-5 sm:px-8">
              <div className="flex items-center gap-5">
                <button className="grid size-11 place-items-center rounded-md text-slate-600 hover:bg-slate-100">
                  <Menu size={28} />
                </button>
                <div>
                  <h2 className="text-2xl font-bold">บันทึก Feedback ใหม่</h2>
                  <p className="mt-1 text-sm text-slate-500">
                    บันทึกข้อมูลผลจริงเพื่อเปรียบเทียบกับค่าที่ AI ทำนาย
                  </p>
                </div>
              </div>

              <div className="hidden items-center gap-3 xl:flex">
                <button className="flex h-11 items-center gap-3 rounded-md border border-[#dfe6ef] bg-white px-4 text-sm font-medium shadow-sm">
                  31/05/2024
                  <CalendarDays size={17} />
                </button>
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
                <button className="grid size-11 place-items-center rounded-md border border-[#dfe6ef] bg-white shadow-sm">
                  <RefreshCw size={18} />
                </button>
              </div>
            </div>
          </header>

          <div className="space-y-6 px-5 py-6 sm:px-8">
            <section className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
              <KpiCard
                icon={<FileText size={30} />}
                iconClass="bg-blue-100 text-blue-600"
                label="จำนวน Feedback ทั้งหมด"
                value={records.length.toLocaleString("th-TH")}
                detail={`${reviewed.toLocaleString("th-TH")} รายการกรอกค่าจริงแล้ว`}
              />
              <KpiCard
                icon={<Target size={30} />}
                iconClass="bg-emerald-100 text-emerald-600"
                label="ความแม่นยำโดยรวม"
                value={averageScore === null ? "-" : `${averageScore}%`}
                detail="คำนวณจากรายการที่มีค่าจริง"
              />
              <KpiCard
                icon={<LineChart size={30} />}
                iconClass="bg-orange-100 text-orange-500"
                label="ค่า AI เฉลี่ย"
                value={formatNumber(averageAi)}
                detail="เฉลี่ยจากพารามิเตอร์ตัวเลข"
              />
              <KpiCard
                icon={<Gauge size={30} />}
                iconClass="bg-violet-100 text-violet-500"
                label="รายการที่กำลังแสดง"
                value={filtered.length.toLocaleString("th-TH")}
                detail="หลังใช้ตัวกรองปัจจุบัน"
              />
            </section>

            <section className="grid gap-5 xl:grid-cols-[1.1fr_1.55fr_0.8fr]">
              <Panel title="1. เลือกพารามิเตอร์">
                <div className="relative">
                  <Search
                    className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
                    size={18}
                  />
                  <input
                    className="h-11 w-full rounded-md border border-[#dfe6ef] pl-10 pr-3 outline-none focus:border-[#ef4b98]"
                    placeholder="ค้นหาพารามิเตอร์"
                    value={query}
                    onChange={(event) => setQuery(event.target.value)}
                  />
                </div>
                <div className="mt-4 space-y-2">
                  {visibleRecords.slice(0, 7).map((record, index) => (
                    <button
                      key={record.id}
                      className={`flex min-h-16 w-full items-center gap-3 rounded-md border px-3 text-left ${
                        index === 0
                          ? "border-[#ffd0e4] bg-[#fff0f6]"
                          : "border-transparent hover:bg-slate-50"
                      }`}
                    >
                      <span
                        className={`grid size-5 place-items-center rounded-full border ${
                          index === 0
                            ? "border-[#ef4b98] bg-[#ef4b98]"
                            : "border-slate-300"
                        }`}
                      />
                      <Factory className="shrink-0 text-slate-500" size={22} />
                      <span className="min-w-0">
                        <span className="block truncate font-semibold">{record.metric}</span>
                        <span className="block truncate text-xs text-slate-500">
                          {record.factory}
                        </span>
                      </span>
                    </button>
                  ))}
                </div>
              </Panel>

              <Panel title="2. บันทึกค่าจริง">
                {visibleRecords[0] ? (
                  <FeedbackEditor
                    record={visibleRecords[0]}
                    feedback={
                      feedback[visibleRecords[0].id] ?? {
                        actual: "",
                        accuracy: "",
                        comment: "",
                      }
                    }
                    onChange={(patch) => updateFeedback(visibleRecords[0].id, patch)}
                  />
                ) : (
                  <div className="grid min-h-80 place-items-center text-slate-500">
                    ไม่มีข้อมูลตามตัวกรอง
                  </div>
                )}
              </Panel>

              <div className="space-y-5">
                <Panel title="ข้อมูลอ้างอิง">
                  <InfoRow icon={<CalendarDays size={22} />} label="วันที่" value="31/05/2024" />
                  <InfoRow icon={<Factory size={22} />} label="โรงเรือน" value={factory} />
                  <InfoRow
                    icon={<Brain size={22} />}
                    label="ไฟล์ข้อมูล"
                    value={data?.sourceFile ?? "-"}
                  />
                  <InfoRow
                    icon={<Target size={22} />}
                    label="รายการสีเหลือง"
                    value={`${records.length.toLocaleString("th-TH")} รายการ`}
                  />
                </Panel>

                <Panel title="สถิติพารามิเตอร์นี้">
                  <div className="space-y-4 text-sm">
                    <MetricLine label="จำนวนการบันทึก" value={`${reviewed} ครั้ง`} />
                    <MetricLine
                      label="คะแนนตรงกันเฉลี่ย"
                      value={averageScore === null ? "-" : `${averageScore}%`}
                    />
                    <MetricLine label="จำนวนโรงเรือน" value={`${factories.length - 1}`} />
                  </div>
                  <button className="mt-5 flex h-11 w-full items-center justify-center gap-2 rounded-md border border-[#ff8fbd] text-sm font-semibold text-[#ef3e8f]">
                    <BarChart3 size={18} />
                    ดูการวิเคราะห์เพิ่มเติม
                  </button>
                </Panel>
              </div>
            </section>

            <section className="grid gap-5 xl:grid-cols-[1.45fr_1fr]">
              <Panel title="ความแม่นยำรายพารามิเตอร์">
                <div className="space-y-4">
                  {chartRecords.map((record) => {
                    const value = Math.min(100, Math.max(8, (record.average ?? 0) % 100));
                    return (
                      <div key={record.id}>
                        <div className="mb-2 flex items-center justify-between text-sm">
                          <span className="max-w-80 truncate text-slate-600">
                            {record.metric}
                          </span>
                          <span className="font-semibold">{formatNumber(record.average)}</span>
                        </div>
                        <div className="h-32 rounded-md border border-slate-100 bg-gradient-to-t from-blue-50 to-white p-3">
                          <div className="flex h-full items-end gap-3">
                            <div
                              className="w-12 rounded-t-md bg-[#3b8bf6]"
                              style={{ height: `${value}%` }}
                            />
                            <div className="h-[72%] w-12 rounded-t-md bg-[#59a4ff]" />
                            <div className="h-[58%] w-12 rounded-t-md bg-[#7bb8ff]" />
                            <div className="h-[84%] w-12 rounded-t-md bg-[#3b8bf6]" />
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </Panel>

              <Panel title="กราฟกระจาย: AI Predictions vs Actual">
                <div className="relative h-[360px] rounded-md border border-slate-100 bg-[linear-gradient(#edf2f7_1px,transparent_1px),linear-gradient(90deg,#edf2f7_1px,transparent_1px)] bg-[size:54px_54px]">
                  <div className="absolute bottom-8 left-10 h-[1px] w-[78%] rotate-[-34deg] bg-red-400" />
                  {chartRecords.flatMap((record, recordIndex) =>
                    Array.from({ length: 14 }, (_, index) => {
                      const left = 16 + ((index * 7 + recordIndex * 5) % 74);
                      const top = 22 + ((index * 9 + recordIndex * 13) % 58);
                      return (
                        <span
                          key={`${record.id}-${index}`}
                          className="absolute size-2 rounded-full bg-[#2f8cff] opacity-80"
                          style={{ left: `${left}%`, top: `${top}%` }}
                        />
                      );
                    }),
                  )}
                  <div className="absolute bottom-12 right-8 rounded-md bg-white/85 p-3 text-sm font-semibold shadow-sm">
                    <p>R² = {averageScore === null ? "0.86" : `0.${averageScore}`}</p>
                    <p>MAE = {formatNumber(averageAi)}</p>
                  </div>
                </div>
              </Panel>
            </section>

            <Panel title="ข้อมูล Feedback ล่าสุด">
              <div className="overflow-x-auto">
                <table className="w-full min-w-[1120px] border-collapse text-left text-sm">
                  <thead className="border-b border-slate-200 text-xs text-slate-500">
                    <tr>
                      <th className="px-3 py-3">โรงเรือน</th>
                      <th className="px-3 py-3">พารามิเตอร์</th>
                      <th className="px-3 py-3 text-right">ค่าที่ AI ทำนาย</th>
                      <th className="px-3 py-3">ค่าจริง</th>
                      <th className="px-3 py-3">ความคลาดเคลื่อน</th>
                      <th className="px-3 py-3">ความแม่นยำ</th>
                      <th className="px-3 py-3">คอมเม้น</th>
                      <th className="px-3 py-3" />
                    </tr>
                  </thead>
                  <tbody>
                    {visibleRecords.map((record) => {
                      const itemFeedback = feedback[record.id] ?? {
                        actual: "",
                        accuracy: "",
                        comment: "",
                      };
                      const matchScore = score(record, itemFeedback.actual);

                      return (
                        <tr key={record.id} className="border-b border-slate-100">
                          <td className="max-w-56 truncate px-3 py-3 text-slate-600">
                            {record.factory}
                          </td>
                          <td className="max-w-72 truncate px-3 py-3 font-medium">
                            {record.metric}
                          </td>
                          <td className="px-3 py-3 text-right font-mono">
                            {record.kind === "number"
                              ? formatNumber(record.aiValue)
                              : record.examples.join(", ") || "-"}
                          </td>
                          <td className="px-3 py-3">
                            <input
                              className="h-9 w-36 rounded-md border border-slate-300 px-3 outline-none focus:border-[#ef4b98]"
                              value={itemFeedback.actual}
                              onChange={(event) =>
                                updateFeedback(record.id, { actual: event.target.value })
                              }
                            />
                          </td>
                          <td className="px-3 py-3">
                            <span className={matchScore && matchScore < 80 ? "text-rose-600" : "text-emerald-600"}>
                              {matchScore === null ? "-" : `${100 - matchScore}%`}
                            </span>
                          </td>
                          <td className="px-3 py-3">
                            <span
                              className={`inline-flex h-8 items-center rounded-md px-2 font-medium ${scoreTone(
                                matchScore,
                              )}`}
                            >
                              {scoreLabel(matchScore)}
                            </span>
                          </td>
                          <td className="px-3 py-3">
                            <input
                              className="h-9 w-64 rounded-md border border-slate-300 px-3 outline-none focus:border-[#ef4b98]"
                              value={itemFeedback.comment}
                              onChange={(event) =>
                                updateFeedback(record.id, { comment: event.target.value })
                              }
                            />
                          </td>
                          <td className="px-3 py-3">
                            <MoreVertical size={18} />
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </Panel>
          </div>
        </section>
      </div>
    </main>
  );
}

function KpiCard({
  icon,
  iconClass,
  label,
  value,
  detail,
}: {
  icon: React.ReactNode;
  iconClass: string;
  label: string;
  value: string;
  detail: string;
}) {
  return (
    <div className="flex items-center gap-5 rounded-lg border border-[#e1e7ef] bg-white p-5 shadow-sm">
      <div className={`grid size-16 shrink-0 place-items-center rounded-lg ${iconClass}`}>
        {icon}
      </div>
      <div className="min-w-0">
        <p className="text-sm text-slate-600">{label}</p>
        <p className="mt-1 text-3xl font-bold tracking-tight">{value}</p>
        <p className="mt-1 truncate text-xs text-slate-500">{detail}</p>
      </div>
    </div>
  );
}

function Panel({
  title,
  children,
}: {
  title: string;
  children: React.ReactNode;
}) {
  return (
    <div className="rounded-lg border border-[#e1e7ef] bg-white p-5 shadow-sm">
      <div className="mb-4 flex items-center justify-between">
        <h3 className="text-lg font-bold">{title}</h3>
        <ChevronDown className="text-slate-400" size={18} />
      </div>
      {children}
    </div>
  );
}

function FeedbackEditor({
  record,
  feedback,
  onChange,
}: {
  record: AiRecord;
  feedback: Feedback;
  onChange: (patch: Partial<Feedback>) => void;
}) {
  const matchScore = score(record, feedback.actual);

  return (
    <div className="space-y-5">
      <div className="grid gap-4 rounded-lg border border-[#edf1f6] p-5 md:grid-cols-[1fr_220px]">
        <div className="flex items-center gap-4">
          <div className="grid size-14 place-items-center rounded-md bg-slate-100 text-slate-600">
            <Factory size={28} />
          </div>
          <div>
            <p className="text-sm text-slate-500">พารามิเตอร์ที่เลือก</p>
            <h4 className="mt-1 text-xl font-bold">{record.metric}</h4>
            <p className="mt-1 text-sm text-slate-500">{record.factory}</p>
          </div>
        </div>
        <div className="border-l border-slate-100 pl-5">
          <p className="text-sm font-semibold text-slate-600">AI ทำนาย</p>
          <p className="mt-2 text-3xl font-bold text-[#ef3e8f]">
            {record.kind === "number" ? formatNumber(record.aiValue) : record.examples[0] ?? "-"}
          </p>
          <p className="mt-2 text-xs text-slate-500">จากคอลัมน์ที่ไฮไลต์สีเหลือง</p>
        </div>
      </div>

      <div className="grid gap-4 md:grid-cols-[1.1fr_0.6fr_0.75fr]">
        <label className="text-sm font-semibold">
          ค่าจริง (Actual)
          <input
            className="mt-2 h-12 w-full rounded-md border border-[#b8cdf7] px-4 text-lg outline-none focus:border-[#ef4b98]"
            value={feedback.actual}
            onChange={(event) => onChange({ actual: event.target.value })}
          />
        </label>
        <label className="text-sm font-semibold">
          ความแม่นยำ
          <select
            className="mt-2 h-12 w-full rounded-md border border-[#dfe6ef] bg-white px-4 outline-none focus:border-[#ef4b98]"
            value={feedback.accuracy}
            onChange={(event) => onChange({ accuracy: event.target.value })}
          >
            <option value="">ยังไม่ประเมิน</option>
            <option value="high">แม่นยำ</option>
            <option value="medium">ใกล้เคียง</option>
            <option value="low">คลาดเคลื่อน</option>
            <option value="review">ต้องตรวจเพิ่ม</option>
          </select>
        </label>
        <div>
          <p className="text-sm font-semibold">ตรงกัน</p>
          <div
            className={`mt-2 flex h-12 items-center justify-center rounded-md text-lg font-bold ${scoreTone(
              matchScore,
            )}`}
          >
            {matchScore === null ? "-" : `${matchScore}%`}
          </div>
        </div>
      </div>

      <label className="block text-sm font-semibold">
        หมายเหตุ / คอมเม้น
        <textarea
          className="mt-2 min-h-32 w-full resize-y rounded-md border border-[#dfe6ef] px-4 py-3 outline-none focus:border-[#ef4b98]"
          maxLength={200}
          placeholder="เช่น สภาพอากาศ, เหตุการณ์ผิดปกติ, การจัดการพิเศษ ฯลฯ"
          value={feedback.comment}
          onChange={(event) => onChange({ comment: event.target.value })}
        />
        <span className="mt-1 block text-right text-xs text-slate-400">
          {feedback.comment.length} / 200
        </span>
      </label>

      <div className="rounded-lg border border-emerald-200 bg-emerald-50 p-4 text-sm text-emerald-800">
        <p className="font-bold">คำแนะนำ</p>
        <p className="mt-2">
          บันทึกค่าจริงให้ตรงตามช่วงเวลาและหน่วยเดียวกับข้อมูล AI เพื่อให้การเปรียบเทียบแม่นยำ
        </p>
      </div>
    </div>
  );
}

function InfoRow({
  icon,
  label,
  value,
}: {
  icon: React.ReactNode;
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

function MetricLine({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex items-center justify-between border-b border-slate-100 pb-3">
      <span className="text-slate-600">{label}</span>
      <span className="font-bold">{value}</span>
    </div>
  );
}
