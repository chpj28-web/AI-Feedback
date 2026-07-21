"use client";

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

function scoreTone(value: number | null) {
  if (value === null) {
    return "bg-slate-100 text-slate-500";
  }
  if (value >= 95) {
    return "bg-emerald-100 text-emerald-800";
  }
  if (value >= 80) {
    return "bg-cyan-100 text-cyan-800";
  }
  if (value >= 60) {
    return "bg-amber-100 text-amber-800";
  }
  return "bg-rose-100 text-rose-800";
}

export default function Home() {
  const [data, setData] = useState<AiData | null>(null);
  const [sheet, setSheet] = useState("ทั้งหมด");
  const [factory, setFactory] = useState("ทั้งหมด");
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
        "ทั้งหมด",
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
      const matchesFactory = factory === "ทั้งหมด" || record.factory === factory;
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
    <main className="min-h-screen bg-slate-50 text-slate-950">
      <section className="border-b border-slate-200 bg-white">
        <div className="mx-auto flex w-full max-w-7xl flex-col gap-6 px-5 py-6 sm:px-8">
          <div className="flex flex-col justify-between gap-4 lg:flex-row lg:items-end">
            <div>
              <p className="text-sm font-semibold uppercase tracking-[0.18em] text-teal-700">
                AI Feedback
              </p>
              <h1 className="mt-2 text-3xl font-semibold sm:text-4xl">
                ตรวจเทียบผล AI จากคอลัมน์สีเหลือง
              </h1>
              <p className="mt-3 max-w-3xl text-base leading-7 text-slate-600">
                ข้อมูล AI มาจากไฟล์ {data?.sourceFile ?? "Excel"} และสรุปแยกตามโรงงาน
                เพื่อกรอกค่าจริง ประเมินความตรงกัน และเก็บ comment สำหรับปรับโมเดลรอบถัดไป
              </p>
            </div>
            <div className="grid grid-cols-3 gap-3 text-sm">
              <div className="rounded-lg border border-slate-200 bg-slate-50 p-3">
                <p className="text-slate-500">รายการ AI</p>
                <p className="mt-1 text-2xl font-semibold">{records.length}</p>
              </div>
              <div className="rounded-lg border border-slate-200 bg-slate-50 p-3">
                <p className="text-slate-500">กรอกแล้ว</p>
                <p className="mt-1 text-2xl font-semibold">{reviewed}</p>
              </div>
              <div className="rounded-lg border border-slate-200 bg-slate-50 p-3">
                <p className="text-slate-500">คะแนนเฉลี่ย</p>
                <p className="mt-1 text-2xl font-semibold">
                  {averageScore === null ? "-" : `${averageScore}%`}
                </p>
              </div>
            </div>
          </div>

          <div className="grid gap-3 rounded-lg border border-slate-200 bg-slate-50 p-4 lg:grid-cols-[1fr_1.4fr_1.2fr]">
            <label className="text-sm font-medium text-slate-700">
              ชีต
              <select
                className="mt-2 h-11 w-full rounded-md border border-slate-300 bg-white px-3 text-slate-950 outline-none focus:border-teal-600"
                value={sheet}
                onChange={(event) => setSheet(event.target.value)}
              >
                {sheets.map((option) => (
                  <option key={option}>{option}</option>
                ))}
              </select>
            </label>
            <label className="text-sm font-medium text-slate-700">
              โรงงาน
              <select
                className="mt-2 h-11 w-full rounded-md border border-slate-300 bg-white px-3 text-slate-950 outline-none focus:border-teal-600"
                value={factory}
                onChange={(event) => setFactory(event.target.value)}
              >
                {factories.map((option) => (
                  <option key={option}>{option}</option>
                ))}
              </select>
            </label>
            <label className="text-sm font-medium text-slate-700">
              ค้นหา metric
              <input
                className="mt-2 h-11 w-full rounded-md border border-slate-300 bg-white px-3 text-slate-950 outline-none focus:border-teal-600"
                placeholder="เช่น Production, Quota, รถ"
                value={query}
                onChange={(event) => setQuery(event.target.value)}
              />
            </label>
          </div>
        </div>
      </section>

      <section className="mx-auto w-full max-w-7xl px-5 py-6 sm:px-8">
        <div className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
          <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3">
            <h2 className="text-lg font-semibold">ตารางประเมินผล AI</h2>
            <span className="rounded-md bg-indigo-50 px-3 py-1 text-sm font-medium text-indigo-700">
              แสดง {filtered.length} รายการ
            </span>
          </div>

          <div className="overflow-x-auto">
            <table className="min-w-[1180px] w-full border-collapse text-left text-sm">
              <thead className="bg-slate-100 text-xs uppercase text-slate-600">
                <tr>
                  <th className="w-[260px] px-4 py-3">รายการ AI</th>
                  <th className="w-[220px] px-4 py-3">โรงงาน</th>
                  <th className="w-[150px] px-4 py-3 text-right">ค่า AI</th>
                  <th className="w-[150px] px-4 py-3">ค่าจริง</th>
                  <th className="w-[120px] px-4 py-3">ตรงกัน</th>
                  <th className="w-[150px] px-4 py-3">ความแม่นยำ AI</th>
                  <th className="w-[260px] px-4 py-3">คอมเม้น</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((record) => {
                  const itemFeedback = feedback[record.id] ?? {
                    actual: "",
                    accuracy: "",
                    comment: "",
                  };
                  const matchScore = score(record, itemFeedback.actual);

                  return (
                    <tr key={record.id} className="border-t border-slate-100 align-top">
                      <td className="px-4 py-4">
                        <p className="font-semibold text-slate-950">{record.metric}</p>
                        <p className="mt-1 text-xs text-slate-500">{record.sheet}</p>
                        <p className="mt-2 text-xs text-slate-500">
                          {record.rows} rows · avg {formatNumber(record.average)} · min{" "}
                          {formatNumber(record.min)} · max {formatNumber(record.max)}
                        </p>
                      </td>
                      <td className="px-4 py-4 text-slate-700">{record.factory}</td>
                      <td className="px-4 py-4 text-right font-mono text-slate-950">
                        {record.kind === "number"
                          ? formatNumber(record.aiValue)
                          : record.examples.join(", ") || "-"}
                      </td>
                      <td className="px-4 py-4">
                        <input
                          className="h-10 w-full rounded-md border border-slate-300 px-3 outline-none focus:border-teal-600"
                          inputMode={record.kind === "number" ? "decimal" : "text"}
                          value={itemFeedback.actual}
                          onChange={(event) =>
                            updateFeedback(record.id, { actual: event.target.value })
                          }
                        />
                      </td>
                      <td className="px-4 py-4">
                        <span
                          className={`inline-flex h-9 min-w-16 items-center justify-center rounded-md px-3 font-semibold ${scoreTone(
                            matchScore,
                          )}`}
                        >
                          {matchScore === null ? "-" : `${matchScore}%`}
                        </span>
                      </td>
                      <td className="px-4 py-4">
                        <select
                          className="h-10 w-full rounded-md border border-slate-300 bg-white px-3 outline-none focus:border-teal-600"
                          value={itemFeedback.accuracy}
                          onChange={(event) =>
                            updateFeedback(record.id, { accuracy: event.target.value })
                          }
                        >
                          <option value="">ยังไม่ประเมิน</option>
                          <option value="high">แม่นยำ</option>
                          <option value="medium">ใกล้เคียง</option>
                          <option value="low">คลาดเคลื่อน</option>
                          <option value="review">ต้องตรวจเพิ่ม</option>
                        </select>
                      </td>
                      <td className="px-4 py-4">
                        <textarea
                          className="min-h-20 w-full resize-y rounded-md border border-slate-300 px-3 py-2 outline-none focus:border-teal-600"
                          value={itemFeedback.comment}
                          onChange={(event) =>
                            updateFeedback(record.id, { comment: event.target.value })
                          }
                        />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </section>
    </main>
  );
}
