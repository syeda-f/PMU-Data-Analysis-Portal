import { Fragment, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { type Grade, type GradeRecord } from "./gradeData";
import { sectionGradeData, type SectionGradeRecord } from "./sectionGradeData";
import { collegeProfessors, findProfessorEmail } from "./collegeProfessors";
import { canonicalCourseTitle } from "./scheduleImport";
import { openOutlookDraft } from "@/lib/outlook-draft";

type DisplayMode = "Student Count" | "Percentage";
type ReportMode = "course" | "section";
type GradeRangeOption = "A+" | "A" | "B+" | "B" | "C+" | "C" | "D+" | "D" | "F" | "W";
type ChartSeries = Pick<GradeRecord, "department" | "course" | "title" | "campus" | "totalStudents" | "grades"> & { label: string; section?: string; instructor?: string; colorIndex?: number };
type SemesterComparisonRecord = {
  sections: number;
  instructors: string[];
  semester: string;
  label: string;
  color: string;
  totalStudents: number;
  grades: Record<Grade, number>;
};

const grades: Grade[] = ["A+", "A", "B+", "B", "C+", "C", "D+", "D&Below", "W"];
const allGrades: Grade[] = [...grades, "P"];
const gradeColors = ["#258C70", "#34805B", "#3E78CE", "#315FC5", "#D17829", "#B86428", "#D75A2B", "#C73C3C", "#7E6C91"];
const sectionColors = ["#2E78A4", "#D85E9C"];
// Bright, clearly different colours for the line chart (plain solid lines with round dots).
const lineGradeColors: Record<Grade, string> = {
  "A+": "#0057FF",
  A: "#00B84A",
  "B+": "#FFB800",
  B: "#FF6A00",
  "C+": "#E6007E",
  C: "#00C2D1",
  "D+": "#8A2BE2",
  "D&Below": "#E60000",
  W: "#6B7280",
  P: "#111111",
};
const instructorColors = ["#0072B2", "#E69F00", "#009E73", "#D55E00", "#CC79A7", "#56B4E9", "#000000", "#A50026", "#6A3D9A", "#7A7A00"];
const semesterColors = ["#4055D6", "#E67E3F", "#258C70", "#A04B8B", "#C44545", "#2E78A4", "#8A6A2F", "#5C6BC0"];
const CCES_DEPARTMENTS = new Set(["ITAP", "COEN", "COSC", "SOEN", "AINT", "CSEC", "GEIT"]);
const semesterLabel = (code: string) => {
  // Preserve the exact six-digit semester code used in imported filenames.
  const match = /^(\d{4})(10|20|30)$/.exec(code.trim());
  if (!match) return code;
  const season = match[2] === "10" ? "Fall" : match[2] === "20" ? "Spring" : "Summer";
  return `${code} (${season})`;
};
const gradeLabels: Record<Exclude<Grade, "P">, string> = {
  "A+": "A+",
  A: "A",
  "B+": "B+",
  B: "B",
  "C+": "C+",
  C: "C",
  "D+": "D+",
  "D&Below": "D & F",
  W: "W",
};
const comparisonGradeLabel = (grade: Grade) => grade === "D&Below" ? "D & Below" : gradeLabels[grade as Exclude<Grade, "P">];

const total = (record: GradeRecord) => record.totalStudents;
const percent = (count: number, record: GradeRecord) => record.totalStudents ? (count / record.totalStudents) * 100 : 0;
const highGrades = (record: GradeRecord) => record.grades["A+"] + record.grades.A;
const passingGrades = (record: GradeRecord) => record.grades["A+"] + record.grades.A + record.grades["B+"] + record.grades.B + record.grades["C+"] + record.grades.C + record.grades.P + record.grades["D+"];
const passingRate = (record: GradeRecord) => percent(passingGrades(record), record);
// Outcome split: passing (A+ to D+, P) + failing (D & F) + withdrawn (W) covers every student, so the three rates sum to 100%.
const failingGrades = (record: GradeRecord) => record.grades["D&Below"];
const failingRate = (record: GradeRecord) => percent(failingGrades(record), record);
const withdrawalRateOf = (record: GradeRecord) => percent(record.grades.W, record);
const outcomeSplit = (record: GradeRecord) => `Pass ${passingRate(record).toFixed(1)}% · Fail ${failingRate(record).toFixed(1)}% · Withdraw ${withdrawalRateOf(record).toFixed(1)}%`;
// A single course is selected by default (CS1-style intro course): GEIT 1411 when present, otherwise the first course alphabetically.
const pickDefaultCourse = (records: SectionGradeRecord[], semesters: string[], department = "All") => {
  const courses = [...new Set(records
    .filter((record) => semesters.includes(record.semester) && CCES_DEPARTMENTS.has(record.department.toUpperCase()) && (department === "All" || record.department.toUpperCase() === department))
    .map((record) => record.course))].sort();
  return courses.includes("GEIT 1411") ? "GEIT 1411" : courses[0] ?? "All";
};
const gradeRangeCount = (record: GradeRecord, range: GradeRangeOption[]) => {
  let count = range.reduce((sum, grade) => {
    if (grade === "D" || grade === "F") return sum;
    return sum + record.grades[grade as Grade];
  }, 0);
  if (range.includes("D") || range.includes("F")) count += record.grades["D&Below"];
  return count;
};
const campusName = (record: GradeRecord) => record.campus.replace(" Campus", "");
const cohortKey = (record: GradeRecord | SectionGradeRecord) => `${"semester" in record ? record.semester : ""}|${record.course}|${record.title}|${record.campus}|${"section" in record ? record.section : ""}`;
const campusColor = (record: GradeRecord) => record.campus === "Male Campus" ? sectionColors[0] : sectionColors[1];
const chartSeriesColor = (record: ChartSeries) => {
  if (!record.section) return campusColor(record);
  const maleColors = ["#123B73", "#B8DAFF", "#2C6FB2", "#DCEEFF", "#184F8C", "#8FC5F7"];
  const femaleColors = ["#7A174F", "#F7B5D8", "#A82E71", "#FFD9EC", "#C34C8D", "#E88BBB"];
  const palette = record.campus === "Male Campus" ? maleColors : femaleColors;
  return palette[(record.colorIndex ?? 0) % palette.length];
};
/** One chart/report section per semester, course, campus and section number.
 * Source rows are additive grade groups, not duplicate copies.
 * Do not combine across semesters, courses or campuses.
 */
// [MERGE] Combines duplicate rows of the SAME section (not the cross-list merge itself; that is done by the backend)
function combineSectionGradeRows(records: SectionGradeRecord[]): SectionGradeRecord[] {
  const grouped = new Map<string, SectionGradeRecord>();
  const instructors = new Map<string, Set<string>>();
  for (const record of records) {
    const key = JSON.stringify([record.semester, record.course.trim().toUpperCase(), record.campus, record.section.trim()]);
    const current = grouped.get(key);
    if (!current) {
      grouped.set(key, { ...record, grades: { ...record.grades } });
      instructors.set(key, new Set(record.instructor && record.instructor !== "Unassigned" ? [record.instructor] : []));
      continue;
    }
    // Different grade rows for the same section represent different students.
    current.totalStudents += record.totalStudents;
    allGrades.forEach((grade) => { current.grades[grade] += record.grades[grade]; });
    if (record.instructor && record.instructor !== "Unassigned") instructors.get(key)!.add(record.instructor);
  }
  grouped.forEach((record, key) => {
    const names = [...(instructors.get(key) ?? [])];
    record.instructor = names.length ? names.join(" / ") : "Unassigned";
  });
  return [...grouped.values()];
}

const sectionKey = (record: SectionGradeRecord) => `${record.semester}|${record.campus}|${record.course}|${record.section}`;
const sectionLabel = (record: SectionGradeRecord) => `Section ${record.section} · ${record.campus.replace(" Campus", "")}`;
const chartSeriesKey = (record: ChartSeries) => `${"semester" in record ? record.semester : ""}|${record.course}|${record.campus}|${record.section ?? ""}|${record.instructor ?? ""}`;
const normalizeHeading = (value: unknown) => String(value ?? "").toLowerCase().replace(/[^a-z0-9+]/g, "");
const normalizeTitle = (value: string) => value.toLowerCase().replace(/[^a-z0-9]/g, "");
const canonicalCrossListedCodes: Record<string, string> = {
  operatingsystems: "COSC 4361",
  systemsprogramming: "ITAP 3411",
  cloudcomputing: "ITAP 4315",
  ecommerce: "SOEN 4371",
  algorithms: "COSC 3351",
  computernetworks: "COEN 3361",
  webprogramming: "SOEN 2312",
  programminglanguages: "COSC 4461",
  networksecurity: "CSEC 3431",
  networkmanagement: "ITAP 2431",
};
const officialProfessorNames = Object.values(collegeProfessors).flat();
const professorNameMap = new Map(officialProfessorNames.map((name) => [name.trim().toLowerCase(), name]));
const uniformProfessorName = (name: string) => professorNameMap.get(name.trim().toLowerCase()) ?? name.trim();
const instructorMatches = (recordInstructor: string, selectedInstructor: string) =>
  selectedInstructor === "All" || recordInstructor.split(/\s+\/\s+/).some((name) => name.trim() === selectedInstructor);
const uniformGradeRecord = (record: SectionGradeRecord): SectionGradeRecord => ({
  ...record,
  title: canonicalCourseTitle(record.course, record.title),
  instructor: uniformProfessorName(record.instructor),
});

function parseGradeDataset(file: File, data: ArrayBuffer): SectionGradeRecord[] {
  const workbook = XLSX.read(data, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
  const headerIndex = rows.findIndex((row) => row.some((cell) => {
    const heading = normalizeHeading(cell);
    return heading === "coursecode" || heading === "course";
  }));
  if (headerIndex < 0) throw new Error("The file does not contain a recognizable course header.");
  const headers = rows[headerIndex].map(normalizeHeading);
  const find = (...aliases: string[]) => headers.findIndex((header) => aliases.includes(header));
  const prefixIndex = find("coursecode", "department", "prefix");
  const courseIndex = find("course", "coursenumber", "courseno");
  const sectionIndex = find("section", "sectionno", "sectionnumber");
  const titleIndex = find("coursetitle", "title");
  const instructorIndex = find("instructorname", "instructor");
  const totalIndex = find("totalnoofstudents", "totalstudents", "students", "total");
  const campusIndex = find("campus");
  const semesterIndex = find("semester", "term", "termcode");
  const gradeIndexes: Record<Grade, number> = {
    "A+": find("a+"),
    A: find("a"),
    "B+": find("b+"),
    B: find("b"),
    "C+": find("c+"),
    C: find("c"),
    P: find("p"),
    "D+": find("d+"),
    "D&Below": find("d&below", "dbelow", "dandbelow", "df"),
    W: find("w", "withdrawra", "withdrawal", "withdrawals"),
  };
  if (prefixIndex < 0 || sectionIndex < 0 || titleIndex < 0) {
    throw new Error("The file is missing required course, section, or title columns.");
  }
  const numberValue = (value: unknown) => {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
  };
  const filenameCampus = /female/i.test(file.name) ? "Female Campus" : /male/i.test(file.name) ? "Male Campus" : null;
  const records = rows.slice(headerIndex + 1).flatMap((row) => {
    const prefixValue = String(row[prefixIndex] ?? "").trim();
    const courseValue = courseIndex >= 0 ? String(row[courseIndex] ?? "").trim() : "";
    const course = /\d/.test(prefixValue) ? prefixValue : `${prefixValue} ${courseValue}`.trim();
    const section = String(row[sectionIndex] ?? "").trim();
    const title = String(row[titleIndex] ?? "").trim();
    if (!course || !section || !title || !/\d/.test(course)) return [];
    const grades = Object.fromEntries(allGrades.map((grade) => [
      grade,
      gradeIndexes[grade] >= 0 ? numberValue(row[gradeIndexes[grade]]) : 0,
    ])) as Record<Grade, number>;
    const campusValue = campusIndex >= 0 ? String(row[campusIndex] ?? "") : "";
    const campus = /female|^f$/i.test(campusValue)
      ? "Female Campus"
      : /male|^m$/i.test(campusValue)
        ? "Male Campus"
        : filenameCampus ?? (section.startsWith("2") ? "Female Campus" : "Male Campus");
    const department = course.split(/\s+/)[0];
    const gradeTotal = allGrades.reduce((sum, grade) => sum + grades[grade], 0);
    return [{
      semester: semesterIndex >= 0 ? String(row[semesterIndex] || "202510") : "202510",
      department,
      course,
      section,
      title: canonicalCourseTitle(course, title),
      instructor: uniformProfessorName(instructorIndex >= 0 ? String(row[instructorIndex] || "Unassigned") : "Unassigned"),
      campus,
      totalStudents: totalIndex >= 0 ? numberValue(row[totalIndex]) : gradeTotal,
      grades,
    } satisfies SectionGradeRecord];
  });
  return records.map(uniformGradeRecord);
}
const API_BASE = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8000";

type BackendGradeRow = Record<string, unknown>;

function backendToSection(record: BackendGradeRow): SectionGradeRecord {
  const string = (key: string) => String(record[key] ?? "").trim();
  const number = (key: string) => Number(record[key] ?? 0) || 0;
  const course = string("course_code");
  const campusValue = string("campus").toLowerCase();
  const campus = campusValue.startsWith("f") ? "Female Campus" : "Male Campus";
  return {
    semester: string("term_code") || "202510",
    department: course.split(/\s+/)[0] || string("department"),
    course,
    section: string("section_number"),
    title: string("course_title"),
    instructor: string("instructor") || "Unassigned",
    campus,
    totalStudents: number("total_students"),
    grades: {
      "A+": number("A+"), A: number("A"), "B+": number("B+"),
      B: number("B"), "C+": number("C+"), C: number("C"),
      P: number("P"), "D+": number("D+"),
      "D&Below": number("D&Below"), W: number("withdraw_count"),
    },
  };
}

// [MERGE] API call for section-level rows (?merge=true|false)
async function requestGradeSections(merged: boolean): Promise<SectionGradeRecord[]> {
  const response = await fetch(`${API_BASE}/api/grades?merge=${merged}&record_type=section`);
  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.detail ?? `Unable to load grades (${response.status}).`);
  }
  const rows: BackendGradeRow[] = await response.json();
  return rows.map(backendToSection);
}

// [MERGE] API call for course-level rows (?merge=true|false)
async function requestGradeCourses(merged: boolean): Promise<SectionGradeRecord[]> {
  const response = await fetch(`${API_BASE}/api/grades?merge=${merged}&record_type=course`);
  if (!response.ok) throw new Error(`Unable to load course grades (${response.status}).`);
  const rows: BackendGradeRow[] = await response.json();
  return rows.map(backendToSection);
}

const downloadBlob = (blob: Blob, filename: string) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
};

const xmlEscape = (value: string) => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const crcTable = Array.from({ length: 256 }, (_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit += 1) value = (value & 1) ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});

const crc32 = (data: Uint8Array) => {
  let crc = 0xffffffff;
  data.forEach((byte) => { crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8); });
  return (crc ^ 0xffffffff) >>> 0;
};

export const createDocxBlob = (lines: string[]) => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${lines.map((line) => `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(line)}</w:t></w:r></w:p>`).join("")}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:body></w:document>`;
  const entries = [
    { name: "[Content_Types].xml", content: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
    { name: "_rels/.rels", content: `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: "word/document.xml", content: documentXml },
  ];
  const encoder = new TextEncoder();
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;
  entries.forEach((entry) => {
    const name = encoder.encode(entry.name);
    const data = encoder.encode(entry.content);
    const checksum = crc32(data);
    const local = new Uint8Array(30 + name.length + data.length);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint32(14, checksum, true);
    localView.setUint32(18, data.length, true);
    localView.setUint32(22, data.length, true);
    localView.setUint16(26, name.length, true);
    local.set(name, 30);
    local.set(data, 30 + name.length);
    localParts.push(local);

    const central = new Uint8Array(46 + name.length);
    const centralView = new DataView(central.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint32(16, checksum, true);
    centralView.setUint32(20, data.length, true);
    centralView.setUint32(24, data.length, true);
    centralView.setUint16(28, name.length, true);
    centralView.setUint32(42, offset, true);
    central.set(name, 46);
    centralParts.push(central);
    offset += local.length;
  });
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, entries.length, true);
  endView.setUint16(10, entries.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, offset, true);
  const blobParts = [...localParts, ...centralParts, end].map((part) =>
    part.buffer.slice(part.byteOffset, part.byteOffset + part.byteLength) as ArrayBuffer
  );
  return new Blob(blobParts, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
};

const createPdfBlob = (sourceLines: string[]) => {
  const clean = (value: string) => value.normalize("NFKD").replace(/[^\x20-\x7e]/g, "-").replace(/([\\()])/g, "\\$1");
  const pages = Array.from({ length: Math.max(1, Math.ceil(sourceLines.length / 48)) }, (_, index) => sourceLines.slice(index * 48, (index + 1) * 48));
  const fontId = 3 + pages.length * 2;
  const objects: string[] = [
    `<< /Type /Catalog /Pages 2 0 R >>`,
    `<< /Type /Pages /Kids [${pages.map((_, index) => `${3 + index * 2} 0 R`).join(" ")}] /Count ${pages.length} >>`,
  ];
  pages.forEach((lines, index) => {
    const pageId = 3 + index * 2;
    const contentId = pageId + 1;
    const stream = `BT /F1 10 Tf 44 790 Td 14 TL ${lines.map((line) => `(${clean(line)}) Tj T*`).join(" ")} ET`;
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 842] /Resources << /Font << /F1 ${fontId} 0 R >> >> /Contents ${contentId} 0 R >>`);
    objects.push(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
  });
  objects.push(`<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>`);
  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((object, index) => {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.slice(1).map((offset) => `${String(offset).padStart(10, "0")} 00000 n `).join("\n")}\ntrailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
  return new Blob([pdf], { type: "application/pdf" });
};

function MiniIcon({ name }: { name: "upload" | "download" | "users" | "sections" | "award" | "alert" | "check" | "chart" }) {
  const paths = {
    upload: <><path d="M12 16V4m0 0L7.5 8.5M12 4l4.5 4.5" /><path d="M4 15v4h16v-4" /></>,
    download: <><path d="M12 4v12m0 0 4.5-4.5M12 16l-4.5-4.5" /><path d="M4 17v3h16v-3" /></>,
    users: <><circle cx="9" cy="8" r="3" /><path d="M3 20c0-4 2-6 6-6s6 2 6 6M16 5.5a3 3 0 0 1 0 5.5M17 14c2.7.4 4 2.4 4 5" /></>,
    sections: <><rect x="4" y="4" width="16" height="16" rx="2" /><path d="M4 10h16M10 4v16" /></>,
    award: <><circle cx="12" cy="9" r="5" /><path d="m8.5 13-1 8 4.5-2 4.5 2-1-8" /></>,
    alert: <><path d="M12 3 2.5 20h19L12 3Z" /><path d="M12 9v5m0 3h.01" /></>,
    check: <path d="m5 12 4 4L19 6" />,
    chart: <><path d="M4 20V10M10 20V4M16 20v-7M22 20V7" /></>,
  };
  return <svg viewBox="0 0 24 24" aria-hidden="true">{paths[name]}</svg>;
}

function GradeStat({ icon, label, value, note, warning = false }: { icon: "users" | "sections" | "award" | "alert"; label: string; value: string; note: string; warning?: boolean }) {
  return (
    <div className={`grade-stat ${warning ? "grade-stat--warning" : ""}`}>
      <div className="grade-stat__icon"><MiniIcon name={icon} /></div>
      <div><span>{label}</span><strong>{value}</strong><small>{note}</small></div>
    </div>
  );
}

function MetricToggle({ mode, onChange }: { mode: DisplayMode; onChange: (mode: DisplayMode) => void }) {
  return <div className="grade-chart-metric" aria-label="Chart display metric">{(["Student Count", "Percentage"] as DisplayMode[]).map((option) => <button className={mode === option ? "active" : ""} onClick={() => onChange(option)} key={option}>{option === "Student Count" ? "# Student Count" : "% Percentage"}</button>)}</div>;
}

function GradeRangeSelect({ values, options, onChange }: { values: GradeRangeOption[]; options: GradeRangeOption[]; onChange: (values: GradeRangeOption[]) => void }) {
  return <details className="grade-range-select"><summary>{values.length} grades selected</summary><div>{options.map((option) => <label key={option}><input type="checkbox" checked={values.includes(option)} onChange={() => onChange(values.includes(option) ? values.filter((value) => value !== option) : [...values, option])} /><span>{option}</span></label>)}</div></details>;
}

function GradeKey({ grade }: { grade: Grade }) {
  return <i style={{ display: "inline-block", width: 14, height: 14, borderRadius: "50%", background: lineGradeColors[grade], flex: "none" }} aria-hidden="true" />;
}

function SemesterComparisonChart({
  data,
  mode,
  selectedGrades,
  course,
}: {
  data: SemesterComparisonRecord[];
  mode: DisplayMode;
  selectedGrades: Grade[];
  course: string;
}) {
  const [hovered, setHovered] = useState<{ semester: string; grade: Grade } | null>(null);
  const width = 920;
  const height = 355;
  const padding = { top: 28, right: 30, bottom: 78, left: 66 };
  const plotWidth = width - padding.left - padding.right;
  const plotHeight = height - padding.top - padding.bottom;
  const share = (record: SemesterComparisonRecord, grade: Grade) => record.totalStudents ? record.grades[grade] / record.totalStudents * 100 : 0;
  const valueFor = (record: SemesterComparisonRecord, grade: Grade) => mode === "Percentage" ? share(record, grade) : record.grades[grade];
  const available = data.filter((record) => record.totalStudents > 0);
  const rawMaximum = Math.max(0, ...available.flatMap((record) => selectedGrades.map((grade) => valueFor(record, grade))));
  const maximum = mode === "Percentage"
    ? Math.max(10, Math.min(100, Math.ceil(rawMaximum / 10) * 10))
    : Math.max(5, Math.ceil(rawMaximum / 5) * 5);
  const xPosition = (index: number) => padding.left + (data.length === 1 ? plotWidth / 2 : index / (data.length - 1) * plotWidth);
  const yPosition = (value: number) => padding.top + plotHeight - value / maximum * plotHeight;
  const overallTotal = available.reduce((sum, record) => sum + record.totalStudents, 0);
  const overallShare = (grade: Grade) => overallTotal ? available.reduce((sum, record) => sum + record.grades[grade], 0) / overallTotal * 100 : 0;
  const hoveredIndex = hovered ? data.findIndex((record) => record.semester === hovered.semester) : -1;
  const hoveredRecord = hoveredIndex >= 0 ? data[hoveredIndex] : null;
  const hoveredGrade = hovered?.grade ?? null;
  const dotX = hoveredRecord ? xPosition(hoveredIndex) : 0;
  const dotY = hoveredRecord && hoveredGrade ? yPosition(valueFor(hoveredRecord, hoveredGrade)) : 0;

  return (
    <div className="grade-semester-chart" style={{ position: "relative" }}>
      <div style={{ position: "relative" }}>
      <svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label={`${course} grade trends by semester`} style={{ width: "100%", height: "auto", overflow: "visible" }} onMouseLeave={() => setHovered(null)}>
        <text x={-height / 2} y="17" transform="rotate(-90)" textAnchor="middle" className="grade-semester-chart__axis-label">{mode}</text>
        {[0, 1, 2, 3, 4].map((tick) => {
          const value = maximum * tick / 4;
          const y = yPosition(value);
          return <g key={tick}>
            <line x1={padding.left} x2={width - padding.right} y1={y} y2={y} className="grade-semester-chart__grid" />
            <text x={padding.left - 10} y={y + 4} textAnchor="end" className="grade-semester-chart__tick">{mode === "Percentage" ? `${value.toFixed(0)}%` : Math.round(value)}</text>
          </g>;
        })}
        {data.map((record, index) => <text key={record.semester} x={xPosition(index)} y={height - 43} textAnchor="middle" className="grade-semester-chart__semester">{record.label}</text>)}
        {hoveredIndex >= 0 && <line x1={xPosition(hoveredIndex)} x2={xPosition(hoveredIndex)} y1={padding.top} y2={padding.top + plotHeight} stroke="#64748B" strokeWidth="1.5" strokeDasharray="5 5" />}
        {selectedGrades.map((grade) => {
          const color = lineGradeColors[grade];
          // Break the line at semesters without records rather than showing a false zero.
          const segments: SemesterComparisonRecord[][] = [];
          let segment: SemesterComparisonRecord[] = [];
          data.forEach((record) => {
            if (record.totalStudents > 0) segment.push(record);
            else if (segment.length) { segments.push(segment); segment = []; }
          });
          if (segment.length) segments.push(segment);
          return <g key={grade}>
            {segments.map((part, partIndex) => <polyline key={partIndex} points={part.map((record) => `${xPosition(data.indexOf(record))},${yPosition(valueFor(record, grade))}`).join(" ")} fill="none" stroke={color} strokeWidth="3.5" strokeLinejoin="round" strokeLinecap="round" />)}
            {data.map((record, index) => record.totalStudents > 0 && <g key={record.semester} pointerEvents="none">
              <circle cx={xPosition(index)} cy={yPosition(valueFor(record, grade))} r={hovered?.semester === record.semester && hovered.grade === grade ? 8 : 5.5} fill={color} stroke="#fff" strokeWidth="1.5" />
            </g>)}
          </g>;
        })}
        {selectedGrades.flatMap((grade) => data.filter((record) => record.totalStudents > 0).map((record) => <circle
          key={`hit-${grade}-${record.semester}`}
          cx={xPosition(data.indexOf(record))} cy={yPosition(valueFor(record, grade))} r="13"
          fill="transparent" style={{ cursor: "pointer" }} tabIndex={0}
          aria-label={`${course}, ${record.label}, ${comparisonGradeLabel(grade)}: ${record.grades[grade]} students`}
          onMouseEnter={() => setHovered({ semester: record.semester, grade })}
          onFocus={() => setHovered({ semester: record.semester, grade })}
        />))}
      </svg>
      {hoveredRecord && hoveredGrade && (() => {
        const count = hoveredRecord.grades[hoveredGrade];
        const now = share(hoveredRecord, hoveredGrade);
        const previous = [...data.slice(0, hoveredIndex)].reverse().find((record) => record.totalStudents > 0);
        const overall = overallShare(hoveredGrade);
        const arrow = (value: number) => value > 0.05 ? "▲" : value < -0.05 ? "▼" : "■";
        // Colour the change by whether it is good news: more A/B/C grades is good, more D, F or withdrawals is bad.
        const higherIsGood = !["D+", "D&Below", "W"].includes(hoveredGrade);
        const tone = (value: number) => Math.abs(value) <= 0.05 ? "#64748B" : (value > 0) === higherIsGood ? "#0A8F3C" : "#D32F2F";
        const signed = (value: number) => `${value > 0.05 ? "+" : value < -0.05 ? "−" : ""}${Math.abs(value).toFixed(1)}`;
        const horizontal = dotX / width < 0.22 ? "-10%" : dotX / width > 0.78 ? "-90%" : "-50%";
        const vertical = dotY < 190 ? "16px" : "calc(-100% - 16px)";
        const row = { display: "flex", justifyContent: "space-between", gap: 16, marginTop: 5 } as const;
        const tipColor = lineGradeColors[hoveredGrade];
        return <div role="status" style={{ position: "absolute", left: `${dotX / width * 100}%`, top: `${dotY / height * 100}%`, transform: `translate(${horizontal}, ${vertical})`, zIndex: 3, background: "#fff", color: "#18243B", border: `2px solid ${tipColor}`, borderRadius: 10, padding: "10px 14px", boxShadow: "0 10px 28px #0003", width: 340, pointerEvents: "none", fontSize: 13, lineHeight: 1.35 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}><GradeKey grade={hoveredGrade} /><strong>{course} · {hoveredRecord.label}</strong></div>
          <div style={{ ...row, marginTop: 8, fontSize: 15 }}><strong>Grade {comparisonGradeLabel(hoveredGrade)}</strong><strong>{count.toLocaleString()} {count === 1 ? "student" : "students"} · {now.toFixed(1)}%</strong></div>
          <div style={{ ...row, borderTop: "1px solid #E2E8F0", paddingTop: 6 }}><span>vs {previous ? previous.label : "previous semester"}</span><strong style={previous ? { color: tone(now - share(previous, hoveredGrade)) } : undefined}>{previous ? `${arrow(now - share(previous, hoveredGrade))} ${signed(now - share(previous, hoveredGrade))} pp (${count - previous.grades[hoveredGrade] > 0 ? "+" : count - previous.grades[hoveredGrade] < 0 ? "−" : ""}${Math.abs(count - previous.grades[hoveredGrade])} students)` : "first semester shown"}</strong></div>
          {available.length >= 3 && <div style={row}><span>vs average of {available.length} semesters ({overall.toFixed(1)}%)</span><strong style={{ color: tone(now - overall) }}>{arrow(now - overall)} {signed(now - overall)} pp</strong></div>}
        </div>;
      })()}
      </div>
      <div className="grade-semester-chart__x-label">Semester</div>
      <div className="grade-semester-chart__grade-key" aria-label="Grade line styles">
        {selectedGrades.map((grade) => <span key={grade} style={{ display: "inline-flex", alignItems: "center", gap: 6 }}><GradeKey grade={grade} />{comparisonGradeLabel(grade)}</span>)}
      </div>
    </div>
  );
}

function GradeTooltip({ record, grade, seriesLabel, instructor, selectionTotal }: { record: GradeRecord; grade: Exclude<Grade, "P">; seriesLabel?: string; instructor?: string; selectionTotal?: number }) {
  const count = record.grades[grade];
  return (
    <span className="grade-tooltip">
      <span className="grade-tooltip__heading"><i style={{ background: gradeColors[grades.indexOf(grade)] }} />Grade {gradeLabels[grade]}</span>
      <span className="grade-tooltip__row"><span>Course Record:</span><b>{seriesLabel ?? record.campus}</b></span>
      {instructor && <span className="grade-tooltip__row"><span>Instructor:</span><b>{instructor}</b></span>}
      <span className="grade-tooltip__row"><span>No. of Students:</span><b className="positive">{count.toLocaleString()} students</b></span>
      <span className="grade-tooltip__row"><span>Section Grade Percentage:</span><b className="accent">{percent(count, record).toFixed(2)}%</b></span>
      {selectionTotal !== undefined && <span className="grade-tooltip__row"><span>Filtered Selection Percentage:</span><b className="accent">{(selectionTotal ? count / selectionTotal * 100 : 0).toFixed(2)}%</b></span>}
      <span className="grade-tooltip__divider" />
      <span className="grade-tooltip__row"><span>Record Total:</span><b>{record.totalStudents.toLocaleString()} enrolled</b></span>
      <span className="grade-tooltip__row"><span>Record Passing Rate:</span><b className="positive">{passingRate(record).toFixed(1)}%</b></span>
      <span className="grade-tooltip__row"><span>Record Failing Rate:</span><b className="warning">{failingRate(record).toFixed(1)}%</b></span>
      <span className="grade-tooltip__row"><span>Record Withdrawal Rate:</span><b className="warning">{percent(record.grades.W, record).toFixed(1)}%</b></span>
    </span>
  );
}

type InstructorComparisonRow = { instructor: string; sections: number; totalStudents: number; grades: Record<Grade, number> };

// Grouped bars, one bar per instructor in each grade group. Percentages are of each instructor's own students.
function InstructorBarsChart({ rows, allNames, mode, course }: { rows: InstructorComparisonRow[]; allNames: string[]; mode: DisplayMode; course: string }) {
  const pct = (count: number, totalStudents: number) => totalStudents ? count / totalStudents * 100 : 0;
  const colorFor = (name: string) => instructorColors[Math.max(0, allNames.indexOf(name)) % instructorColors.length];
  const valueFor = (row: InstructorComparisonRow, grade: Grade) => mode === "Percentage" ? pct(row.grades[grade], row.totalStudents) : row.grades[grade];
  const rawMaximum = Math.max(0, ...rows.flatMap((row) => grades.map((grade) => valueFor(row, grade))));
  const maximum = mode === "Percentage" ? Math.max(10, Math.min(100, Math.ceil(rawMaximum / 10) * 10)) : Math.max(5, Math.ceil(rawMaximum / 5) * 5);
  const width = 920, height = 372, padding = { top: 58, right: 20, bottom: 40, left: 56 };
  const plotWidth = width - padding.left - padding.right, plotHeight = height - padding.top - padding.bottom;
  const groupWidth = plotWidth / grades.length;
  const barWidth = Math.min(40, groupWidth * 0.82 / Math.max(1, rows.length));
  const y = (value: number) => padding.top + plotHeight - value / maximum * plotHeight;
  const label = (row: InstructorComparisonRow, grade: Grade) => mode === "Percentage" ? `${pct(row.grades[grade], row.totalStudents).toFixed(1)}%` : String(row.grades[grade]);
  if (rows.length === 0) return <p role="status">No records for the selected instructors.</p>;
  return <>
    <div className="grade-chart-legend grade-chart-legend--single-line">{rows.map((row, index) => <span key={row.instructor}><i style={{ background: colorFor(row.instructor) }} /><b>{index + 1}. {row.instructor} · {row.sections} {row.sections === 1 ? "section" : "sections"} · {row.totalStudents.toLocaleString()} students</b></span>)}</div>
    <svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label={`${course} grades by instructor`} style={{ width: "100%", height: "auto" }}>
      {[0, 1, 2, 3, 4].map((tick) => {
        const value = maximum * tick / 4;
        return <g key={tick}>
          <line x1={padding.left} x2={width - padding.right} y1={y(value)} y2={y(value)} stroke="#CBD5E1" strokeDasharray="3 4" />
          <text x={padding.left - 8} y={y(value) + 4} textAnchor="end" fontSize="11" fill="#64748B">{mode === "Percentage" ? `${value.toFixed(0)}%` : Math.round(value)}</text>
        </g>;
      })}
      {grades.map((grade, gradeIndex) => {
        const groupStart = padding.left + gradeIndex * groupWidth + (groupWidth - barWidth * rows.length) / 2;
        return <g key={grade}>
          {rows.map((row, rowIndex) => {
            const value = valueFor(row, grade);
            const top = y(value);
            const x = groupStart + rowIndex * barWidth;
            return <g key={row.instructor}>
              <rect x={x} y={top} width={Math.max(1, barWidth - 3)} height={padding.top + plotHeight - top} fill={colorFor(row.instructor)} stroke="#fff" strokeWidth="1">
                <title>{`${row.instructor} · Grade ${gradeLabels[grade as Exclude<Grade, "P">]}: ${row.grades[grade]} students (${pct(row.grades[grade], row.totalStudents).toFixed(1)}% of ${row.totalStudents} students) across ${row.sections} section${row.sections === 1 ? "" : "s"}`}</title>
              </rect>
              {rows.length <= 6 && <text transform={`translate(${x + (barWidth - 3) / 2 + 4},${top - 6}) rotate(-90)`} textAnchor="start" fontSize="12" fontWeight="700" fill="#18243B">{label(row, grade)}</text>}
              {barWidth >= 14 && padding.top + plotHeight - top >= 18 && <text x={x + (barWidth - 3) / 2} y={padding.top + plotHeight - 6} textAnchor="middle" fontSize="11" fontWeight="800" fill="#fff" stroke="#000" strokeWidth="3" paintOrder="stroke">{rowIndex + 1}</text>}
            </g>;
          })}
          <text x={padding.left + gradeIndex * groupWidth + groupWidth / 2} y={height - 14} textAnchor="middle" fontSize="12" fontWeight="700" fill={gradeColors[gradeIndex]}>{gradeLabels[grade as Exclude<Grade, "P">]}</text>
        </g>;
      })}
    </svg>
  </>;
}

function CohortChart({ initialRecord, initialMode }: { initialRecord: GradeRecord; initialMode: DisplayMode }) {
  const [mode, setMode] = useState<DisplayMode>(initialMode);
  const record = initialRecord;
  const maximum = Math.max(1, ...grades.map((grade) => mode === "Percentage" ? percent(record.grades[grade], record) : record.grades[grade]));
  return (
    <div className="grade-cohort-chart">
      <div className="grade-cohort-chart__heading"><div><strong>{record.course} — {record.title}</strong><span>{record.campus} · {record.totalStudents.toLocaleString()} students · {passingRate(record).toFixed(1)}% passing</span></div><MetricToggle mode={mode} onChange={setMode} /></div>
      <div className="grade-cohort-chart__plot">
        {grades.map((grade, index) => {
          const value = mode === "Percentage" ? percent(record.grades[grade], record) : record.grades[grade];
          return <div className="grade-cohort-chart__group" key={grade}><div className="grade-cohort-chart__bar-wrap"><div className="grade-cohort-chart__bar" tabIndex={0} style={{ height: `${value / maximum * 100}%`, background: campusColor(record) }}><span>{mode === "Percentage" ? `${value.toFixed(1)}%` : value}</span><GradeTooltip record={record} grade={grade as Exclude<Grade, "P">} /></div></div><strong style={{ color: gradeColors[index] }}>{gradeLabels[grade as Exclude<Grade, "P">]}</strong></div>;
        })}
      </div>
    </div>
  );
}

export default function GradeAnalysis() {
  // Cross-listed courses are merged by default; the hero button toggles the original view.
  // [MERGE] State: merged by default; isMerging = button loading state
  const [mergeEnabled, setMergeEnabled] = useState(true);
  const [isMerging, setIsMerging] = useState(false);
  // [MERGE] Initial bundled data is combined too
  const [gradeRecords, setGradeRecords] = useState<SectionGradeRecord[]>(() => combineSectionGradeRows(sectionGradeData.map(uniformGradeRecord)).filter((record) => CCES_DEPARTMENTS.has(record.department.toUpperCase())));
  const [activePanel, setActivePanel] = useState<"analysis" | "report">("analysis");
  const [courseGradeRecords, setCourseGradeRecords] = useState<SectionGradeRecord[]>([]);
  // The dashboard starts on one course; the report starts on all courses.
  const activePanelRef = useRef(activePanel);
  activePanelRef.current = activePanel;
  const panelCourseRef = useRef<{ analysis: string | null; report: string | null }>({ analysis: null, report: null });
  const defaultCourse = (records: SectionGradeRecord[], semesters: string[], department = "All") =>
    activePanelRef.current === "analysis" ? pickDefaultCourse(records, semesters, department) : "All";
  const [reportMode, setReportMode] = useState<ReportMode>("course");
  const [filters, setFilters] = useState(() => ({ semesters: ["202510"], department: "All", course: pickDefaultCourse(sectionGradeData.map(uniformGradeRecord), ["202510"]), instructor: "All", campus: "All", mode: "Student Count" as DisplayMode }));
  const [selectedSections, setSelectedSections] = useState<string[]>([]);
  const [comparisonGrades, setComparisonGrades] = useState<Grade[]>(() => [...grades]);
  const [chartInstructors, setChartInstructors] = useState<string[]>([]);
  const [multiSemesterView, setMultiSemesterView] = useState<"lines" | "bars">("lines");
  const [aType, setAType] = useState<DisplayMode>("Percentage");
  const [aThreshold, setAThreshold] = useState(20);
  const [riskType, setRiskType] = useState<DisplayMode>("Percentage");
  const [riskThreshold, setRiskThreshold] = useState(20);
  const [highGradeRange, setHighGradeRange] = useState<GradeRangeOption[]>(["A+", "A"]);
  const [lowGradeRange, setLowGradeRange] = useState<GradeRangeOption[]>(["C+", "C", "D+", "D", "F", "W"]);
  const [minimumStudents, setMinimumStudents] = useState(10);
  const [topCourseCount, setTopCourseCount] = useState(5);
  const [atRiskCourseCount, setAtRiskCourseCount] = useState(5);
  const [topSectionCount, setTopSectionCount] = useState(5);
  const [atRiskSectionCount, setAtRiskSectionCount] = useState(5);
  const [variationThreshold, setVariationThreshold] = useState(15);
  const [reportGenerated, setReportGenerated] = useState(false);
  const [expandedCohort, setExpandedCohort] = useState<string | null>(null);
  const [sortKey, setSortKey] = useState<"campus" | "total" | "department" | "course" | "title">("course");
  const [sortAsc, setSortAsc] = useState(true);
  const [isImporting, setIsImporting] = useState(false);


  const [toast, setToast] = useState<{ type: "success" | "error" | "info"; message: string } | null>(null);
  const [showEmailPanel, setShowEmailPanel] = useState(false);
  const [selectedEmailCourses, setSelectedEmailCourses] = useState<string[]>([]);
  const [reviewEmailDrafts, setReviewEmailDrafts] = useState(false);
  const [chairEmails, setChairEmails] = useState("mkhan@pmu.edu.sa, sabdulrahman@pmu.edu.sa, nmohammad@pmu.edu.sa");
  const [senderEmail, setSenderEmail] = useState("mkhan@pmu.edu.sa");
  const [sendingEmails, setSendingEmails] = useState(false);
  const [emailAttachment, setEmailAttachment] = useState<"pdf" | "csv">("pdf");
  const importInputRef = useRef<HTMLInputElement>(null);
  const toastTimerRef = useRef<number | null>(null);

  const showToast = (type: "success" | "error" | "info", message: string) => {
    setToast({ type, message });
    if (toastTimerRef.current !== null) window.clearTimeout(toastTimerRef.current);
    toastTimerRef.current = window.setTimeout(() => setToast(null), 3500);
  };

  const applyRecords = (records: SectionGradeRecord[], courseRows?: SectionGradeRecord[]) => {
    if (courseRows) setCourseGradeRecords(courseRows.filter((record) => CCES_DEPARTMENTS.has(record.department.toUpperCase())));
    // [MERGE] Combines duplicate cross-listed rows when new records are applied
    const nextRecords = combineSectionGradeRows(records.map(uniformGradeRecord)).filter((record) => CCES_DEPARTMENTS.has(record.department.toUpperCase()));
    setGradeRecords(nextRecords);
    setSelectedSections([]);
    setExpandedCohort(null);
    setReportGenerated(false);
    const availableTerms = new Set([...records, ...(courseRows ?? courseGradeRecords)].map((record) => record.semester));
    setFilters((current) => {
      const kept = current.semesters.filter((semester) => availableTerms.has(semester));
      const semesters = kept.length ? kept : [[...availableTerms].sort()[0] ?? "202510"];
      panelCourseRef.current = { analysis: null, report: null };
      return { ...current, semesters, department: "All", course: defaultCourse(nextRecords, semesters), instructor: "All", campus: "All" };
    });
  };

  useEffect(() => {
    let cancelled = false;
    // [MERGE] Initial load: fetches sections + courses with merge=mergeEnabled
    Promise.all([requestGradeSections(mergeEnabled), requestGradeCourses(mergeEnabled)])
      .then(([records, courseRows]) => {
        if (cancelled) return;
        applyRecords(records, courseRows);
      })
      .catch(() => {
        // Keep the bundled example dataset until files are imported.
      });
    return () => { cancelled = true; };
  }, [mergeEnabled]);

  // [MERGE] Toggle handler: loads merged or original data from the API, then updates mergeEnabled
  const handleMergeCrossListed = async () => {
    if (isMerging || isImporting) return;
    setIsMerging(true);
    const next = !mergeEnabled;
    try {
      const [records, courseRows] = await Promise.all([requestGradeSections(next), requestGradeCourses(next)]);
      applyRecords(records, courseRows);
      setMergeEnabled(next);
      showToast("success", next ? "Merged cross-listed view enabled." : "Original course view restored.");
    } catch (error) {
      showToast("error", error instanceof Error ? error.message : "Unable to load grade data.");
    } finally {
      setIsMerging(false);
    }
  };

  const handleGradeImport = async (files?: FileList | null) => {
    if (!files?.length || isImporting) return;
    setIsImporting(true);
    try {
      const selected = Array.from(files);
      if (selected.some((file) => !file.name.toLowerCase().endsWith(".csv"))) {
        throw new Error("Select CSV files only.");
      }
      const upload = async (replaceExisting: boolean) => {
        const form = new FormData();
        selected.forEach((file) => form.append("files", file, file.name));
        form.append("replace_existing", String(replaceExisting));
        const response = await fetch(`${API_BASE}/api/grades/import-batch`, { method: "POST", body: form });
        const result = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(typeof result.detail === "string" ? result.detail : "Grade batch import failed.");
        return result as { imported: Array<{ filename: string; semester: string; source: string }>; errors: Array<{ filename: string; error: string }> };
      };
      let result = await upload(false);
      if (result.errors.some((item) => item.error.includes("already imported"))) {
        if (window.confirm("Some datasets are already imported. Replace only those matching semester/source datasets?")) {
          result = await upload(true);
        }
      }
      if (result.imported.length) {
        // [MERGE] Re-fetch after a CSV import using the current merge state
        const [records, courseRows] = await Promise.all([requestGradeSections(mergeEnabled), requestGradeCourses(mergeEnabled)]);
        applyRecords(records, courseRows);
        const importedSemesters = [...new Set(result.imported.map((item) => item.semester))];
        setFilters((current) => ({ ...current, semesters: importedSemesters, course: defaultCourse(records.map(uniformGradeRecord), importedSemesters) }));
      }
      const errors = result.errors.map((item) => `${item.filename}: ${item.error}`);
      showToast(errors.length ? "error" : "success", `${result.imported.length} file(s) imported, ${errors.length} failed.${errors.length ? " " + errors.slice(0, 2).join(" | ") : ""}`);
    } catch (error) {
      showToast("error", error instanceof Error ? error.message : "Unable to import grade files.");
    } finally {
      setIsImporting(false);
      if (importInputRef.current) importInputRef.current.value = "";
    }
  };

  const semesterOptions = useMemo(() => [...new Set([...gradeRecords, ...courseGradeRecords].map((record) => record.semester))].sort(), [gradeRecords, courseGradeRecords]);
  // Course options depend on the selected instructor, not the other way around.
  // This uses section-level records so an instructor is matched to courses they
  // actually taught in the selected semester(s), campus and department.
  const courseOptions = useMemo(() => {
    const matching = gradeRecords.filter((record) =>
      filters.semesters.includes(record.semester) &&
      CCES_DEPARTMENTS.has(record.department.toUpperCase()) &&
      (filters.department === "All" || record.department.toUpperCase() === filters.department) &&
      (filters.campus === "All" || record.campus.startsWith(filters.campus)) &&
      instructorMatches(record.instructor, filters.instructor)
    );
    const byCourse = new Map<string, typeof matching>();
    for (const record of matching) {
      const group = byCourse.get(record.course) ?? [];
      group.push(record);
      byCourse.set(record.course, group);
    }
    return [...byCourse.entries()].map(([course, records]) => ({
      value: course,
      course,
      title: [...new Set(records.map((record) => record.title))].join(" / "),
      department: records[0].department,
    })).sort((a, b) => a.course.localeCompare(b.course));
  }, [filters.semesters, filters.department, filters.campus, filters.instructor, gradeRecords]);

  const departmentOptions = useMemo(() => [...new Set(gradeRecords.filter((record) => filters.semesters.includes(record.semester)).map((record) => record.department.toUpperCase()).filter((department) => CCES_DEPARTMENTS.has(department)))].sort(), [filters.semesters, gradeRecords]);

  // Cascading filters: semester/campus -> department -> course -> instructor.
  // When a course is selected, show only instructors who taught that course.
  // When all courses are selected, show instructors teaching any course in
  // the chosen department. The instructor can still be selected first.
  const instructorOptions = useMemo(() => [...new Set(gradeRecords.filter((item) =>
    filters.semesters.includes(item.semester) &&
    CCES_DEPARTMENTS.has(item.department.toUpperCase()) &&
    (filters.department === "All" || item.department.toUpperCase() === filters.department) &&
    (filters.course === "All" || item.course === filters.course) &&
    (filters.campus === "All" || item.campus.startsWith(filters.campus))
  ).flatMap((item) => item.instructor.split(/\s+\/\s+/)).map((name) => name.trim()).filter(Boolean))].sort(), [filters.department, filters.course, filters.campus, filters.semesters, gradeRecords]);

  const availableSections = useMemo(() => gradeRecords.filter((item) =>
    filters.semesters.includes(item.semester) &&
    (filters.department === "All" || item.department === filters.department) &&
    (filters.course === "All" || item.course === filters.course) &&
    instructorMatches(item.instructor, filters.instructor) &&
    (filters.campus === "All" || item.campus.startsWith(filters.campus))
  ), [filters.campus, filters.course, filters.department, filters.instructor, filters.semesters, gradeRecords]);

  const sections = useMemo(() => {
    const result = availableSections.filter((item) =>
      !selectedSections.length || selectedSections.includes(sectionKey(item))
    );
    return [...result].sort((a, b) => {
      const first = sortKey === "total" ? total(a) : a[sortKey];
      const second = sortKey === "total" ? total(b) : b[sortKey];
      return (typeof first === "number" ? first - (second as number) : String(first).localeCompare(String(second))) * (sortAsc ? 1 : -1);
    });
  }, [availableSections, selectedSections, sortAsc, sortKey]);

  const studentTotal = sections.reduce((sum, section) => sum + total(section), 0);
  const selectedCourseCount = new Set(sections.map((section) => section.course)).size;
  const overallA = studentTotal ? sections.reduce((sum, section) => sum + highGrades(section), 0) / studentTotal * 100 : 0;
  const overallPassing = studentTotal ? sections.reduce((sum, section) => sum + passingGrades(section), 0) / studentTotal * 100 : 0;
  const gradeBasedRisk = studentTotal ? sections.reduce((sum, section) => sum + section.grades["D+"] + section.grades["D&Below"], 0) / studentTotal * 100 : 0;
  const failingRateOverall = studentTotal ? sections.reduce((sum, section) => sum + failingGrades(section), 0) / studentTotal * 100 : 0;
  const withdrawalRate = studentTotal ? sections.reduce((sum, section) => sum + section.grades.W, 0) / studentTotal * 100 : 0;
  const atRiskCourseGroups = sections.filter((section) => {
    const lowGradeCount = section.grades["D+"] + section.grades["D&Below"];
    const value = riskType === "Percentage" ? percent(lowGradeCount, section) : lowGradeCount;
    return section.totalStudents >= minimumStudents && value >= riskThreshold;
  }).length;
  const chartSeries = useMemo<ChartSeries[]>(() => {
    const counters = { "Male Campus": 0, "Female Campus": 0 };
    return [...sections]
      .sort((a, b) => Number(a.campus.startsWith("Female")) - Number(b.campus.startsWith("Female")) || a.section.localeCompare(b.section, undefined, { numeric: true }))
      .map((record) => ({
        ...record,
        grades: { ...record.grades },
        label: sectionLabel(record),
        section: record.section,
        instructor: record.instructor,
        colorIndex: counters[record.campus]++,
      }));
  }, [sections]);
  const chartCampusGroups = useMemo(() => (["Male Campus", "Female Campus"] as const)
    .map((campus) => ({ campus, series: chartSeries.filter((item) => item.campus === campus) }))
    .filter((group) => group.series.length), [chartSeries]);
  const maxCount = Math.max(1, ...grades.map((grade) => chartSeries.reduce((sum, section) => sum + section.grades[grade], 0)));
  const maxPercentage = Math.max(10, Math.min(100, Math.ceil(Math.max(...grades.map((grade) => studentTotal ? chartSeries.reduce((sum, section) => sum + section.grades[grade], 0) / studentTotal * 100 : 0), 0) / 10) * 10));
  const comparisonEnabled = filters.semesters.length >= 2 && (filters.course !== "All" || filters.instructor !== "All");
  // Instructors who taught the selected course in the selected semesters (drives the chart's instructor dropdown).
  const chartInstructorOptions = useMemo(() => filters.course === "All" ? [] : [...new Set(gradeRecords.filter((record) =>
    filters.semesters.includes(record.semester) &&
    record.course === filters.course &&
    (filters.campus === "All" || record.campus.startsWith(filters.campus))
  ).flatMap((record) => record.instructor.split(/\s+\/\s+/)).map((name) => name.trim()).filter(Boolean))].sort(), [filters.campus, filters.course, filters.semesters, gradeRecords]);
  const activeChartInstructors = useMemo(() => chartInstructors.filter((name) => chartInstructorOptions.includes(name)), [chartInstructors, chartInstructorOptions]);
  const chartInstructorMatches = (recordInstructor: string) => !activeChartInstructors.length || activeChartInstructors.some((name) => instructorMatches(recordInstructor, name));
  const semesterComparison = useMemo<SemesterComparisonRecord[]>(() => {
    if (!comparisonEnabled) return [];
    return semesterOptions.filter((semester) => filters.semesters.includes(semester)).map((semester) => {
      const records = gradeRecords.filter((record) =>
        record.semester === semester &&
        (filters.course === "All" || record.course === filters.course) &&
        (filters.department === "All" || record.department === filters.department) &&
        instructorMatches(record.instructor, filters.instructor) &&
        (!activeChartInstructors.length || activeChartInstructors.some((name) => instructorMatches(record.instructor, name))) &&
        (filters.campus === "All" || record.campus.startsWith(filters.campus))
      );
      return {
        semester,
        label: semesterLabel(semester),
        color: semesterColors[semesterOptions.indexOf(semester) % semesterColors.length],
        sections: records.length,
        instructors: [...new Set(records.flatMap((record) => record.instructor.split(/\s+\/\s+/)).map((name) => name.trim()).filter(Boolean))].sort(),
        totalStudents: records.reduce((sum, record) => sum + record.totalStudents, 0),
        grades: Object.fromEntries(allGrades.map((grade) => [
          grade,
          records.reduce((sum, record) => sum + record.grades[grade], 0),
        ])) as Record<Grade, number>,
      };
    });
  }, [activeChartInstructors, comparisonEnabled, filters.campus, filters.course, filters.department, filters.instructor, filters.semesters, gradeRecords, semesterOptions]);
  const selectedSemesterLabels = filters.semesters.map(semesterLabel).join(", ");

  const rankedCourses = useMemo(() => {
    const grouped = new Map<string, ChartSeries>();
    sections.forEach((record) => {
      const current = grouped.get(record.course);
      if (current) {
        current.totalStudents += record.totalStudents;
        allGrades.forEach((grade) => { current.grades[grade] += record.grades[grade]; });
      } else {
        grouped.set(record.course, { ...record, grades: { ...record.grades }, label: record.course });
      }
    });
    return [...grouped.values()];
  }, [sections]);
  // Prefer the actual uploaded course-grade datasets for course reports.
  // For terms/campuses with section data only, retain section-derived aggregates.
  const uploadedCourseReportRecords = useMemo(() => courseGradeRecords.filter((record) =>
    filters.instructor === "All" &&
    filters.semesters.includes(record.semester) &&
    (filters.department === "All" || record.department === filters.department) &&
    (filters.course === "All" || record.course === filters.course) &&
    (filters.campus === "All" || record.campus.startsWith(filters.campus))
  ), [courseGradeRecords, filters.instructor, filters.semesters, filters.department, filters.course, filters.campus]);
  const courseReportRecords = useMemo(() => {
    const groups = new Map<string, ChartSeries>();
    const courseSourceKeys = new Set(uploadedCourseReportRecords.map((r) => `${r.semester}|${r.campus}|${r.course}`));
    const sectionFallback = sections.filter((r) => !courseSourceKeys.has(`${r.semester}|${r.campus}|${r.course}`));
    for (const record of [...uploadedCourseReportRecords, ...sectionFallback]) {
      const key = `${record.semester}|${record.course}`;
      const previous = groups.get(key);
      if (previous) {
        previous.totalStudents += record.totalStudents;
        allGrades.forEach((grade) => { previous.grades[grade] += record.grades[grade]; });
      } else {
        groups.set(key, { ...record, grades: { ...record.grades }, label: record.course });
      }
    }
    return [...groups.values()];
  }, [uploadedCourseReportRecords, sections]);
  // Instructor names come from the filtered section records, since uploaded
  // course-level aggregates may not contain an instructor assignment.
  const courseInstructorNames = useMemo(() => {
    const names = new Map<string, Set<string>>();
    for (const record of sections) {
      const key = `${record.semester}|${record.course}`;
      const courseNames = names.get(key) ?? new Set<string>();
      record.instructor.split(/\s+\/\s+/).map((name) => name.trim()).filter((name) => name && name !== "Unassigned").forEach((name) => courseNames.add(name));
      names.set(key, courseNames);
    }
    return new Map([...names.entries()].map(([course, instructors]) => [course, [...instructors].sort().join(", ") || "Not available"]));
  }, [sections]);
  const instructorsForCourse = (course: string, semester: string) => courseInstructorNames.get(`${semester}|${course}`) ?? "Not available";
  const reportRecords = reportMode === "course" ? courseReportRecords : sections;
  const reportPotentiallyEasy = useMemo(() => reportRecords
    .filter((record) => {
      const count = gradeRangeCount(record, highGradeRange);
      const value = aType === "Percentage" ? percent(count, record) : count;
      return record.totalStudents >= minimumStudents && value >= aThreshold;
    })
    .sort((a, b) => percent(gradeRangeCount(b, highGradeRange), b) - percent(gradeRangeCount(a, highGradeRange), a)), [aThreshold, aType, highGradeRange, minimumStudents, reportRecords]);
  const reportPotentiallyDifficult = useMemo(() => reportRecords
    .filter((record) => {
      const lowGradeCount = gradeRangeCount(record, lowGradeRange);
      const value = riskType === "Percentage" ? percent(lowGradeCount, record) : lowGradeCount;
      return record.totalStudents >= minimumStudents && value >= riskThreshold;
    })
    .sort((a, b) => percent(gradeRangeCount(b, lowGradeRange), b) - percent(gradeRangeCount(a, lowGradeRange), a)), [lowGradeRange, minimumStudents, reportRecords, riskThreshold, riskType]);
  const allNotableSectionDifferences = useMemo(() => {
    const grouped = new Map<string, SectionGradeRecord[]>();
    sections.forEach((record) => {
      const key = `${record.semester}|${record.course}`;
      grouped.set(key, [...(grouped.get(key) ?? []), record]);
    });
    return [...grouped.entries()].flatMap(([course, records]) => {
      if (records.length < 2) return [];
      const measures = [
        { label: "Combined A+/A", count: (record: SectionGradeRecord) => highGrades(record) },
        { label: "Combined D & F", count: (record: SectionGradeRecord) => record.grades["D&Below"] },
      ];
      return measures.map((measure) => {
        const ranked = records.map((record) => ({ record, rate: percent(measure.count(record), record) })).sort((a, b) => b.rate - a.rate);
        return { course: records[0].course, semester: records[0].semester, title: records[0].title, measure: measure.label, highest: ranked[0], lowest: ranked[ranked.length - 1], difference: ranked[0].rate - ranked[ranked.length - 1].rate };
      }).filter((item) => item.difference >= variationThreshold);
    }).sort((a, b) => b.difference - a.difference);
  }, [sections, variationThreshold]);
  const notableSectionDifferences = allNotableSectionDifferences;
  // One row per instructor picked in the chart dropdown: all of that instructor's sections of the course are totalled.
  const instructorBarRows = useMemo<InstructorComparisonRow[]>(() => activeChartInstructors.map((name) => {
    const row: InstructorComparisonRow = { instructor: name, sections: 0, totalStudents: 0, grades: Object.fromEntries(allGrades.map((grade) => [grade, 0])) as Record<Grade, number> };
    gradeRecords.filter((record) =>
      filters.semesters.includes(record.semester) &&
      record.course === filters.course &&
      (filters.campus === "All" || record.campus.startsWith(filters.campus)) &&
      instructorMatches(record.instructor, name)
    ).forEach((record) => {
      row.sections += 1;
      row.totalStudents += record.totalStudents;
      allGrades.forEach((grade) => { row.grades[grade] += record.grades[grade]; });
    });
    return row;
  }).filter((row) => row.sections > 0), [activeChartInstructors, filters.campus, filters.course, filters.semesters, gradeRecords]);
  // Bars (one per instructor, totalled over every selected semester) replace the line chart when picked in multi-semester view.
  const showInstructorBars = instructorBarRows.length > 0 && (!comparisonEnabled || multiSemesterView === "bars");
  const instructorVariations = useMemo(() => {
    const byCourse = new Map<string, SectionGradeRecord[]>();
    sections.forEach((record) => {
      const key = `${record.semester}|${record.course}`;
      byCourse.set(key, [...(byCourse.get(key) ?? []), record]);
    });
    const measures = [
      { label: "Combined A+/A", count: (value: Record<Grade, number>) => value["A+"] + value.A },
      { label: "Combined D & F", count: (value: Record<Grade, number>) => value["D&Below"] },
    ];
    return [...byCourse.values()].flatMap((records) => {
      const byInstructor = new Map<string, InstructorComparisonRow>();
      records.forEach((record) => {
        const row = byInstructor.get(record.instructor) ?? { instructor: record.instructor, sections: 0, totalStudents: 0, grades: Object.fromEntries(allGrades.map((grade) => [grade, 0])) as Record<Grade, number> };
        row.sections += 1;
        row.totalStudents += record.totalStudents;
        allGrades.forEach((grade) => { row.grades[grade] += record.grades[grade]; });
        byInstructor.set(record.instructor, row);
      });
      const instructorGroups = [...byInstructor.values()].filter((row) => row.totalStudents >= minimumStudents);
      if (instructorGroups.length < 2) return [];
      return measures.map((measure) => {
        const ranked = instructorGroups.map((row) => ({ row, rate: row.totalStudents ? measure.count(row.grades) / row.totalStudents * 100 : 0 })).sort((a, b) => b.rate - a.rate);
        return { course: records[0].course, semester: records[0].semester, title: records[0].title, measure: measure.label, highest: ranked[0], lowest: ranked[ranked.length - 1], difference: ranked[0].rate - ranked[ranked.length - 1].rate };
      }).filter((item) => item.difference >= variationThreshold);
    }).sort((a, b) => b.difference - a.difference);
  }, [sections, variationThreshold, minimumStudents]);
  // Instructor Comparison: every instructor's sections of a course are totalled across the selected semesters, with a count per grade.
  const instructorGradeTable = useMemo(() => {
    const groups = new Map<string, { course: string; title: string; instructor: string; sections: number; totalStudents: number; grades: Record<Grade, number> }>();
    sections.forEach((record) => {
      const key = `${record.course}|${record.instructor}`;
      const row = groups.get(key) ?? { course: record.course, title: record.title, instructor: record.instructor, sections: 0, totalStudents: 0, grades: Object.fromEntries(allGrades.map((grade) => [grade, 0])) as Record<Grade, number> };
      row.sections += 1;
      row.totalStudents += record.totalStudents;
      allGrades.forEach((grade) => { row.grades[grade] += record.grades[grade]; });
      groups.set(key, row);
    });
    return [...groups.values()].sort((a, b) => a.course.localeCompare(b.course) || a.instructor.localeCompare(b.instructor));
  }, [sections]);
  const topPerformingSections = useMemo(() => sections.filter((record) => {
    const count = gradeRangeCount(record, highGradeRange);
    const value = aType === "Percentage" ? percent(count, record) : count;
    return record.totalStudents >= minimumStudents && value >= aThreshold;
  })
    .sort((a, b) => percent(gradeRangeCount(b, highGradeRange), b) - percent(gradeRangeCount(a, highGradeRange), a))
    .slice(0, topSectionCount), [aThreshold, aType, highGradeRange, minimumStudents, sections, topSectionCount]);
  const atRiskSections = useMemo(() => sections.filter((record) => {
    const lowGradeCount = gradeRangeCount(record, lowGradeRange);
    const value = riskType === "Percentage" ? percent(lowGradeCount, record) : lowGradeCount;
    return record.totalStudents >= minimumStudents && value >= riskThreshold;
  }).sort((a, b) => percent(gradeRangeCount(b, lowGradeRange), b) - percent(gradeRangeCount(a, lowGradeRange), a))
    .slice(0, atRiskSectionCount), [atRiskSectionCount, lowGradeRange, minimumStudents, riskThreshold, riskType, sections]);
  const emailReportCourses = useMemo(() => {
    const records = reportMode === "course"
      ? [...reportPotentiallyEasy.slice(0, topCourseCount), ...reportPotentiallyDifficult.slice(0, atRiskCourseCount)]
      : [...topPerformingSections, ...atRiskSections];
    return [...new Set(records.map((record) => record.course))].sort();
  }, [atRiskCourseCount, atRiskSections, reportMode, reportPotentiallyDifficult, reportPotentiallyEasy, topCourseCount, topPerformingSections]);
  const reportFileName = (course: string) => `${course.replace(/\s+/g, "_")}_Grade_Report.${emailAttachment}`;
  const buildReportFile = (course: string, lines: string[], courseSections: SectionGradeRecord[]) => {
    // Wrap long lines so they stay inside the PDF page width.
    if (emailAttachment === "pdf") return createPdfBlob(lines.flatMap((line) => line.match(/.{1,90}(?:\s|$)/g)?.map((part) => part.trim()) ?? [line]));
    const escape = (value: string | number) => `"${String(value).replace(/"/g, '""')}"`;
    const header = ["Semester", "Course", "Title", "Section", "Campus", "Instructor", "Total Students", ...allGrades, "Passing Rate", "Failing Rate", "Withdrawal Rate"];
    const rows = courseSections.map((record) => [semesterLabel(record.semester), record.course, record.title, record.section, record.campus, record.instructor, record.totalStudents, ...allGrades.map((grade) => record.grades[grade]), passingRate(record).toFixed(1) + "%", failingRate(record).toFixed(1) + "%", withdrawalRateOf(record).toFixed(1) + "%"]);
    // BOM so Excel opens the CSV as UTF-8.
    return new Blob(["\ufeff" + [header, ...rows].map((row) => row.map(escape).join(",")).join("\r\n")], { type: "text/csv;charset=utf-8" });
  };
  const buildCourseEmail = (course: string) => {
    const courseSections = sections.filter((record) => record.course === course);
    const instructors = [...new Set(courseSections.map((record) => record.instructor))];
    const missingRecipients = instructors.filter((instructor) => !findProfessorEmail(instructor));
    const recipients = [...new Set(instructors.flatMap((instructor) => {
      const email = findProfessorEmail(instructor);
      return email ? [email] : [];
    }))];
    const highFlag = reportPotentiallyEasy.some((record) => record.course === course) || topPerformingSections.some((record) => record.course === course);
    const lowFlag = reportPotentiallyDifficult.some((record) => record.course === course) || atRiskSections.some((record) => record.course === course);
    const details = courseSections.map((record) => {
      const breakdown = allGrades.map((grade) => `${grade}: ${record.grades[grade]}`).join(", ");
      return `Section ${record.section} — ${record.instructor} — ${record.totalStudents} students\n${breakdown}`;
    }).join("\n\n");
    return {
      course,
      to: recipients,
      missingRecipients,
      subject: `${course} Grade Distribution Review`,
      body: `Report flag: ${[highFlag ? "High average performance" : "", lowFlag ? "Low average performance" : ""].filter(Boolean).join(" and ")}\n\nCourse: ${course} — ${courseSections[0]?.title ?? ""}\nInstructors: ${instructors.join(", ")}\n\nThe grade breakdown by section is in the attached ${emailAttachment.toUpperCase()} report: ${reportFileName(course)}.`,
      reportLines: [`PMU - CCES`, `Grade Distribution Review: ${course} — ${courseSections[0]?.title ?? ""}`, `Semesters: ${selectedSemesterLabels}`, `Report flag: ${[highFlag ? "High average performance" : "", lowFlag ? "Low average performance" : ""].filter(Boolean).join(" and ")}`, "", ...courseSections.flatMap((record) => [`Section ${record.section} (${record.campus}) — ${record.instructor} — ${record.totalStudents} students`, allGrades.map((grade) => `${grade}: ${record.grades[grade]} (${percent(record.grades[grade], record).toFixed(1)}%)`).join(", "), outcomeSplit(record), ""])],
      sections: courseSections,
    };
  };
  const sendReportEmails = async () => {
    setSendingEmails(true);
    try {
      const cc = chairEmails.split(/[;,]/).map((email) => email.trim()).filter(Boolean);
      const drafts = selectedEmailCourses.map(buildCourseEmail);
      const unresolved = drafts.flatMap((draft) => draft.missingRecipients);
      if (unresolved.length) {
        throw new Error(`No directory email found for: ${[...new Set(unresolved)].join(", ")}.`);
      }
      const messages = drafts.map(({ course, to, subject, body }) => ({ course, to, subject, body, cc }));
      const response = await fetch("/api/send-report", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ messages }) });
      const result = await response.json() as { error?: string; sent?: number };
      if (!response.ok) throw new Error(result.error ?? "Unable to send report emails.");
      showToast("success", `${result.sent ?? messages.length} report emails sent successfully.`);
      setShowEmailPanel(false);
    } catch (error) {
      showToast("error", error instanceof Error ? error.message : "Unable to send report emails.");
    } finally {
      setSendingEmails(false);
    }
  };
  const openReportDrafts = async () => {
    const cc = chairEmails.split(/[;,]/).map((email) => email.trim()).filter(Boolean);
    const drafts = selectedEmailCourses.map(buildCourseEmail);
    const unresolved = drafts.flatMap((draft) => draft.missingRecipients);
    if (unresolved.length) {
      showToast("error", `No directory email found for: ${[...new Set(unresolved)].join(", ")}.`);
      return;
    }
    // Outlook on the web does not reliably honour the CC parameter, so also put the CC list on the
    // clipboard. Copy first: the clipboard needs the click's user activation, which window.open keeps.
    let copied = false;
    try {
      await navigator.clipboard.writeText(cc.join("; "));
      copied = true;
    } catch { /* Clipboard may be unavailable. */ }
    // Outlook compose links cannot carry attachments, so each report file is downloaded
    // and has to be attached to its draft (drag it in from Downloads).
    let opened = 0;
    for (const [index, draft] of drafts.entries()) {
      const { to, subject, body, course, reportLines, sections: courseSections } = draft;
      window.setTimeout(() => downloadBlob(buildReportFile(course, reportLines, courseSections), reportFileName(course)), index * 500);
      if (openOutlookDraft({ to, cc, subject, body })) opened++;
    }
    if (opened < drafts.length) {
      showToast("error", `Opened ${opened} of ${drafts.length} drafts. Allow pop-ups for this site, then try again.`);
    } else {
      showToast("success", `${opened} Outlook draft${opened === 1 ? "" : "s"} opened and ${drafts.length} ${emailAttachment.toUpperCase()} file${drafts.length === 1 ? "" : "s"} downloaded: attach each file to its draft.${copied && cc.length ? " CC list copied: if CC is empty, paste it into the CC box." : ""}`);
    }
  };
  const toggleSort = (key: typeof sortKey) => {
    if (sortKey === key) setSortAsc(!sortAsc);
    else { setSortKey(key); setSortAsc(true); }
  };

  const courseReportLines = [
    "PMU - CCES",
    "Course Difficulty & Ease Analysis Report",
    `Scope: ${selectedCourseCount} courses, ${studentTotal} students`,
    `Filters: Semesters ${selectedSemesterLabels}; Department ${filters.department}; Course ${filters.course}; Instructor ${filters.instructor}; Campus ${filters.campus}; Sections ${selectedSections.length ? selectedSections.length : "All"}`,
    `Potentially easy threshold: A/A+ >= ${aThreshold}${aType === "Percentage" ? "%" : " students"}`,
    `Potentially difficult threshold: D+ & Below >= ${riskThreshold}${riskType === "Percentage" ? "%" : " students"}`,
    `Minimum students per course group: ${minimumStudents}`,
    "",
    `Courses with High Average Performance (${Math.min(reportPotentiallyEasy.length, topCourseCount)})`,
    ...reportPotentiallyEasy.slice(0, topCourseCount).map((record) => `${semesterLabel(record.semester)} | ${record.course} - ${record.title} | Instructor: ${instructorsForCourse(record.course, record.semester)} | ${gradeRangeCount(record, highGradeRange)} of ${record.totalStudents} students (${percent(gradeRangeCount(record, highGradeRange), record).toFixed(1)}%) in ${highGradeRange.join(", ")} | ${outcomeSplit(record)}`),
    "",
    `Courses with Low Average Performance (${Math.min(reportPotentiallyDifficult.length, atRiskCourseCount)})`,
    ...reportPotentiallyDifficult.slice(0, atRiskCourseCount).map((record) => {
      const lowGradeCount = gradeRangeCount(record, lowGradeRange);
      return `${semesterLabel(record.semester)} | ${record.course} - ${record.title} | Instructor: ${instructorsForCourse(record.course, record.semester)} | ${lowGradeCount} of ${record.totalStudents} (${percent(lowGradeCount, record).toFixed(1)}%) in ${lowGradeRange.join(", ")} | Passing: ${passingGrades(record)} of ${record.totalStudents} (${passingRate(record).toFixed(1)}%) | Failing: ${failingGrades(record)} of ${record.totalStudents} (${failingRate(record).toFixed(1)}%) | Withdrawal: ${record.grades.W} of ${record.totalStudents} (${percent(record.grades.W, record).toFixed(1)}%) | Grade Distribution: ${grades.map((grade) => `${grade}: ${record.grades[grade]}`).join(", ")}`;
    }),
    "",
    "These configurable rules are illustrative and are not official PMU academic standards.",
  ];
  const sectionReportLines = [
    "PMU - CCES",
    "Section-Level Grade Analysis Report",
    `Scope: ${sections.length} sections, ${studentTotal} students`,
    `Filters: Semesters ${selectedSemesterLabels}; Department ${filters.department}; Course ${filters.course}; Instructor ${filters.instructor}; Campus ${filters.campus}; Sections ${selectedSections.length ? selectedSections.length : "All"}`,
    "",
    `Sections with High Average Performance (${topPerformingSections.length})`,
    ...topPerformingSections.map((record, index) => `${index + 1}. ${semesterLabel(record.semester)} | ${record.course} - ${record.title} | Section ${record.section} | ${record.instructor} | ${record.totalStudents} students | Selected range ${percent(gradeRangeCount(record, highGradeRange), record).toFixed(1)}%`),
    "",
    `Sections with Low Average Performance (${atRiskSections.length})`,
    ...atRiskSections.map((record, index) => `${index + 1}. ${semesterLabel(record.semester)} | ${record.course} - ${record.title} | Section ${record.section} | ${record.instructor} | ${record.totalStudents} students | Selected range ${percent(gradeRangeCount(record, lowGradeRange), record).toFixed(1)}%`),
    "",
    "Variation measures: Combined A+/A and Combined D & F",
    `Section Grade Variation Analysis >= ${variationThreshold} percentage points (${notableSectionDifferences.length})`,
    ...notableSectionDifferences.map((item) => `${semesterLabel(item.semester)} | ${item.course} - ${item.title} | ${item.measure}: ${item.difference.toFixed(1)} pp | Section ${item.highest.record.section} (${item.highest.record.campus}, ${item.highest.record.totalStudents} students) ${item.highest.rate.toFixed(1)}% vs Section ${item.lowest.record.section} (${item.lowest.record.campus}, ${item.lowest.record.totalStudents} students) ${item.lowest.rate.toFixed(1)}%`),
  ];
  const instructorVariationLines = [
    "",
    `Instructor Grade Variation >= ${variationThreshold} percentage points (${instructorVariations.length})`,
    ...instructorVariations.map((item) => `${semesterLabel(item.semester)} | ${item.course} - ${item.title} | ${item.measure}: ${item.difference.toFixed(1)} pp | ${item.highest.row.instructor} (${item.highest.row.sections} sections, ${item.highest.row.totalStudents} students) ${item.highest.rate.toFixed(1)}% vs ${item.lowest.row.instructor} (${item.lowest.row.sections} sections, ${item.lowest.row.totalStudents} students) ${item.lowest.rate.toFixed(1)}%`),
  ];
  const instructorComparisonLines = [
    "",
    `Instructor Comparison (${instructorGradeTable.length} instructor-course rows, totalled across the selected semesters)`,
    ...instructorGradeTable.map((row) => `${row.course} - ${row.title} | ${row.instructor} | ${row.sections} sections, ${row.totalStudents} students | ${grades.map((grade) => `${comparisonGradeLabel(grade)} ${row.grades[grade]} (${row.totalStudents ? (row.grades[grade] / row.totalStudents * 100).toFixed(1) : "0.0"}%)`).join(", ")}`),
  ];
  const reportLines = reportMode === "course" ? [...courseReportLines, ...instructorVariationLines, ...instructorComparisonLines] : [...sectionReportLines, ...instructorVariationLines, ...instructorComparisonLines];

  const exportExcel = () => {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([
      { Setting: "College", Value: "CCES" },
      { Setting: "Report View", Value: reportMode === "course" ? "By Course" : "By Section" },
      { Setting: "Semesters", Value: selectedSemesterLabels },
      { Setting: "Department", Value: filters.department },
      { Setting: "Course", Value: filters.course },
      { Setting: "Instructor", Value: filters.instructor },
      { Setting: "Campus", Value: filters.campus },
      { Setting: "Sections", Value: selectedSections.length ? selectedSections.length : "All Sections" },
      { Setting: "High Average Performance Threshold", Value: aThreshold, Type: aType },
      { Setting: "Low Average Performance Threshold", Value: riskThreshold, Type: riskType },
      { Setting: "Minimum Students", Value: minimumStudents },
      { Setting: "High-performance Courses to Display", Value: topCourseCount },
      { Setting: "Low-performance Courses to Display", Value: atRiskCourseCount },
      { Setting: "High-performance Sections to Display", Value: topSectionCount },
      { Setting: "Low-performance Sections to Display", Value: atRiskSectionCount },
      { Setting: "Section Variation Threshold", Value: variationThreshold, Type: "Percentage Points" },
    ]), "Criteria");
    const appendInstructorComparisonSheet = () => {
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(instructorGradeTable.map((row) => ({
        Course: row.course,
        "Course Name": row.title,
        Instructor: row.instructor,
        Sections: row.sections,
        Students: row.totalStudents,
        ...Object.fromEntries(grades.flatMap((grade) => [
          [`${comparisonGradeLabel(grade)} Count`, row.grades[grade]],
          [`${comparisonGradeLabel(grade)} %`, row.totalStudents ? row.grades[grade] / row.totalStudents : 0],
        ])),
      }))), "Instructor Comparison");
    };
    const appendInstructorVariationSheet = () => {
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(instructorVariations.map((item) => ({
        Semester: semesterLabel(item.semester),
        Course: item.course,
        "Course Name": item.title,
        "Grade Measure": item.measure,
        "Difference (percentage points)": item.difference,
        "Highest Instructor": item.highest.row.instructor,
        "Highest Instructor Sections": item.highest.row.sections,
        "Highest Instructor Students": item.highest.row.totalStudents,
        "Highest Percentage": item.highest.rate / 100,
        "Lowest Instructor": item.lowest.row.instructor,
        "Lowest Instructor Sections": item.lowest.row.sections,
        "Lowest Instructor Students": item.lowest.row.totalStudents,
        "Lowest Percentage": item.lowest.rate / 100,
      }))), "Instructor Variation");
    };
    if (reportMode === "section") {
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(topPerformingSections.map((record, index) => ({
        Rank: index + 1,
        Semester: semesterLabel(record.semester),
        Course: record.course,
        "Course Name": record.title,
        Section: record.section,
        Instructor: record.instructor,
        Students: record.totalStudents,
        "A/A+ Count": highGrades(record),
        "A/A+ Percentage": percent(highGrades(record), record) / 100,
      }))), "Top Sections");
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(atRiskSections.map((record, index) => ({
        Rank: index + 1,
        Semester: semesterLabel(record.semester),
        Course: record.course,
        "Course Name": record.title,
        Section: record.section,
        Instructor: record.instructor,
        Students: record.totalStudents,
        "D+ & Below Count": record.grades["D+"] + record.grades["D&Below"],
        "D+ & Below Rate": percent(record.grades["D+"] + record.grades["D&Below"], record) / 100,
      }))), "At Risk Sections");
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(notableSectionDifferences.map((item) => ({
        Semester: semesterLabel(item.semester),
        Course: item.course,
        "Course Name": item.title,
        "Grade Measure": item.measure,
        "Difference (percentage points)": item.difference,
        "Highest Section": item.highest.record.section,
        "Highest Campus": item.highest.record.campus,
        "Highest Section Students": item.highest.record.totalStudents,
        "Highest Percentage": item.highest.rate / 100,
        "Lowest Section": item.lowest.record.section,
        "Lowest Campus": item.lowest.record.campus,
        "Lowest Section Students": item.lowest.record.totalStudents,
        "Lowest Percentage": item.lowest.rate / 100,
      }))), "Notable Differences");
      appendInstructorVariationSheet();
      appendInstructorComparisonSheet();
      XLSX.writeFileXLSX(workbook, "PMU_CCES_Section_Grade_Analysis_Report.xlsx");
      return;
    }
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(reportPotentiallyEasy.map((record) => ({
      Semester: semesterLabel(record.semester),
      Department: record.department,
      Course: record.course,
      Title: record.title,
      Instructor: instructorsForCourse(record.course, record.semester),
      Students: record.totalStudents,
      "A/A+ Count": highGrades(record),
      "A/A+ Rate": percent(highGrades(record), record) / 100,
      "Passing Rate": passingRate(record) / 100,
      "Failing Rate": failingRate(record) / 100,
      "Withdrawal Rate": withdrawalRateOf(record) / 100,
      ...record.grades,
    }))), "High Performance");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(reportPotentiallyDifficult.map((record) => {
      const lowGradeCount = record.grades["D+"] + record.grades["D&Below"];
      return {
        Department: record.department,
        Course: record.course,
        Title: record.title,
        Instructor: instructorsForCourse(record.course, record.semester),
        Students: record.totalStudents,
        "D+ & Below Count": lowGradeCount,
        "D+ & Below Rate": percent(lowGradeCount, record) / 100,
        "Passing Rate": passingRate(record) / 100,
        "Failing Rate": failingRate(record) / 100,
        "Withdrawal Rate": percent(record.grades.W, record) / 100,
        ...record.grades,
      };
    })), "Low Performance");
    appendInstructorVariationSheet();
    appendInstructorComparisonSheet();
    XLSX.writeFileXLSX(workbook, "PMU_CCES_Course_Difficulty_Ease_Report.xlsx");
  };

  const switchPanel = (next: "analysis" | "report") => {
    if (next === activePanel) return;
    panelCourseRef.current[activePanel] = filters.course;
    const saved = panelCourseRef.current[next];
    const target = saved !== null && (saved === "All" || gradeRecords.some((record) => filters.semesters.includes(record.semester) && record.course === saved))
      ? saved
      : next === "analysis" ? pickDefaultCourse(gradeRecords, filters.semesters, filters.department) : "All";
    setActivePanel(next);
    if (target === filters.course) return;
    setSelectedSections([]);
    setReportGenerated(false);
    setFilters((current) => ({
      ...current,
      course: target,
      instructor: current.instructor === "All" || gradeRecords.some((record) => current.semesters.includes(record.semester) && (target === "All" || record.course === target) && instructorMatches(record.instructor, current.instructor)) ? current.instructor : "All",
    }));
  };

  const toggleSemester = (semester: string) => {
    const selected = filters.semesters.includes(semester);
    if (selected && filters.semesters.length === 1) return;
    const semesters = selected
      ? filters.semesters.filter((value) => value !== semester)
      : [...filters.semesters, semester].sort();
    const courseAvailable = filters.course === "All" || gradeRecords.some((record) => semesters.includes(record.semester) && record.course === filters.course);
    setSelectedSections([]);
    setReportGenerated(false);
    setFilters((current) => ({
      ...current,
      semesters,
      course: (() => {
        if (courseAvailable) return current.course;
        if (current.instructor !== "All") return "All";
        if (semesters.length < 2) return defaultCourse(gradeRecords, semesters, current.department);
        const counts = new Map<string, Set<string>>();
        gradeRecords.forEach((record) => {
          if (!semesters.includes(record.semester) || !CCES_DEPARTMENTS.has(record.department.toUpperCase())) return;
          if (!counts.has(record.course)) counts.set(record.course, new Set());
          counts.get(record.course)!.add(record.semester);
        });
        return [...counts.entries()].filter(([, terms]) => terms.size >= 2).map(([course]) => course).sort()[0] ?? defaultCourse(gradeRecords, semesters, current.department);
      })(),
      instructor: current.instructor,
      mode: semesters.length >= 2 && current.semesters.length < 2 ? "Percentage" : current.mode,
    }));
  };

  const renderFilterControls = () => (
    <section className="grade-filters grade-panel">
      <div className="grade-filter-field grade-semester-filter">
        <label>Semester</label>
        <details>
          <summary>{filters.semesters.length === 1 ? semesterLabel(filters.semesters[0]) : `${filters.semesters.length} semesters selected`}</summary>
          <div className="grade-semester-filter__menu">
            <strong>Select semesters</strong>
            {semesterOptions.map((semester) => <label key={semester}><input type="checkbox" checked={filters.semesters.includes(semester)} onChange={() => toggleSemester(semester)} /><span>{semesterLabel(semester)}</span></label>)}
          </div>
        </details>
      </div>
      <div className="grade-filter-field"><label>College</label><select value="CCES" disabled><option>CCES</option></select></div>
      <div className="grade-filter-field"><label>Department</label><select value={filters.department} onChange={(e) => { setSelectedSections([]); setReportGenerated(false); const department = e.target.value; setFilters((current) => ({ ...current, department, course: defaultCourse(gradeRecords, current.semesters, department), instructor: "All" })); }}><option value="All">All Departments</option>{departmentOptions.map((department) => <option key={department}>{department}</option>)}</select></div>
      <div className="grade-filter-field"><label>Course</label><select value={filters.course} onChange={(e) => { setSelectedSections([]); setReportGenerated(false); setFilters((current) => { const course = e.target.value; const instructorStillTeachesCourse = current.instructor === "All" || gradeRecords.some((record) => current.semesters.includes(record.semester) && (current.department === "All" || record.department.toUpperCase() === current.department) && (current.campus === "All" || record.campus.startsWith(current.campus)) && (course === "All" || record.course === course) && instructorMatches(record.instructor, current.instructor)); return { ...current, course, instructor: instructorStillTeachesCourse ? current.instructor : "All" }; }); }}><option value="All">All Courses</option>{courseOptions.filter((option) => filters.department === "All" || option.department === filters.department || option.value === filters.course).map((option) => <option key={option.value} value={option.value}>{option.course} — {option.title}</option>)}</select></div>
      <div className="grade-filter-field"><label>Instructor</label><select value={filters.instructor} onChange={(e) => { setReportGenerated(false); setSelectedSections([]); setFilters((current) => { const instructor = e.target.value; const eligible = gradeRecords.some((record) => current.semesters.includes(record.semester) && CCES_DEPARTMENTS.has(record.department.toUpperCase()) && (current.department === "All" || record.department.toUpperCase() === current.department) && (current.campus === "All" || record.campus.startsWith(current.campus)) && instructorMatches(record.instructor, instructor) && record.course === current.course); return { ...current, instructor, course: eligible ? current.course : "All" }; }); }}><option value="All">All Instructors</option>{instructorOptions.map((instructor) => <option key={instructor}>{instructor}</option>)}</select></div>
      <div className="grade-filter-field grade-section-filter">
        <label>Section</label>
        <details>
          <summary>{comparisonEnabled ? "All sections aggregated" : selectedSections.length ? `${selectedSections.length} Sections Selected` : "All Sections"}</summary>
          <div className="grade-section-filter__menu">
            {comparisonEnabled ? <span>Cross-semester comparison includes all matching sections.</span> : filters.course === "All" ? <span>Select a course to view sections.</span> : <>
              <button className={!selectedSections.length ? "active" : ""} onClick={() => { setSelectedSections([]); setReportGenerated(false); }}><i>{!selectedSections.length ? "✓" : ""}</i><span>All Sections</span><b>{availableSections.length}</b></button>
              {availableSections.map((record) => {
                const key = sectionKey(record);
                const active = selectedSections.includes(key);
                return <button className={active ? "active" : ""} key={key} onClick={() => { setReportGenerated(false); setSelectedSections((current) => active ? (current.length === 1 ? [] : current.filter((item) => item !== key)) : [...current, key]); }}><i>{active ? "✓" : ""}</i><span>{sectionLabel(record)}</span><b>{record.totalStudents}</b></button>;
              })}
            </>}
          </div>
        </details>
      </div>
      <div className="grade-filter-actions"><button className="grade-reset" onClick={() => { setSelectedSections([]); setReportGenerated(false); setFilters({ semesters: [gradeRecords[0]?.semester ?? "202510"], department: "All", course: defaultCourse(gradeRecords, [gradeRecords[0]?.semester ?? "202510"]), instructor: "All", campus: "All", mode: "Student Count" }); }}>Reset</button></div>
    </section>
  );

  return (
    <div className="grade-page">
      <section className="grade-hero">
        <div>
          <span className="grade-eyebrow"><MiniIcon name="chart" /> Academic performance intelligence</span>
          <h1>Grade Distribution Analysis</h1>
          <p>Analyze grades and compare course records across campuses.</p>
        </div>
        <div className="grade-hero-buttons">
          {/* [MERGE] Header button: toggles merged / original cross-listed courses */}
          <button className="grade-button grade-button--secondary grade-merge-action" disabled={isMerging || isImporting} onClick={() => void handleMergeCrossListed()}>
            {isMerging ? <span className="grade-button-spinner" aria-hidden="true" /> : <MiniIcon name="sections" />}
            {isMerging ? "Loading…" : mergeEnabled ? "Show Original Courses" : "Merge Cross-Listed Courses"}
          </button>
          <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
            Select any number of CSV files; semester and source detected automatically
          </label>
          <input
            ref={importInputRef}
            className="sr-only"
            type="file"
            accept=".csv,text/csv"
            multiple
            onChange={(event) => void handleGradeImport(event.target.files)}
          />
          <button className="grade-button grade-button--primary" disabled={isImporting} onClick={() => importInputRef.current?.click()}>
            {isImporting ? <span className="grade-button-spinner" aria-hidden="true" /> : <MiniIcon name="upload" />}
            {isImporting ? "Importing…" : "Import Files"}
          </button>
        </div>
      </section>
      {toast && <div className={`grade-toast grade-toast--${toast.type}`} role="status" aria-live="polite"><MiniIcon name={toast.type === "error" ? "alert" : toast.type === "success" ? "check" : "sections"} /><span>{toast.message}</span></div>}

      <div className="grade-view-tabs" role="tablist" aria-label="Grade analysis views">
        <button className={activePanel === "analysis" ? "active" : ""} onClick={() => switchPanel("analysis")} role="tab" aria-selected={activePanel === "analysis"}><MiniIcon name="chart" /><span>Analysis Dashboard<small>Explore distributions and course records</small></span></button>
        <button className={activePanel === "report" ? "active" : ""} onClick={() => switchPanel("report")} role="tab" aria-selected={activePanel === "report"}><MiniIcon name="sections" /><span>Report<small>Key findings and detailed summary</small></span></button>
      </div>

      {activePanel === "analysis" ? <>
      {renderFilterControls()}
      <section className="grade-summary">
        <GradeStat icon="users" label="Students in Selection" value={studentTotal.toLocaleString()} note={`Across ${selectedCourseCount} filtered ${selectedCourseCount === 1 ? "course" : "courses"}`} />
        <GradeStat icon="sections" label="Sections in Selection" value={String(sections.length)} note={selectedSections.length ? `${selectedSections.length} selected sections` : "All sections included"} />
        <GradeStat icon="award" label="A/A+ Rate" value={`${overallA.toFixed(1)}%`} note="Percentage of selected students receiving A or A+." />
        <GradeStat icon="award" label="Overall Passing Rate" value={`${overallPassing.toFixed(1)}%`} note="Students passing (A+ to D+, and P)" />
        <GradeStat icon="alert" label="Failing Rate" value={`${failingRateOverall.toFixed(1)}%`} note="Percentage of selected students receiving D & F." warning />
        <GradeStat icon="sections" label="At-Risk Course Groups" value={String(atRiskCourseGroups)} note="Based on risk threshold" warning />
        <GradeStat icon="alert" label="Academic Risk Rate" value={`${gradeBasedRisk.toFixed(1)}%`} note="Percentage of selected students receiving D+ or D & Below." warning />
        <GradeStat icon="alert" label="Withdrawal Rate" value={`${withdrawalRate.toFixed(1)}%`} note={`Pass ${overallPassing.toFixed(1)}% + Fail ${failingRateOverall.toFixed(1)}% + Withdraw ${withdrawalRate.toFixed(1)}% = ${(overallPassing + failingRateOverall + withdrawalRate).toFixed(1)}%`} warning />
      </section>

      <section className="grade-panel grade-chart-card">
        <div className="grade-card-heading">
          <div><h2>{showInstructorBars ? `${filters.course} Grade Distribution by Instructor` : comparisonEnabled ? `${filters.course !== "All" ? filters.course : filters.instructor} Grade Distribution Across Semesters` : "Grade Distribution Across Letter Grades"}</h2>{comparisonEnabled && !showInstructorBars && <p>All matching sections are aggregated within each semester before percentages are calculated{activeChartInstructors.length ? ` · instructors: ${activeChartInstructors.join(", ")}` : ""}.</p>}{showInstructorBars && <p>All sections each instructor taught for this course{filters.semesters.length > 1 ? ` across the ${filters.semesters.length} selected semesters` : ""} are totalled; percentages are of that instructor&apos;s own students.</p>}</div>
          <div className="grade-chart-actions">
            {filters.course !== "All" && chartInstructorOptions.length > 0 && <details className="grade-comparison-grade-filter">
              <summary>{activeChartInstructors.length ? `${activeChartInstructors.length} ${activeChartInstructors.length === 1 ? "instructor" : "instructors"}` : "All instructors"}</summary>
              <div>
                <strong>Compare instructors</strong>
                <label><input type="checkbox" checked={activeChartInstructors.length === 0} onChange={() => setChartInstructors([])} /><span>All instructors</span></label>
                {chartInstructorOptions.map((name) => <label key={name}><input type="checkbox" checked={activeChartInstructors.includes(name)} onChange={() => setChartInstructors((current) => { const base = current.filter((value) => chartInstructorOptions.includes(value)); return base.includes(name) ? base.filter((value) => value !== name) : [...base, name]; })} /><i style={{ background: instructorColors[chartInstructorOptions.indexOf(name) % instructorColors.length] }} /><span>{name}</span></label>)}
              </div>
            </details>}
            {comparisonEnabled && activeChartInstructors.length > 0 && <div className="grade-segment" role="tablist" aria-label="Chart type">{([["lines", "Lines"], ["bars", "Bars (total)"]] as const).map(([value, label]) => <button key={value} className={multiSemesterView === value ? "active" : ""} onClick={() => setMultiSemesterView(value)}>{label}</button>)}</div>}
            {comparisonEnabled && !showInstructorBars && <details className="grade-comparison-grade-filter">
              <summary>{comparisonGrades.length} grade categories</summary>
              <div>
                <strong>Display grades</strong>
                {grades.map((grade, index) => <label key={grade}><input type="checkbox" checked={comparisonGrades.includes(grade)} onChange={() => setComparisonGrades((current) => current.includes(grade) ? (current.length === 1 ? current : current.filter((value) => value !== grade)) : [...current, grade])} /><GradeKey grade={grade} /><span>{comparisonGradeLabel(grade)}</span></label>)}
              </div>
            </details>}
            <div className="grade-chart-toggles"><div className="grade-segment">{["All", "Male", "Female"].map((campus) => <button className={filters.campus === campus ? "active" : ""} onClick={() => { setReportGenerated(false); setFilters((current) => ({ ...current, campus })); }} key={campus}>{campus}</button>)}</div><MetricToggle mode={filters.mode} onChange={(mode) => setFilters((current) => ({ ...current, mode }))} /></div>
          </div>
        </div>
        {filters.semesters.length >= 2 && !comparisonEnabled && <p role="status">Choose a course or instructor to compare grade distributions across the selected semesters.</p>}
        {comparisonEnabled && !showInstructorBars ? <>
                    <SemesterComparisonChart data={semesterComparison} mode={filters.mode} selectedGrades={comparisonGrades} course={filters.course !== "All" ? filters.course : filters.instructor} />
        </> : showInstructorBars ? <InstructorBarsChart rows={instructorBarRows} allNames={chartInstructorOptions} mode={filters.mode} course={filters.course} /> : <>
          <div className="grade-chart-legend grade-chart-legend--single-line">{chartSeries.map((section) => <span key={chartSeriesKey(section)}><i style={{ background: chartSeriesColor(section) }} /><b>{section.course} · Sec {section.section}</b></span>)}</div>
          <div className="grade-chart-wrap">
            <span className="grade-y-label">{filters.mode}</span>
            <div className="grade-y-axis">{[4, 3, 2, 1, 0].map((step) => <span key={step}>{filters.mode === "Percentage" ? `${Math.round(maxPercentage * step / 4)}%` : Math.ceil(maxCount * step / 4)}</span>)}</div>
            <div className="grade-chart">
              {grades.map((grade, gradeIndex) => (
                <div className="grade-group" key={grade}>
                  <div className="grade-bars grade-bars--stacked">
                    <div className="grade-stack">
                    {chartSeries.map((section) => {
                        const count = section.grades[grade];
                        const pct = studentTotal ? count / studentTotal * 100 : 0;
                        const chartValue = filters.mode === "Percentage" ? pct / maxPercentage : count / maxCount;
                        const displayValue = filters.mode === "Percentage" ? `${pct.toFixed(1)}%` : count.toLocaleString();
                        return <div className={`grade-bar grade-bar--stacked grade-bar--${section.campus.startsWith("Male") ? "male" : "female"}`} key={chartSeriesKey(section)} style={{ height: `${Math.max(count ? 1 : 0, chartValue * 100)}%`, background: chartSeriesColor(section) }} tabIndex={0}><span className="grade-stack-value">{displayValue}</span><GradeTooltip record={section} grade={grade as Exclude<Grade, "P">} seriesLabel={section.label} instructor={section.instructor} selectionTotal={studentTotal} /></div>;
                      })}
                    </div>
                  </div>
                  <span className="grade-category"><strong style={{ color: gradeColors[gradeIndex] }}>{gradeLabels[grade as Exclude<Grade, "P">]}</strong></span>
                </div>
              ))}
            </div>
          </div>
          <div className="grade-x-label">Grade category</div>
        </>}
      </section>

      <section className="grade-panel grade-table-card">
        <div className="grade-card-heading"><div><h2>Course, Campus &amp; Section Breakdown</h2></div><span className="grade-table-count">{sections.length} section records</span></div>
        <div className="grade-table-scroll"><table><thead><tr>
          <th>Semester</th>
          <th><button onClick={() => toggleSort("course")}>Course {sortKey === "course" ? (sortAsc ? "↑" : "↓") : ""}</button></th>
          <th><button onClick={() => toggleSort("title")}>Course Title {sortKey === "title" ? (sortAsc ? "↑" : "↓") : ""}</button></th>
          <th><button onClick={() => toggleSort("campus")}>Campus {sortKey === "campus" ? (sortAsc ? "↑" : "↓") : ""}</button></th>
          <th>Section</th><th>Instructor</th>
          <th><button onClick={() => toggleSort("total")}>Total {sortKey === "total" ? (sortAsc ? "↑" : "↓") : ""}</button></th>
          {(["A+", "A", "D+", "D&Below", "W"] as Grade[]).map((grade) => <th key={grade}>{grade === "D&Below" ? "D & F" : grade}</th>)}<th>Record Chart</th>
        </tr></thead><tbody>{sections.map((section) => { const key = cohortKey(section); const expanded = expandedCohort === key; return <Fragment key={key}><tr><td><strong>{semesterLabel(section.semester)}</strong></td><td><strong>{section.course}</strong></td><td>{section.title}</td><td><span className={`grade-campus grade-campus--${section.campus.startsWith("Male") ? "male" : "female"}`}>{section.campus}</span></td><td><strong>{section.section}</strong></td><td>{section.instructor}</td><td><strong>{total(section)}</strong></td>{(["A+", "A", "D+", "D&Below", "W"] as Grade[]).map((grade) => <td key={grade}><strong>{section.grades[grade]}</strong><span>{percent(section.grades[grade], section).toFixed(1)}%</span></td>)}<td><button className={`grade-chart-toggle ${expanded ? "active" : ""}`} onClick={() => setExpandedCohort(expanded ? null : key)}><MiniIcon name="chart" />{expanded ? "Hide chart" : "View chart"}</button></td></tr>{expanded && <tr className="grade-expanded-row"><td colSpan={13}><CohortChart initialRecord={section} initialMode={filters.mode} /></td></tr>}</Fragment>; })}</tbody></table></div>
      </section>
      </> : <>
      <section className="grade-report-hero grade-panel">
        <div>
          <span className="grade-report-kicker">PMU · CCES</span>
          <h2>{reportMode === "course" ? "Course Difficulty & Ease Analysis Report" : "Section-Level Grade Analysis Report"}</h2>
          <p>{reportMode === "course" ? "Identify courses with grade patterns that may warrant academic review." : "Compare individual sections and identify notable differences requiring review."}</p>
        </div>
        <div className="grade-report-hero__meta">
          <span className="grade-report-scope">{selectedCourseCount} courses · {sections.length} sections · {studentTotal.toLocaleString()} students</span>
          <div className="grade-export-block"><span>{reportGenerated ? "Download generated report" : "Generate the report to enable downloads"}</span><div className="grade-export-actions"><button className="grade-button grade-button--secondary" disabled={!reportGenerated} onClick={() => downloadBlob(createPdfBlob(reportLines), reportMode === "course" ? "PMU_CCES_Course_Difficulty_Ease_Report.pdf" : "PMU_CCES_Section_Grade_Analysis_Report.pdf")}><MiniIcon name="download" /> PDF</button><button className="grade-button grade-button--secondary" disabled={!reportGenerated} onClick={exportExcel}><MiniIcon name="download" /> Excel</button><button className="grade-button grade-button--secondary" disabled={!reportGenerated} onClick={() => downloadBlob(createDocxBlob(reportLines), reportMode === "course" ? "PMU_CCES_Course_Difficulty_Ease_Report.docx" : "PMU_CCES_Section_Grade_Analysis_Report.docx")}><MiniIcon name="download" /> DOCX</button></div></div>
        </div>
      </section>
      <div className="grade-report-controls">
      <section className="grade-report-view-switch grade-panel" aria-labelledby="grade-report-view-label">
        <div>
          <span id="grade-report-view-label">Report grouping</span>
          <strong>{reportMode === "course" ? "Course-level analysis" : "Section-level analysis"}</strong>
        </div>
        <div className="grade-segment" role="tablist" aria-label="Report grouping">
          {([
            { value: "course", label: "By Course" },
            { value: "section", label: "By Section" },
          ] as const).map((option) => (
            <button
              type="button"
              role="tab"
              aria-selected={reportMode === option.value}
              className={reportMode === option.value ? "active" : ""}
              onClick={() => {
                setReportMode(option.value);
                setReportGenerated(false);
              }}
              key={option.value}
            >
              {option.label}
            </button>
          ))}
        </div>
      </section>
      <div className="grade-report-filter-row grade-report-filter-row--responsive">
        <style>{`
          .grade-report-filter-row--responsive { width: 100%; min-width: 0; max-width: 100%; overflow: visible; }
          .grade-report-filter-row--responsive .grade-filters {
            display: grid !important;
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
            width: 100% !important;
            min-width: 0 !important;
            max-width: 100% !important;
            box-sizing: border-box;
            gap: 14px !important;
            overflow: visible !important;
          }
          .grade-report-filter-row--responsive .grade-filter-field,
          .grade-report-filter-row--responsive .grade-filter-actions {
            min-width: 0 !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box;
          }
          .grade-report-filter-row--responsive select,
          .grade-report-filter-row--responsive details,
          .grade-report-filter-row--responsive summary {
            min-width: 0;
            max-width: 100%;
            width: 100%;
            box-sizing: border-box;
          }
          .grade-report-filter-row--responsive summary {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
          }
          @media (max-width: 850px) {
            .grade-report-filter-row--responsive .grade-filters {
              grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
            }
          }
          @media (max-width: 520px) {
            .grade-report-filter-row--responsive .grade-filters {
              grid-template-columns: minmax(0, 1fr) !important;
            }
          }
        `}</style>
        {renderFilterControls()}
      </div>
      <section className="grade-panel" style={{ padding: "14px 18px" }}>
        <strong>Available uploaded data</strong>
        <p style={{ margin: "8px 0 0" }}>
          {semesterOptions.length ? semesterOptions.map((semester) => {
            const sectionCount = gradeRecords.filter((record) => record.semester === semester).length;
            const courseCount = courseGradeRecords.filter((record) => record.semester === semester).length;
            return `${semesterLabel(semester)}: ${sectionCount} sections, ${courseCount} course records`;
          }).join(" · ") : "No uploaded semester data is available."}
        </p>
      </section>
      <section className="grade-panel grade-report grade-analysis-criteria">
        <div className="grade-report-section-heading"><span>01</span><div><h2>Analysis Criteria</h2><p>Choose when {reportMode === "course" ? "a course group" : "an individual section"} should be classified.</p></div></div>
        <div className="grade-classification-settings">
          <article className="grade-classification-rule grade-classification-rule--easy">
            <h3>{reportMode === "course" ? "Courses with High Average Performance" : "Sections with High Average Performance"}</h3>
            <div className="grade-classification-rule__controls">
              <label>Threshold Type<div className="grade-segment">{(["Percentage", "Student Count"] as DisplayMode[]).map((type) => <button className={aType === type ? "active" : ""} onClick={() => { setAType(type); setReportGenerated(false); }} key={type}>{type}</button>)}</div></label>
              <label>Threshold Value<div><input type="number" min="0" max={aType === "Percentage" ? 100 : undefined} value={aThreshold} onChange={(e) => { setAThreshold(Number(e.target.value)); setReportGenerated(false); }} /><span>{aType === "Percentage" ? "%" : "students"}</span></div></label>
              <label>Grade Range<GradeRangeSelect values={highGradeRange} options={["A+", "A", "B+", "B", "C+", "C", "D+"]} onChange={(values) => { setHighGradeRange(values); setReportGenerated(false); }} /></label>
            </div>
            <p>Flag {reportMode === "course" ? "courses" : "sections"} where at least <strong>{aThreshold}{aType === "Percentage" ? "%" : " students"}</strong> {aType === "Percentage" ? "of students " : ""}received A or A+.</p>
          </article>
          <article className="grade-classification-rule grade-classification-rule--difficult">
            <h3>{reportMode === "course" ? "Courses with Low Average Performance" : "Sections with Low Average Performance"}</h3>
            <div className="grade-classification-rule__controls">
              <label>Threshold Type<div className="grade-segment">{(["Percentage", "Student Count"] as DisplayMode[]).map((type) => <button className={riskType === type ? "active" : ""} onClick={() => { setRiskType(type); setReportGenerated(false); }} key={type}>{type}</button>)}</div></label>
              <label>Threshold Value<div><input type="number" min="0" max={riskType === "Percentage" ? 100 : undefined} value={riskThreshold} onChange={(e) => { setRiskThreshold(Number(e.target.value)); setReportGenerated(false); }} /><span>{riskType === "Percentage" ? "%" : "students"}</span></div></label>
              <label>Grade Range<GradeRangeSelect values={lowGradeRange} options={["C+", "C", "D+", "D", "F", "W"]} onChange={(values) => { setLowGradeRange(values); setReportGenerated(false); }} /></label>
            </div>
            <p>Flag {reportMode === "course" ? "courses" : "sections"} where at least <strong>{riskThreshold}{riskType === "Percentage" ? "%" : " students"}</strong> {riskType === "Percentage" ? "of students " : ""}received D+ or D &amp; Below.</p>
          </article>
        </div>
        <div className="grade-ranking-controls">
          <label className="grade-minimum-students">Minimum Students per {reportMode === "course" ? "Course Group" : "Section"}<div><input type="number" min="1" value={minimumStudents} onChange={(e) => { setMinimumStudents(Number(e.target.value)); setReportGenerated(false); }} /><span>students</span></div></label>
          {<label className="grade-minimum-students">{reportMode === "section" ? "Section & Instructor" : "Instructor"} Variation Threshold<div><input type="number" min="0" max="100" value={variationThreshold} onChange={(e) => { setVariationThreshold(Math.max(0, Number(e.target.value))); setReportGenerated(false); }} /><span>percentage points</span></div></label>}
        </div>
        <div className="grade-criteria-note"><MiniIcon name="alert" /><span>Report results show both student counts and percentages. These illustrative rules are not official PMU academic standards.</span></div>
        <div className="grade-report-action"><span>High range ≥ {aThreshold}{aType === "Percentage" ? "%" : " students"} · Low range ≥ {riskThreshold}{riskType === "Percentage" ? "%" : " students"} · Minimum {minimumStudents} students{` · Variation ≥ ${variationThreshold} pp`}</span><div><button className="grade-button grade-button--secondary" disabled={!reportGenerated} onClick={() => { setSelectedEmailCourses(emailReportCourses); setReviewEmailDrafts(false); setShowEmailPanel(true); }}><MiniIcon name="upload" /> Send report by email</button><button className="grade-button grade-button--primary" onClick={() => setReportGenerated(true)}><MiniIcon name="chart" /> Generate Report</button></div></div>
      </section>
      </div>

      {reportGenerated && <>
      <div className="grade-report-campus-view grade-panel"><span>Campus View</span><div className="grade-segment">{["All", "Male", "Female"].map((campus) => <button className={filters.campus === campus ? "active" : ""} onClick={() => setFilters((current) => ({ ...current, campus }))} key={campus}>{campus}</button>)}</div></div>
      {reportMode === "course" ? <>
      <section className="grade-panel grade-report grade-detailed-summary grade-signal-section--easy">
        <div className="grade-report-section-heading"><span>02</span><div><h2>Courses with High Average Performance</h2><p>Courses meeting the configured grade-range threshold.</p></div><select value={topCourseCount} onChange={(e) => setTopCourseCount(Number(e.target.value))}>{[5,10,15,20,9999].map((count) => <option key={count} value={count}>{count === 9999 ? "All" : `Top ${count}`}</option>)}</select><b>{Math.min(reportPotentiallyEasy.length, topCourseCount)} courses</b></div>
        <div className="grade-table-scroll"><table><thead><tr><th>Semester</th><th>Course</th><th>Instructor</th><th>Selected Grade Range Outcome</th><th>Pass / Fail / Withdraw</th><th>Grade Distribution</th></tr></thead><tbody>{reportPotentiallyEasy.slice(0, topCourseCount).map((record) => { const count = gradeRangeCount(record, highGradeRange); return <tr key={`easy-${record.semester}-${record.course}`}><td><strong>{semesterLabel(record.semester)}</strong></td><td><strong>{record.course} — {record.title}</strong></td><td>{instructorsForCourse(record.course, record.semester)}</td><td><span className="grade-signal grade-signal--easy">{count} of {record.totalStudents} students ({percent(count, record).toFixed(1)}%)</span></td><td>{outcomeSplit(record)}</td><td className="grade-distribution-summary">{grades.map((grade) => `${gradeLabels[grade as Exclude<Grade, "P">]} ${record.grades[grade]} (${percent(record.grades[grade], record).toFixed(1)}%)`).join(" · ")}</td></tr>; })}</tbody></table></div>
        {!reportPotentiallyEasy.length && <div className="grade-review-empty"><MiniIcon name="check" /><span><strong>No courses meet the potentially easy rule</strong>Adjust the criteria or broaden the filters to review another selection.</span></div>}
      </section>

      <section className="grade-panel grade-report grade-detailed-summary grade-signal-section--difficult">
        <div className="grade-report-section-heading"><span>03</span><div><h2>Courses with Low Average Performance</h2><p>Courses meeting the configured low-performance grade range.</p></div><select value={atRiskCourseCount} onChange={(e) => setAtRiskCourseCount(Number(e.target.value))}>{[5,10,15,20,9999].map((count) => <option key={count} value={count}>{count === 9999 ? "All" : `Top ${count}`}</option>)}</select><b>{Math.min(reportPotentiallyDifficult.length, atRiskCourseCount)} courses</b></div>
        <div className="grade-table-scroll"><table><thead><tr><th>Semester</th><th>Course</th><th>Instructor</th><th>Selected Grade Range Outcome</th><th>Passing Outcome</th><th>Failing Outcome</th><th>Withdrawal Outcome</th><th>Grade Distribution</th></tr></thead><tbody>{reportPotentiallyDifficult.slice(0, atRiskCourseCount).map((record) => { const lowGradeCount = gradeRangeCount(record, lowGradeRange); const passingCount = passingGrades(record); return <tr key={`difficult-${record.semester}-${record.course}`}><td><strong>{semesterLabel(record.semester)}</strong></td><td><strong>{record.course} — {record.title}</strong></td><td>{instructorsForCourse(record.course, record.semester)}</td><td><span className="grade-signal grade-signal--difficult">{lowGradeCount} of {record.totalStudents} students ({percent(lowGradeCount, record).toFixed(1)}%)</span></td><td>{passingCount} of {record.totalStudents} ({passingRate(record).toFixed(1)}%)</td><td>{failingGrades(record)} of {record.totalStudents} ({failingRate(record).toFixed(1)}%)</td><td>{record.grades.W} of {record.totalStudents} ({percent(record.grades.W, record).toFixed(1)}%)</td><td className="grade-distribution-summary">{grades.map((grade) => `${gradeLabels[grade as Exclude<Grade, "P">]} ${record.grades[grade]} (${percent(record.grades[grade], record).toFixed(1)}%)`).join(" · ")}</td></tr>; })}</tbody></table></div>
        {!reportPotentiallyDifficult.length && <div className="grade-review-empty"><MiniIcon name="check" /><span><strong>No courses meet the potentially difficult rule</strong>Adjust the criteria or broaden the filters to review another selection.</span></div>}
      </section>

      <section className="grade-panel grade-report grade-detailed-summary">
        <div className="grade-report-section-heading"><span>04</span><div><h2>Instructor Grade Variation</h2><p>Each instructor&apos;s sections of a course are totalled, then the highest and lowest instructors are compared on Combined A+/A and Combined D &amp; F (differences of {variationThreshold} percentage points or more; instructors with fewer than {minimumStudents} students are excluded).</p></div><b>{instructorVariations.length} findings</b></div>
        <div className="grade-table-scroll"><table><thead><tr><th>Semester</th><th>Course</th><th>Course Name</th><th>Grade Measure</th><th>Highest Instructor</th><th>Lowest Instructor</th><th>Difference</th></tr></thead><tbody>{instructorVariations.map((item) => <tr key={`${item.semester}-${item.course}-${item.measure}`}><td>{semesterLabel(item.semester)}</td><td><strong>{item.course}</strong></td><td>{item.title}</td><td><strong>{item.measure}</strong></td><td>{item.highest.row.instructor} · {item.highest.row.sections} {item.highest.row.sections === 1 ? "section" : "sections"} · {item.highest.row.totalStudents} students · {item.highest.rate.toFixed(1)}%</td><td>{item.lowest.row.instructor} · {item.lowest.row.sections} {item.lowest.row.sections === 1 ? "section" : "sections"} · {item.lowest.row.totalStudents} students · {item.lowest.rate.toFixed(1)}%</td><td><span className="grade-gap">{item.difference.toFixed(1)} pp</span></td></tr>)}</tbody></table></div>
        {!instructorVariations.length && <div className="grade-review-empty"><MiniIcon name="check" /><span><strong>No notable instructor variations</strong>No course has two instructors whose grade measures differ by {variationThreshold} percentage points or more.</span></div>}
      </section>
      <section className="grade-panel grade-report grade-detailed-summary">
        <div className="grade-report-section-heading"><span>05</span><div><h2>Instructor Comparison</h2><p>Every instructor&apos;s sections of a course are totalled across the selected semesters. Each cell shows the number of students and the share of that instructor&apos;s own students (for example, how many A+ grades an instructor gave in a course across all of their sections).</p></div><b>{instructorGradeTable.length} rows</b></div>
        <div className="grade-table-scroll"><table><thead><tr><th>Course</th><th>Instructor</th><th>Sections</th><th>Students</th>{grades.map((grade) => <th key={grade}>{comparisonGradeLabel(grade)}</th>)}</tr></thead><tbody>{instructorGradeTable.map((row) => <tr key={`${row.course}-${row.instructor}`}><td><strong>{row.course}</strong><br /><small>{row.title}</small></td><td>{row.instructor}</td><td>{row.sections}</td><td>{row.totalStudents}</td>{grades.map((grade) => <td key={grade}><strong>{row.grades[grade]}</strong> <small>({row.totalStudents ? (row.grades[grade] / row.totalStudents * 100).toFixed(1) : "0.0"}%)</small></td>)}</tr>)}</tbody></table></div>
        {!instructorGradeTable.length && <div className="grade-review-empty"><MiniIcon name="check" /><span><strong>No grade records</strong>No sections match the current filters.</span></div>}
      </section>

      </> : <>
      <section className="grade-panel grade-report grade-detailed-summary grade-signal-section--easy">
        <div className="grade-report-section-heading"><span>02</span><div><h2>Sections with High Average Performance</h2><p>Sections meeting the configured grade-range threshold.</p></div><select value={topSectionCount} onChange={(e) => setTopSectionCount(Number(e.target.value))}>{[5,10,15,20,9999].map((count) => <option key={count} value={count}>{count === 9999 ? "All" : `Top ${count}`}</option>)}</select><b>{topPerformingSections.length} sections</b></div>
        <div className="grade-table-scroll"><table><thead><tr><th>Rank</th><th>Semester</th><th>Course / Section</th><th>Course Name</th><th>Instructor</th><th>Students</th><th>Selected Grade Range Outcome</th></tr></thead><tbody>{topPerformingSections.map((record, index) => { const count = gradeRangeCount(record, highGradeRange); return <tr key={`top-section-${cohortKey(record)}`}><td><strong>#{index + 1}</strong></td><td>{semesterLabel(record.semester)}</td><td><strong>{record.course} · Section {record.section}</strong></td><td>{record.title}</td><td>{record.instructor}</td><td>{record.totalStudents}</td><td><span className="grade-signal grade-signal--easy">{count} of {record.totalStudents} ({percent(count, record).toFixed(1)}%)</span></td></tr>; })}</tbody></table></div>
      </section>
      <section className="grade-panel grade-report grade-detailed-summary grade-signal-section--difficult">
        <div className="grade-report-section-heading"><span>03</span><div><h2>Sections with Low Average Performance</h2><p>Sections meeting the configured low-performance grade range.</p></div><select value={atRiskSectionCount} onChange={(e) => setAtRiskSectionCount(Number(e.target.value))}>{[5,10,15,20,9999].map((count) => <option key={count} value={count}>{count === 9999 ? "All" : `Top ${count}`}</option>)}</select><b>{atRiskSections.length} sections</b></div>
        <div className="grade-table-scroll"><table><thead><tr><th>Rank</th><th>Semester</th><th>Course / Section</th><th>Course Name</th><th>Instructor</th><th>Students</th><th>Selected Grade Range Outcome</th></tr></thead><tbody>{atRiskSections.map((record, index) => { const lowGrades = gradeRangeCount(record, lowGradeRange); return <tr key={`risk-section-${cohortKey(record)}`}><td><strong>#{index + 1}</strong></td><td>{semesterLabel(record.semester)}</td><td><strong>{record.course} · Section {record.section}</strong></td><td>{record.title}</td><td>{record.instructor}</td><td>{record.totalStudents}</td><td><span className="grade-signal grade-signal--difficult">{lowGrades} of {record.totalStudents} ({percent(lowGrades, record).toFixed(1)}%)</span></td></tr>; })}</tbody></table></div>
      </section>
      <section className="grade-panel grade-report grade-detailed-summary">
        <div className="grade-report-section-heading"><span>04</span><div><h2>Section Grade Variation Analysis</h2><p>Combined A+/A and Combined D &amp; F · highest-to-lowest differences of {variationThreshold} percentage points or more.</p></div><b>{notableSectionDifferences.length} findings</b></div>
        <div className="grade-table-scroll"><table><thead><tr><th>Semester</th><th>Course</th><th>Course Name</th><th>Grade Measure</th><th>Highest Section</th><th>Lowest Section</th><th>Difference</th></tr></thead><tbody>{notableSectionDifferences.map((item) => <tr key={`${item.semester}-${item.course}-${item.measure}`}><td>{semesterLabel(item.semester)}</td><td><strong>{item.course}</strong></td><td>{item.title}</td><td><strong>{item.measure}</strong></td><td>Section {item.highest.record.section} · {item.highest.record.campus} · {item.highest.record.totalStudents} students · {item.highest.rate.toFixed(1)}%</td><td>Section {item.lowest.record.section} · {item.lowest.record.campus} · {item.lowest.record.totalStudents} students · {item.lowest.rate.toFixed(1)}%</td><td><span className="grade-gap">{item.difference.toFixed(1)} pp</span></td></tr>)}</tbody></table></div>
        {!notableSectionDifferences.length && <div className="grade-review-empty"><MiniIcon name="check" /><span><strong>No notable section variations</strong>No grade measure differs by {variationThreshold} percentage points or more.</span></div>}
      </section>
      <section className="grade-panel grade-report grade-detailed-summary">
        <div className="grade-report-section-heading"><span>05</span><div><h2>Instructor Grade Variation</h2><p>Each instructor&apos;s sections of a course are totalled, then the highest and lowest instructors are compared on Combined A+/A and Combined D &amp; F (differences of {variationThreshold} percentage points or more; instructors with fewer than {minimumStudents} students are excluded).</p></div><b>{instructorVariations.length} findings</b></div>
        <div className="grade-table-scroll"><table><thead><tr><th>Semester</th><th>Course</th><th>Course Name</th><th>Grade Measure</th><th>Highest Instructor</th><th>Lowest Instructor</th><th>Difference</th></tr></thead><tbody>{instructorVariations.map((item) => <tr key={`${item.semester}-${item.course}-${item.measure}`}><td>{semesterLabel(item.semester)}</td><td><strong>{item.course}</strong></td><td>{item.title}</td><td><strong>{item.measure}</strong></td><td>{item.highest.row.instructor} · {item.highest.row.sections} {item.highest.row.sections === 1 ? "section" : "sections"} · {item.highest.row.totalStudents} students · {item.highest.rate.toFixed(1)}%</td><td>{item.lowest.row.instructor} · {item.lowest.row.sections} {item.lowest.row.sections === 1 ? "section" : "sections"} · {item.lowest.row.totalStudents} students · {item.lowest.rate.toFixed(1)}%</td><td><span className="grade-gap">{item.difference.toFixed(1)} pp</span></td></tr>)}</tbody></table></div>
        {!instructorVariations.length && <div className="grade-review-empty"><MiniIcon name="check" /><span><strong>No notable instructor variations</strong>No course has two instructors whose grade measures differ by {variationThreshold} percentage points or more.</span></div>}
      </section>
      <section className="grade-panel grade-report grade-detailed-summary">
        <div className="grade-report-section-heading"><span>06</span><div><h2>Instructor Comparison</h2><p>Every instructor&apos;s sections of a course are totalled across the selected semesters. Each cell shows the number of students and the share of that instructor&apos;s own students (for example, how many A+ grades an instructor gave in a course across all of their sections).</p></div><b>{instructorGradeTable.length} rows</b></div>
        <div className="grade-table-scroll"><table><thead><tr><th>Course</th><th>Instructor</th><th>Sections</th><th>Students</th>{grades.map((grade) => <th key={grade}>{comparisonGradeLabel(grade)}</th>)}</tr></thead><tbody>{instructorGradeTable.map((row) => <tr key={`${row.course}-${row.instructor}`}><td><strong>{row.course}</strong><br /><small>{row.title}</small></td><td>{row.instructor}</td><td>{row.sections}</td><td>{row.totalStudents}</td>{grades.map((grade) => <td key={grade}><strong>{row.grades[grade]}</strong> <small>({row.totalStudents ? (row.grades[grade] / row.totalStudents * 100).toFixed(1) : "0.0"}%)</small></td>)}</tr>)}</tbody></table></div>
        {!instructorGradeTable.length && <div className="grade-review-empty"><MiniIcon name="check" /><span><strong>No grade records</strong>No sections match the current filters.</span></div>}
      </section>
      </>}
      </>}
      </>}
      {showEmailPanel && <div className="grade-email-backdrop" onMouseDown={() => setShowEmailPanel(false)}>
        <aside className="grade-email-drawer" onMouseDown={(event) => event.stopPropagation()}>
          <div className="grade-email-head"><div><span>Email delivery</span><h2>Send grade reports</h2><p>Select courses and review each email draft here before sending.</p></div><button onClick={() => setShowEmailPanel(false)} aria-label="Close">×</button></div>
          {!reviewEmailDrafts ? <>
            <label className="grade-email-field">Sender<input value={senderEmail} onChange={(event) => setSenderEmail(event.target.value)} /></label>
            <label className="grade-email-field">Department chairs (CC)<textarea value={chairEmails} onChange={(event) => setChairEmails(event.target.value)} rows={3} /></label>
            <label className="grade-email-field">Report attachment<select value={emailAttachment} onChange={(event) => setEmailAttachment(event.target.value as "pdf" | "csv")}><option value="pdf">PDF report</option><option value="csv">CSV (opens in Excel)</option></select></label>
            <div className="grade-email-course-list">{emailReportCourses.map((course) => <label key={course}><input type="checkbox" checked={selectedEmailCourses.includes(course)} onChange={() => setSelectedEmailCourses((current) => current.includes(course) ? current.filter((item) => item !== course) : [...current, course])} /><span><strong>{course}</strong><small>{sections.find((record) => record.course === course)?.title}</small></span></label>)}</div>
            <button className="grade-button grade-button--primary" disabled={!selectedEmailCourses.length} onClick={() => setReviewEmailDrafts(true)}>Review {selectedEmailCourses.length} email {selectedEmailCourses.length === 1 ? "draft" : "drafts"}</button>
          </> : <>
            <button className="grade-email-back" onClick={() => setReviewEmailDrafts(false)}>Edit draft settings</button>
            <div className="grade-email-drafts">{selectedEmailCourses.map((course) => { const draft = buildCourseEmail(course); return <article key={course}><div><strong>{draft.subject}</strong><small>To: {draft.to.join(", ") || "No directory email found"}</small>{draft.missingRecipients.length > 0 && <small>Missing: {draft.missingRecipients.join(", ")}</small>}<small>CC: {chairEmails}</small><small>Attachment: {reportFileName(course)}</small></div><pre style={{ whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}>{draft.body}</pre></article>; })}</div>
            <button className="grade-button grade-button--primary" disabled={selectedEmailCourses.length === 0 || selectedEmailCourses.some((course) => buildCourseEmail(course).missingRecipients.length > 0)} onClick={openReportDrafts}>{`Open ${selectedEmailCourses.length} ${selectedEmailCourses.length === 1 ? "draft" : "drafts"} in Outlook`}</button>
          </>}
        </aside>
      </div>}
    </div>
  );
}
