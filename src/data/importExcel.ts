import * as XLSX from "xlsx";
import { Activity, ImportError, ImportWarning, Phase, Process, ProjectData, ValidationResult } from "./types";
import { toIsoDate } from "../utils/date";

const requiredSheets = ["Project Tracker", "Settings"];
const phaseNames = ["Assessment", "PDD Share", "PDD Approval", "Development", "UAT", "Go Live"];
const validPriorities = new Set(["", "Critical", "High", "Medium", "Low"]);
const validBlocked = new Set(["", "Yes", "No"]);

type Row = Record<string, unknown>;

function normalizeKey(value: unknown): string {
  return String(value ?? "")
    .replace(/[▼▽]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function canonicalHeader(value: unknown): string {
  const key = normalizeKey(value);
  const aliases: Record<string, string> = {
    "task id": "process id",
    infiermint: "process",
    "infiermint / phase": "process / phase",
  };
  return aliases[key] || key;
}

function canonicalPhase(value: unknown): string {
  const phase = String(value ?? "").replace(/\s+/g, " ").trim();
  const aliases: Record<string, string> = {
    "pdd approve": "PDD Approval",
  };
  return aliases[normalizeKey(phase)] || phase;
}

function findHeaderRow(rows: unknown[][], sheetName: string): number {
  if (sheetName !== "Project Tracker") return 0;
  const index = rows.findIndex((row) => {
    const headers = row.map(canonicalHeader);
    return headers.includes("process / phase") && headers.includes("status");
  });
  return index >= 0 ? index : 0;
}

function readSheet(workbook: XLSX.WorkBook, sheetName: string): Row[] {
  const sheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true });
  const headerIndex = findHeaderRow(rawRows, sheetName);
  const headers = (rawRows[headerIndex] || []).map(canonicalHeader);
  return rawRows
    .slice(headerIndex + 1)
    .map((values) => {
      const normalized: Row = {};
      headers.forEach((key, index) => {
        if (key) normalized[key] = values[index] ?? "";
      });
      return normalized;
    })
    .filter((row) => Object.values(row).some((value) => String(value ?? "").trim()))
    .filter((row) => canonicalHeader(text(row, "Process ID")) !== "process id" && canonicalHeader(text(row, "Process / Phase")) !== "process / phase");
}

function text(row: Row, ...keys: string[]): string {
  for (const key of keys) {
    const value = row[normalizeKey(key)];
    if (value !== undefined && value !== null && String(value).trim() !== "") return String(value).trim();
  }
  return "";
}

function numberValue(row: Row, ...keys: string[]): number {
  const raw = text(row, ...keys);
  if (!raw) return 0;
  if (raw.endsWith("%")) return Number(raw.replace("%", "")) / 100;
  const value = Number(raw);
  return Number.isFinite(value) ? value : 0;
}

function dateValue(row: Row, ...keys: string[]) {
  for (const key of keys) {
    const iso = toIsoDate(row[normalizeKey(key)]);
    if (iso) return iso;
  }
  return undefined;
}

function phaseValue(row: Row): string {
  return canonicalPhase(text(row, "Process / Phase", "Phase"));
}

function blockedValue(row: Row): boolean {
  return text(row, "Blocked").toLowerCase() === "yes";
}

function currentStage(phases: Phase[], status: string): string {
  const active = phases.find((phase) => ["In Progress", "On Hold"].includes(phase.status));
  if (active) return active.phaseName;
  if (String(status).toLowerCase() === "completed") return "Completed";
  const firstOpen = phases.find((phase) => phase.status !== "Completed" && phase.status !== "Cancelled");
  return firstOpen?.phaseName || "Not provided";
}

function mapPhase(row: Row): Phase {
  return {
    phaseName: phaseValue(row),
    responsibility: text(row, "Responsibility"),
    status: text(row, "Status"),
    progress: numberValue(row, "Progress"),
    currentOwner: text(row, "Current Owner"),
    waitingFor: text(row, "Waiting For"),
    plannedStart: dateValue(row, "Start", "Planned Start"),
    plannedFinish: dateValue(row, "Finish", "Planned Finish"),
    actualStart: dateValue(row, "Actual Start"),
    actualFinish: dateValue(row, "Actual Finish"),
    varianceDays: numberValue(row, "Variance Days"),
    health: text(row, "Health"),
    blocked: blockedValue(row),
    blockerDescription: text(row, "Blocker Description"),
    delayReason: text(row, "Delay Reason"),
    nextAction: text(row, "Next Action"),
    lastUpdated: dateValue(row, "Last Updated"),
    updatedBy: text(row, "Updated By"),
  };
}

function mapProcess(row: Row, phases: Phase[], index: number, warnings: ImportWarning[], usedProcessIds: Set<string>): Process {
  let processId = text(row, "Process ID");
  if (!processId) {
    let generatedIndex = index + 1;
    do {
      processId = `AUTO-${String(generatedIndex).padStart(3, "0")}`;
      generatedIndex += 1;
    } while (usedProcessIds.has(processId));
    warnings.push({ code: "generated-process-id", message: `Generated an internal Process ID for process ${index + 1} because Excel was blank.` });
  }
  usedProcessIds.add(processId);
  const status = text(row, "Status");
  const plannedFinish = dateValue(row, "Finish", "Planned Finish");
  const actualFinish = dateValue(row, "Actual Finish");
  return {
    processId,
    processName: text(row, "Process", "Process Name") || "Not provided",
    department: text(row, "Department"),
    businessOwner: text(row, "Business Owner"),
    currentStage: currentStage(phases, status),
    overallStatus: status || "Not provided",
    progress: numberValue(row, "Progress"),
    priority: text(row, "Priority") || "Medium",
    health: text(row, "Health"),
    currentOwner: text(row, "Current Owner"),
    waitingFor: text(row, "Waiting For"),
    plannedStart: dateValue(row, "Start", "Planned Start"),
    plannedFinish,
    actualStart: dateValue(row, "Actual Start"),
    actualFinish,
    dueDate: plannedFinish,
    completionDate: actualFinish,
    varianceDays: numberValue(row, "Variance Days"),
    blocked: blockedValue(row) || phases.some((phase) => phase.blocked),
    blockerDescription: text(row, "Blocker Description") || phases.find((phase) => phase.blockerDescription)?.blockerDescription || "",
    delayReason: text(row, "Delay Reason") || phases.find((phase) => phase.delayReason)?.delayReason || "",
    nextAction: text(row, "Next Action") || phases.find((phase) => phase.nextAction)?.nextAction || "",
    lastUpdated: dateValue(row, "Last Updated"),
    updatedBy: text(row, "Updated By"),
    responsibility: text(row, "Responsibility"),
    phases,
  };
}

function mapActivities(rows: Row[]): Activity[] {
  return rows
    .filter((row) => text(row, "Activity ID") && text(row, "Activity ID") !== "EXAMPLE-DELETE")
    .map((row) => ({
      activityId: text(row, "Activity ID"),
      processId: text(row, "Process ID"),
      processName: text(row, "Process Name"),
      phase: text(row, "Phase"),
      updateDate: dateValue(row, "Update Date"),
      updatedBy: text(row, "Updated By"),
      previousStatus: text(row, "Previous Status"),
      newStatus: text(row, "New Status"),
      updateDescription: text(row, "Update Description"),
      nextAction: text(row, "Next Action"),
      waitingFor: text(row, "Waiting For"),
      blocker: text(row, "Blocker"),
      dueDate: dateValue(row, "Due Date"),
    }));
}

export async function readWorkbook(file: File | ArrayBuffer): Promise<XLSX.WorkBook> {
  const buffer = file instanceof File ? await file.arrayBuffer() : file;
  return XLSX.read(buffer, { type: "array", cellDates: true, raw: true });
}

export function validateWorkbook(workbook: XLSX.WorkBook, fileName: string, uploadedAt = new Date().toISOString()): ValidationResult {
  const errors: ImportError[] = [];
  const warnings: ImportWarning[] = [];
  requiredSheets.forEach((sheet) => {
    if (!workbook.SheetNames.includes(sheet)) errors.push({ code: "missing-sheet", message: `Missing required sheet: ${sheet}` });
  });
  if (errors.length) return { ok: false, errors, warnings, processCount: 0, phaseCount: 0, activityCount: 0, fileName, uploadedAt };

  const rows = readSheet(workbook, "Project Tracker");
  const headers = Object.keys(rows[0] || {});
  ["process id", "process", "process / phase", "status", "responsibility", "next action"].forEach((header) => {
    if (!headers.includes(header)) errors.push({ code: "missing-column", message: `Missing required Project Tracker column: ${header}` });
  });

  const processRows = rows.filter((row) => phaseValue(row) === "Process Summary");
  const phaseRows = rows.filter((row) => phaseNames.includes(phaseValue(row)));
  const ids = processRows.map((row) => text(row, "Process ID")).filter(Boolean);
  const duplicate = ids.find((id, index) => ids.indexOf(id) !== index);
  if (duplicate) errors.push({ code: "duplicate-process-id", message: `Duplicate Process ID found: ${duplicate}` });

  rows.forEach((row, index) => {
    const priority = text(row, "Priority");
    const blocked = text(row, "Blocked");
    if (!validPriorities.has(priority)) errors.push({ code: "invalid-priority", message: `Invalid priority "${priority}" on row ${index + 2}` });
    if (!validBlocked.has(blocked)) errors.push({ code: "invalid-blocked", message: `Invalid blocked value "${blocked}" on row ${index + 2}` });
  });

  const activities = workbook.SheetNames.includes("Activity Log") ? mapActivities(readSheet(workbook, "Activity Log")) : [];
  return { ok: errors.length === 0, errors, warnings, processCount: processRows.length, phaseCount: phaseRows.length, activityCount: activities.length, fileName, uploadedAt };
}

export function normalizeWorkbook(workbook: XLSX.WorkBook, fileName: string, uploadedAt = new Date().toISOString()): ProjectData {
  const validation = validateWorkbook(workbook, fileName, uploadedAt);
  if (!validation.ok) {
    throw new Error(validation.errors.map((error) => error.message).join("\n"));
  }
  const warnings = [...validation.warnings];
  const rows = readSheet(workbook, "Project Tracker");
  const processes: Process[] = [];
  const usedProcessIds = new Set(rows.map((row) => text(row, "Process ID")).filter(Boolean));
  let currentSummary: Row | null = null;
  let currentPhases: Phase[] = [];

  const flush = () => {
    if (!currentSummary) return;
    processes.push(mapProcess(currentSummary, currentPhases, processes.length, warnings, usedProcessIds));
    currentSummary = null;
    currentPhases = [];
  };

  rows.forEach((row) => {
    const phase = phaseValue(row);
    if (phase === "Process Summary") {
      flush();
      currentSummary = row;
    } else if (phaseNames.includes(phase)) {
      if (!currentSummary) {
        warnings.push({ code: "orphan-phase", message: `Phase "${phase}" could not be linked to a parent process.` });
      } else {
        currentPhases.push(mapPhase(row));
      }
    }
  });
  flush();

  return {
    projectName: "RPA Project Control Center",
    uploadedFileName: fileName,
    uploadedAt,
    totalProcesses: processes.length,
    processes,
    activities: mapActivities(readSheet(workbook, "Activity Log")),
    warnings,
  };
}

export async function importExcel(file: File): Promise<{ validation: ValidationResult; data?: ProjectData }> {
  const uploadedAt = new Date().toISOString();
  const workbook = await readWorkbook(file);
  const validation = validateWorkbook(workbook, file.name, uploadedAt);
  if (!validation.ok) return { validation };
  return { validation, data: normalizeWorkbook(workbook, file.name, uploadedAt) };
}

export async function importDefaultExcel(): Promise<ProjectData> {
  const response = await fetch("/data/RPA_Project_Tracker_Web_Ready.xlsx");
  if (!response.ok) throw new Error("Default Excel tracker could not be loaded.");
  const workbook = await readWorkbook(await response.arrayBuffer());
  return normalizeWorkbook(workbook, "RPA_Project_Tracker_Web_Ready.xlsx", new Date().toISOString());
}
