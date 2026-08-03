import * as XLSX from "xlsx";
import { Activity, ImportError, ImportWarning, Phase, Process, ProjectData, ValidationResult } from "./types";
import { toIsoDate } from "../utils/date";

export const DEFAULT_TRACKER_FILE = "RPA_Project_Tracker_Web_Ready.xlsx";
export const DEFAULT_TRACKER_LABEL = "RPA Tracker.xlsx";

const requiredSheets = ["Project Tracker"];
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
    "process / task": "process",
    phase: "process / phase",
    "duration days": "days",
    "planned start": "start",
    "planned finish": "finish",
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
  if (!["Project Tracker", "Process Summary"].includes(sheetName)) return 0;
  const index = rows.findIndex((row) => {
    const headers = row.map(canonicalHeader);
    return sheetName === "Project Tracker"
      ? headers.includes("process / phase") && headers.includes("status")
      : headers.includes("process") && headers.includes("status") && headers.includes("progress");
  });
  return index >= 0 ? index : 0;
}

function readSheet(workbook: XLSX.WorkBook, sheetName: string): Row[] {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
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
    if (value === 0) continue;
    if (value !== undefined && value !== null && String(value).trim() !== "") return String(value).trim();
  }
  return "";
}

function numberValue(row: Row, ...keys: string[]): number {
  return optionalNumberValue(row, ...keys) ?? 0;
}

function optionalNumberValue(row: Row, ...keys: string[]): number | undefined {
  const raw = text(row, ...keys);
  if (!raw) return undefined;
  if (raw.endsWith("%")) return Number(raw.replace("%", "")) / 100;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function progressValue(value: unknown): number | undefined {
  if (value === undefined || value === null || value === "") return undefined;
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return undefined;
    return value > 1 ? value / 100 : value;
  }
  const raw = String(value).trim();
  if (!raw) return undefined;
  const normalized = raw.endsWith("%") ? Number(raw.replace("%", "")) / 100 : Number(raw);
  return Number.isFinite(normalized) ? (normalized > 1 ? normalized / 100 : normalized) : undefined;
}

function dateValue(row: Row, ...keys: string[]) {
  for (const key of keys) {
    const iso = toIsoDate(row[normalizeKey(key)]);
    if (iso) return iso;
  }
  return undefined;
}

function varianceDaysValue(row: Row): number {
  const value = optionalNumberValue(row, "Variance Days");
  return value !== undefined && Math.abs(value) < 1000 ? value : 0;
}

function phaseValue(row: Row): string {
  return canonicalPhase(text(row, "Process / Phase", "Phase"));
}

function isProcessSummary(row: Row): boolean {
  return phaseValue(row) === "Process Summary";
}

function isPhaseRow(row: Row): boolean {
  const phase = phaseValue(row);
  return Boolean(phase) && phase !== "Process Summary";
}

function hasProcessId(row: Row): boolean {
  return Boolean(text(row, "Process ID"));
}

function blockedValue(row: Row): boolean {
  return text(row, "Blocked").toLowerCase() === "yes";
}

function processNameValue(row: Row, summaryRow?: Row): string {
  const summaryName = text(summaryRow || {}, "Process", "Process Name");
  const trackerName = text(row, "Process", "Process Name");
  return summaryName || trackerName || "Not provided";
}

function currentStage(phases: Phase[], status: string): string {
  const active = phases.find((phase) => ["in progress", "on hold"].includes(normalizeKey(phase.status)));
  if (active) return active.phaseName;
  if (String(status).toLowerCase() === "completed") return "Completed";
  const firstOpen = phases.find((phase) => !["completed", "cancelled"].includes(normalizeKey(phase.status)));
  return firstOpen?.phaseName || "Not provided";
}

function statusFromPhases(phases: Phase[]): string {
  if (!phases.length) return "Not Started";
  const statuses = phases.map((phase) => normalizeKey(phase.status));
  if (statuses.includes("on hold")) return "On Hold";
  if (statuses.includes("in progress")) return "In Progress";
  if (statuses.every((status) => status === "completed")) return "Completed";
  return "Not Started";
}

function progressFromPhases(phases: Phase[]): number {
  if (!phases.length) return 0;
  return phases.reduce((sum, phase) => sum + phase.progress, 0) / phases.length;
}

function healthFromPhases(phases: Phase[]): string {
  if (phases.some((phase) => phase.health === "Red")) return "Red";
  if (phases.some((phase) => phase.health === "Amber")) return "Amber";
  if (phases.some((phase) => phase.health === "Green")) return "Green";
  return "";
}

function mapPhase(row: Row): Phase {
  return {
    phaseName: phaseValue(row),
    responsibility: text(row, "Responsibility"),
    status: text(row, "Status"),
    progress: numberValue(row, "Progress"),
    durationDays: numberValue(row, "Days", "Duration Days", "Delivery Days"),
    currentOwner: text(row, "Current Owner"),
    waitingFor: text(row, "Waiting For"),
    plannedStart: dateValue(row, "Start", "Planned Start"),
    plannedFinish: dateValue(row, "Finish", "Planned Finish"),
    actualStart: dateValue(row, "Actual Start"),
    actualFinish: dateValue(row, "Actual Finish"),
    varianceDays: varianceDaysValue(row),
    health: text(row, "Health"),
    blocked: blockedValue(row),
    blockerDescription: text(row, "Blocker Description"),
    delayReason: text(row, "Delay Reason"),
    nextAction: text(row, "Next Action"),
    lastUpdated: dateValue(row, "Last Updated"),
    updatedBy: text(row, "Updated By"),
  };
}

function mapProcess(row: Row, phases: Phase[], index: number, warnings: ImportWarning[], usedProcessIds: Set<string>, summaryRow?: Row): Process {
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
  const status = text(summaryRow || {}, "Status") || text(row, "Status") || statusFromPhases(phases);
  const progress = optionalNumberValue(summaryRow || {}, "Progress") ?? optionalNumberValue(row, "Progress") ?? progressFromPhases(phases);
  const plannedFinish = dateValue(summaryRow || {}, "Finish", "Planned Finish") || dateValue(row, "Finish", "Planned Finish");
  const actualFinish = dateValue(row, "Actual Finish");
  return {
    processId,
    processName: processNameValue(row, summaryRow),
    department: text(row, "Department"),
    businessOwner: text(row, "Business Owner"),
    currentStage: currentStage(phases, status),
    overallStatus: status || "Not provided",
    progress,
    durationDays: optionalNumberValue(summaryRow || {}, "Days", "Duration Days", "Delivery Days") ?? optionalNumberValue(row, "Days", "Duration Days", "Delivery Days") ?? phases.reduce((sum, phase) => sum + phase.durationDays, 0),
    priority: text(row, "Priority") || "Medium",
    health: text(summaryRow || {}, "Health") || text(row, "Health") || healthFromPhases(phases),
    currentOwner: text(row, "Current Owner") || phases.find((phase) => ["in progress", "on hold"].includes(normalizeKey(phase.status)))?.currentOwner || text(summaryRow || {}, "Developer") || text(row, "Developer"),
    waitingFor: text(row, "Waiting For"),
    plannedStart: dateValue(summaryRow || {}, "Effective Start", "Auto Start", "Start", "Planned Start") || dateValue(row, "Start", "Planned Start"),
    plannedFinish,
    actualStart: dateValue(row, "Actual Start"),
    actualFinish,
    dueDate: plannedFinish,
    completionDate: actualFinish,
    varianceDays: varianceDaysValue(row) || Math.max(0, ...phases.map((phase) => phase.varianceDays)),
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
      phase: text(row, "Phase", "Process / Phase"),
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

function mapInfrastructureActivity(parentRow: Row, row: Row, index: number): Activity {
  const status = text(row, "Status");
  const blocker = text(row, "Blocker Description") || text(row, "Delay Reason") || text(row, "Blocked");
  const description = [phaseValue(row), status ? `is ${status}` : "", blocker && blocker !== "No" ? `- ${blocker}` : ""].filter(Boolean).join(" ");
  return {
    activityId: `INFRA-${String(index + 1).padStart(3, "0")}`,
    processId: "INFRA",
    processName: text(parentRow, "Process", "Process Name") || "Infrastructure",
    phase: phaseValue(row),
    progress: numberValue(row, "Progress"),
    durationDays: numberValue(row, "Days", "Duration Days", "Delivery Days"),
    plannedStart: dateValue(row, "Start", "Planned Start"),
    plannedFinish: dateValue(row, "Finish", "Planned Finish"),
    updateDate: dateValue(row, "Last Updated"),
    updatedBy: text(row, "Updated By"),
    previousStatus: "",
    newStatus: status,
    updateDescription: description || `${phaseValue(row)} is Not provided.`,
    nextAction: text(row, "Next Action"),
    waitingFor: text(row, "Waiting For"),
    blocker,
    dueDate: dateValue(row, "Finish", "Planned Finish"),
  };
}

function countInfrastructureActivities(rows: Row[]): number {
  let inInfrastructureBlock = false;
  let count = 0;
  rows.forEach((row) => {
    if (isProcessSummary(row)) {
      inInfrastructureBlock = !hasProcessId(row);
    } else if (inInfrastructureBlock && isPhaseRow(row)) {
      count += 1;
    }
  });
  return count;
}

function dashboardInfrastructureActivities(workbook: XLSX.WorkBook): Activity[] {
  const sheet = workbook.Sheets.Dashboard;
  if (!sheet?.["!ref"]) return [];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true });
  const headerRow = rows.find((row) => row.map(normalizeKey).includes("provide servers") && row.map(normalizeKey).includes("uipath platform access"));
  if (!headerRow) return [];
  const headerIndex = rows.indexOf(headerRow);
  const activityColumns = headerRow
    .map((header, index) => ({ header: String(header || "").trim(), index }))
    .filter(({ header }) => ["Provide Servers", "Team Access", "Application Setup", "UiPath Platform Access"].includes(header));
  const activities: Activity[] = [];
  rows.slice(headerIndex + 1).forEach((row) => {
    const processName = String(row[10] || "").trim();
    if (!processName) return;
    activityColumns.forEach(({ header, index }) => {
      const status = String(row[index] || "").trim();
      if (!status) return;
      const progress = progressValue(status === "Completed" ? 1 : status === "In Progress" ? 0.5 : 0) ?? 0;
      activities.push({
        activityId: `INFRA-${String(activities.length + 1).padStart(3, "0")}`,
        processId: "INFRA",
        processName,
        phase: header,
        progress,
        updatedBy: "Project Manager",
        previousStatus: "",
        newStatus: status,
        updateDescription: `${processName} - ${header} is ${status}.`,
        nextAction: "",
        waitingFor: "",
        blocker: "",
      });
    });
  });
  return activities;
}

function readDashboardOverallProgress(workbook: XLSX.WorkBook): number | undefined {
  const sheet = workbook.Sheets.Dashboard;
  if (!sheet?.["!ref"]) return undefined;
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
      if (normalizeKey(cell?.v) !== "overall progress") continue;
      const valueCell = sheet[XLSX.utils.encode_cell({ r: row + 1, c: col })];
      return progressValue(valueCell?.v) ?? progressValue(valueCell?.w);
    }
  }
  return undefined;
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

  const processRows = rows.filter((row) => isProcessSummary(row) && hasProcessId(row));
  const phaseRows = rows.filter(isPhaseRow);
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
  return { ok: errors.length === 0, errors, warnings, processCount: processRows.length, phaseCount: phaseRows.length, activityCount: activities.length + countInfrastructureActivities(rows) + dashboardInfrastructureActivities(workbook).length, fileName, uploadedAt };
}

export function normalizeWorkbook(workbook: XLSX.WorkBook, fileName: string, uploadedAt = new Date().toISOString()): ProjectData {
  const validation = validateWorkbook(workbook, fileName, uploadedAt);
  if (!validation.ok) {
    throw new Error(validation.errors.map((error) => error.message).join("\n"));
  }
  const warnings = [...validation.warnings];
  const rows = readSheet(workbook, "Project Tracker");
  const summaryRows = readSheet(workbook, "Process Summary").filter((row) => text(row, "Process"));
  const processes: Process[] = [];
  const infrastructureActivities: Activity[] = [];
  const usedProcessIds = new Set(rows.map((row) => text(row, "Process ID")).filter(Boolean));
  let currentSummary: Row | null = null;
  let currentPhases: Phase[] = [];
  let currentInfrastructureSummary: Row | null = null;

  const flush = () => {
    if (!currentSummary) return;
    processes.push(mapProcess(currentSummary, currentPhases, processes.length, warnings, usedProcessIds, summaryRows[processes.length]));
    currentSummary = null;
    currentPhases = [];
  };

  rows.forEach((row) => {
    const phase = phaseValue(row);
    if (isProcessSummary(row)) {
      flush();
      if (hasProcessId(row)) {
        currentSummary = row;
        currentInfrastructureSummary = null;
      } else {
        currentSummary = null;
        currentInfrastructureSummary = row;
      }
    } else if (isPhaseRow(row)) {
      if (!currentSummary) {
        if (currentInfrastructureSummary) {
          infrastructureActivities.push(mapInfrastructureActivity(currentInfrastructureSummary, row, infrastructureActivities.length));
        } else {
          warnings.push({ code: "orphan-phase", message: `Phase "${phase}" could not be linked to a parent process.` });
        }
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
    overallProgress: readDashboardOverallProgress(workbook),
    totalProcesses: processes.length,
    processes,
    activities: [...infrastructureActivities, ...dashboardInfrastructureActivities(workbook), ...mapActivities(readSheet(workbook, "Activity Log"))],
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
  const response = await fetch(`/data/${DEFAULT_TRACKER_FILE}`);
  if (!response.ok) throw new Error("Default Excel tracker could not be loaded.");
  const workbook = await readWorkbook(await response.arrayBuffer());
  return normalizeWorkbook(workbook, DEFAULT_TRACKER_LABEL, new Date().toISOString());
}
