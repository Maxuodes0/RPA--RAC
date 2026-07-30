import { FilterState, Process } from "../data/types";
import { isDelayed } from "./calculations";

export const defaultFilters: FilterState = {
  search: "",
  statuses: [],
  stages: [],
  health: [],
  priorities: [],
  departments: [],
  owners: [],
  waitingFor: [],
  blocked: "all",
  delayed: "all",
  dueFrom: "",
  dueTo: "",
  updatedFrom: "",
  updatedTo: "",
};

function inSet(value: string, selected: string[]) {
  return selected.length === 0 || selected.includes(value || "Not provided");
}

function dateInRange(value: string | undefined, from: string, to: string) {
  if (!from && !to) return true;
  if (!value) return false;
  const date = new Date(value).getTime();
  if (from && date < new Date(from).getTime()) return false;
  if (to && date > new Date(to).getTime() + 86_399_999) return false;
  return true;
}

export function applyFilters(processes: Process[], filters: FilterState) {
  const query = filters.search.trim().toLowerCase();
  return processes.filter((process) => {
    const haystack = [
      process.processId,
      process.processName,
      process.department,
      process.currentOwner,
      process.waitingFor,
      process.currentStage,
      process.overallStatus,
      process.blockerDescription,
      process.delayReason,
      process.nextAction,
    ]
      .join(" ")
      .toLowerCase();
    if (query && !haystack.includes(query)) return false;
    if (!inSet(process.overallStatus, filters.statuses)) return false;
    if (!inSet(process.currentStage, filters.stages)) return false;
    if (!inSet(process.health, filters.health)) return false;
    if (!inSet(process.priority, filters.priorities)) return false;
    if (!inSet(process.department || "Not provided", filters.departments)) return false;
    if (!inSet(process.currentOwner || "Not provided", filters.owners)) return false;
    if (!inSet(process.waitingFor || "Not provided", filters.waitingFor)) return false;
    if (filters.blocked === "yes" && !process.blocked) return false;
    if (filters.blocked === "no" && process.blocked) return false;
    if (filters.delayed === "yes" && !isDelayed(process)) return false;
    if (filters.delayed === "no" && isDelayed(process)) return false;
    if (!dateInRange(process.dueDate, filters.dueFrom, filters.dueTo)) return false;
    if (!dateInRange(process.lastUpdated, filters.updatedFrom, filters.updatedTo)) return false;
    return true;
  });
}

export function uniqueValues(processes: Process[], selector: (process: Process) => string | undefined) {
  return Array.from(new Set(processes.map((process) => selector(process) || "Not provided"))).sort();
}
