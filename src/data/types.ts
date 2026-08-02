export type Status = "Not Started" | "In Progress" | "On Hold" | "Completed" | "Cancelled" | string;
export type Health = "Green" | "Amber" | "Red" | "Gray" | string;
export type Priority = "Critical" | "High" | "Medium" | "Low" | string;

export interface Phase {
  phaseName: string;
  responsibility: string;
  status: Status;
  progress: number;
  durationDays: number;
  currentOwner: string;
  waitingFor: string;
  plannedStart?: string;
  plannedFinish?: string;
  actualStart?: string;
  actualFinish?: string;
  varianceDays: number;
  health: Health;
  blocked: boolean;
  blockerDescription: string;
  delayReason: string;
  nextAction: string;
  lastUpdated?: string;
  updatedBy: string;
}

export interface Process {
  processId: string;
  processName: string;
  department: string;
  businessOwner: string;
  currentStage: string;
  overallStatus: Status;
  progress: number;
  durationDays: number;
  priority: Priority;
  health: Health;
  currentOwner: string;
  waitingFor: string;
  plannedStart?: string;
  plannedFinish?: string;
  actualStart?: string;
  actualFinish?: string;
  dueDate?: string;
  completionDate?: string;
  varianceDays: number;
  blocked: boolean;
  blockerDescription: string;
  delayReason: string;
  nextAction: string;
  lastUpdated?: string;
  updatedBy: string;
  responsibility: string;
  phases: Phase[];
}

export interface Activity {
  activityId: string;
  processId: string;
  processName: string;
  phase: string;
  progress?: number;
  updateDate?: string;
  updatedBy: string;
  previousStatus: string;
  newStatus: string;
  updateDescription: string;
  nextAction: string;
  waitingFor: string;
  blocker: string;
  dueDate?: string;
}

export interface ImportWarning {
  code: string;
  message: string;
}

export interface ImportError {
  code: string;
  message: string;
}

export interface ProjectData {
  projectName: string;
  uploadedFileName: string;
  uploadedAt: string;
  overallProgress?: number;
  totalProcesses: number;
  processes: Process[];
  activities: Activity[];
  warnings: ImportWarning[];
}

export interface ValidationResult {
  ok: boolean;
  errors: ImportError[];
  warnings: ImportWarning[];
  processCount: number;
  phaseCount: number;
  activityCount: number;
  fileName: string;
  uploadedAt: string;
}

export interface FilterState {
  search: string;
  statuses: string[];
  stages: string[];
  health: string[];
  priorities: string[];
  departments: string[];
  owners: string[];
  waitingFor: string[];
  blocked: "all" | "yes" | "no";
  delayed: "all" | "yes" | "no";
  dueFrom: string;
  dueTo: string;
  updatedFrom: string;
  updatedTo: string;
}

export interface DashboardMetrics {
  total: number;
  completed: number;
  inProgress: number;
  notStarted: number;
  delayed: number;
  blocked: number;
  atRisk: number;
  completion: number;
}

export interface Insight {
  type: "Fact" | "Warning" | "Recommendation";
  title: string;
  evidence: string;
  processIds?: string[];
}
