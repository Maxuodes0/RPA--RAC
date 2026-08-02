import { Activity, DashboardMetrics, Insight, Process } from "../data/types";
import { daysBetween, isPast } from "./date";

const DAY_MS = 86_400_000;

function clampProgress(value: number) {
  return Math.max(0, Math.min(1, value));
}

function normalizedStatus(value: string) {
  return String(value || "").trim().toLowerCase();
}

export function isCompleted(process: Process) {
  return String(process.overallStatus).toLowerCase() === "completed" || process.progress >= 1;
}

export function isDelayed(process: Process) {
  if (isCompleted(process) || String(process.overallStatus).toLowerCase() === "cancelled") return false;
  return process.varianceDays > 0 || isPast(process.dueDate);
}

export function staleDays(process: Process) {
  if (isCompleted(process) || String(process.overallStatus).toLowerCase() === "cancelled") return undefined;
  return daysBetween(process.lastUpdated);
}

export function needsAttention(process: Process) {
  const stale = staleDays(process);
  return (
    process.blocked ||
    process.health === "Red" ||
    process.priority === "Critical" ||
    isDelayed(process) ||
    (stale !== undefined && stale > 7) ||
    (isDelayed(process) && !process.nextAction)
  );
}

function activityProgress(activity: Activity) {
  if (typeof activity.progress === "number") return activity.progress;
  return String(activity.newStatus).toLowerCase() === "completed" ? 1 : 0;
}

function dateSpanDays(start?: string, finish?: string): number {
  if (!start || !finish) return 0;
  const startDate = new Date(start);
  const finishDate = new Date(finish);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(finishDate.getTime())) return 0;
  return Math.max(1, Math.round((finishDate.getTime() - startDate.getTime()) / DAY_MS) + 1);
}

function scheduledProgress(start?: string, finish?: string, today = new Date()): number | undefined {
  if (!start || !finish) return undefined;
  const startDate = new Date(start);
  const finishDate = new Date(finish);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(finishDate.getTime())) return undefined;
  if (finishDate <= startDate) return today >= finishDate ? 1 : 0;
  if (today <= startDate) return 0;
  if (today >= finishDate) return 1;
  return clampProgress((today.getTime() - startDate.getTime()) / (finishDate.getTime() - startDate.getTime()));
}

function statusProgress(status: string, progress?: number) {
  const normalized = normalizedStatus(status);
  if (normalized === "completed") return 1;
  if (normalized === "cancelled") return 0;
  if (typeof progress === "number" && Number.isFinite(progress) && progress > 0) return clampProgress(progress);
  return undefined;
}

function phaseEstimate(process: Process, today = new Date()) {
  if (!process.phases.length) {
    return statusProgress(process.overallStatus, process.progress) ?? scheduledProgress(process.plannedStart, process.plannedFinish, today) ?? 0;
  }
  let totalDays = 0;
  let earnedDays = 0;
  process.phases.forEach((phase) => {
    const days = phase.durationDays || dateSpanDays(phase.plannedStart, phase.plannedFinish) || 1;
    const progress = statusProgress(phase.status, phase.progress) ?? scheduledProgress(phase.plannedStart, phase.plannedFinish, today) ?? 0;
    totalDays += days;
    earnedDays += days * progress;
  });
  return totalDays ? earnedDays / totalDays : 0;
}

function workItemWeight(process: Process) {
  const phaseDays = process.phases.reduce((sum, phase) => sum + (phase.durationDays || dateSpanDays(phase.plannedStart, phase.plannedFinish)), 0);
  return process.durationDays || phaseDays || dateSpanDays(process.plannedStart, process.plannedFinish) || 1;
}

function infrastructureWeight(activity: Activity) {
  return activity.durationDays || dateSpanDays(activity.plannedStart, activity.plannedFinish) || 1;
}

function infrastructureProgress(activity: Activity, today = new Date()) {
  return statusProgress(activity.newStatus, activity.progress) ?? scheduledProgress(activity.plannedStart, activity.plannedFinish, today) ?? activityProgress(activity);
}

export function calculateMetrics(processes: Process[], activities: Activity[] = []): DashboardMetrics {
  const total = processes.length;
  const completed = processes.filter(isCompleted).length;
  const inProgress = processes.filter((process) => process.overallStatus === "In Progress").length;
  const notStarted = processes.filter((process) => process.overallStatus === "Not Started").length;
  const delayed = processes.filter(isDelayed).length;
  const blocked = processes.filter((process) => process.blocked).length;
  const atRisk = processes.filter((process) => process.health === "Amber" || process.priority === "Critical").length;
  const infrastructureActivities = activities.filter((activity) => activity.processId === "INFRA");
  const processWork = processes.reduce(
    (acc, process) => {
      const weight = workItemWeight(process);
      acc.weight += weight;
      acc.earned += weight * phaseEstimate(process);
      return acc;
    },
    { weight: 0, earned: 0 },
  );
  const infrastructureWork = infrastructureActivities.reduce(
    (acc, activity) => {
      const weight = infrastructureWeight(activity);
      acc.weight += weight;
      acc.earned += weight * infrastructureProgress(activity);
      return acc;
    },
    { weight: 0, earned: 0 },
  );
  const workWeight = processWork.weight + infrastructureWork.weight;
  const completion = workWeight ? (processWork.earned + infrastructureWork.earned) / workWeight : 0;
  return { total, completed, inProgress, notStarted, delayed, blocked, atRisk, completion };
}

export function groupCount<T>(items: T[], selector: (item: T) => string | undefined) {
  const counts = new Map<string, number>();
  items.forEach((item) => {
    const key = selector(item) || "Not provided";
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  return Array.from(counts, ([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
}

export function bottlenecks(processes: Process[]) {
  const stages = groupCount(processes.filter((process) => !isCompleted(process)), (process) => process.currentStage);
  const waiting = groupCount(processes.filter((process) => !isCompleted(process)), (process) => process.waitingFor);
  const delayReasons = groupCount(processes.filter(isDelayed), (process) => process.delayReason);
  const blockers = processes.filter((process) => process.blocked);
  const oldestBlocker = blockers
    .slice()
    .sort((a, b) => (new Date(a.lastUpdated || 0).getTime() || 0) - (new Date(b.lastUpdated || 0).getTime() || 0))[0];
  const longestDelay = processes.slice().sort((a, b) => b.varianceDays - a.varianceDays)[0];
  const stale = processes.filter((process) => (staleDays(process) || 0) > 7);
  return {
    topStage: stages[0],
    topWaiting: waiting[0],
    topDelayReason: delayReasons[0],
    oldestBlocker,
    longestDelay,
    stale,
  };
}

export function buildInsights(processes: Process[]): Insight[] {
  const delayed = processes.filter(isDelayed);
  const blocked = processes.filter((process) => process.blocked);
  const criticalBlocked = blocked.filter((process) => process.priority === "Critical");
  const topStage = groupCount(processes.filter((process) => !isCompleted(process)), (process) => process.currentStage)[0];
  const topOwner = groupCount(processes.filter((process) => !isCompleted(process)), (process) => process.currentOwner)[0];
  const stale = processes.filter((process) => (staleDays(process) || 0) > 7);
  const insights: Insight[] = [
    {
      type: "Fact",
      title: `${delayed.length} processes are delayed.`,
      evidence: "A process is counted as delayed when Variance Days is positive or the due date has passed while it is not completed.",
      processIds: delayed.map((process) => process.processId),
    },
    {
      type: "Fact",
      title: `${blocked.length} processes are blocked.`,
      evidence: "This uses the official Blocked field from the uploaded Excel tracker.",
      processIds: blocked.map((process) => process.processId),
    },
  ];
  if (topStage) {
    insights.push({
      type: "Warning",
      title: `${topStage.name} contains the highest number of active processes.`,
      evidence: `${topStage.value} active processes are currently in this stage.`,
    });
  }
  if (topOwner && topOwner.name !== "Not provided") {
    insights.push({
      type: "Warning",
      title: `${topOwner.name} has the largest active workload.`,
      evidence: `${topOwner.value} active processes list this owner.`,
    });
  }
  if (criticalBlocked.length) {
    insights.push({
      type: "Recommendation",
      title: "Prioritize Critical blocked processes before assigning additional work.",
      evidence: `${criticalBlocked.length} Critical processes are blocked in the uploaded tracker.`,
      processIds: criticalBlocked.map((process) => process.processId),
    });
  }
  if (stale.length) {
    insights.push({
      type: "Recommendation",
      title: "Refresh stale active process updates.",
      evidence: `${stale.length} active processes have not been updated for more than 7 days.`,
      processIds: stale.map((process) => process.processId),
    });
  }
  return insights;
}
