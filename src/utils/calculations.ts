import { Activity, DashboardMetrics, Insight, Process } from "../data/types";
import { daysBetween, isPast } from "./date";

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

export function calculateMetrics(processes: Process[], activities: Activity[] = []): DashboardMetrics {
  const total = processes.length;
  const completed = processes.filter(isCompleted).length;
  const inProgress = processes.filter((process) => process.overallStatus === "In Progress").length;
  const notStarted = processes.filter((process) => process.overallStatus === "Not Started").length;
  const delayed = processes.filter(isDelayed).length;
  const blocked = processes.filter((process) => process.blocked).length;
  const atRisk = processes.filter((process) => process.health === "Amber" || process.priority === "Critical").length;
  const infrastructureActivities = activities.filter((activity) => activity.processId === "INFRA");
  const workItemCount = total + infrastructureActivities.length;
  const workProgress = processes.reduce((sum, process) => sum + process.progress, 0) + infrastructureActivities.reduce((sum, activity) => sum + activityProgress(activity), 0);
  const completion = workItemCount ? workProgress / workItemCount : 0;
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
