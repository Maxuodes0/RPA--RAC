import { DashboardMetrics, Process, ProjectData } from "./types";
import { calculateMetrics } from "../utils/calculations";
import { importExcel } from "./importExcel";

export interface DataSource {
  getProcesses(): Process[];
  getProcessById(id: string): Process | undefined;
  getActivities(): ProjectData["activities"];
  getDashboardMetrics(): DashboardMetrics;
  importExcel(file: File): ReturnType<typeof importExcel>;
}

export class ExcelDataSource implements DataSource {
  constructor(private readonly data: ProjectData) {}

  getProcesses() {
    return this.data.processes;
  }

  getProcessById(id: string) {
    return this.data.processes.find((process) => process.processId === id);
  }

  getActivities() {
    return this.data.activities;
  }

  getDashboardMetrics() {
    return calculateMetrics(this.data.processes);
  }

  importExcel(file: File) {
    return importExcel(file);
  }
}
