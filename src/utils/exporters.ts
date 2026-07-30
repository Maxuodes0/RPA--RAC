import jsPDF from "jspdf";
import * as XLSX from "xlsx";
import { Process } from "../data/types";

export function rowsForExport(processes: Process[]) {
  return processes.map((process) => ({
    Process: process.processName,
    Department: process.department,
    "Current Stage": process.currentStage,
    Status: process.overallStatus,
    Progress: Math.round(process.progress * 100),
    Health: process.health,
    Priority: process.priority,
    "Current Owner": process.currentOwner,
    "Waiting For": process.waitingFor,
    Blocked: process.blocked ? "Yes" : "No",
    "Due Date": process.dueDate,
    "Delay Days": process.varianceDays,
    "Next Action": process.nextAction,
    "Last Updated": process.lastUpdated,
  }));
}

export function downloadCsv(processes: Process[], filename = "filtered-processes.csv") {
  const worksheet = XLSX.utils.json_to_sheet(rowsForExport(processes));
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), filename);
}

export function downloadExcel(processes: Process[], filename = "filtered-processes.xlsx") {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rowsForExport(processes)), "Processes");
  XLSX.writeFile(workbook, filename);
}

export function downloadPdf(processes: Process[], filename = "rpa-report.pdf") {
  const doc = new jsPDF();
  doc.setFontSize(16);
  doc.text("RPA Project Control Center Report", 14, 18);
  doc.setFontSize(10);
  processes.slice(0, 28).forEach((process, index) => {
    doc.text(`${process.processName || "Not provided"} | ${process.overallStatus} | ${process.health}`, 14, 30 + index * 8);
  });
  doc.save(filename);
}

export function printReport() {
  window.print();
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}
