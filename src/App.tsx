import {
  AlertTriangle,
  BarChart3,
  CalendarDays,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  ClipboardList,
  Columns3,
  Database,
  Download,
  FileSpreadsheet,
  Filter,
  Gauge,
  Home,
  KanbanSquare,
  LayoutDashboard,
  List,
  Lock,
  Moon,
  PanelLeftClose,
  PanelLeftOpen,
  Printer,
  RefreshCcw,
  Search,
  Server,
  Settings,
  ShieldAlert,
  Sparkles,
  Sun,
  Table2,
  Upload,
  UserCircle,
  X,
} from "lucide-react";
import { motion } from "framer-motion";
import {
  ColumnDef,
  SortingState,
  VisibilityState,
  flexRender,
  getCoreRowModel,
  getPaginationRowModel,
  getSortedRowModel,
  useReactTable,
} from "@tanstack/react-table";
import { useEffect, useMemo, useRef, useState } from "react";
import { DEFAULT_TRACKER_FILE, DEFAULT_TRACKER_LABEL, importDefaultExcel, importExcel, readWorkbook, validateWorkbook } from "./data/importExcel";
import { clearActiveProject, getActiveProject, saveActiveProject } from "./data/storage";
import { Activity, FilterState, Process, ProjectData, ValidationResult } from "./data/types";
import { applyFilters, defaultFilters, uniqueValues } from "./utils/filters";
import { bottlenecks, buildInsights, calculateMetrics, isDelayed, needsAttention, staleDays } from "./utils/calculations";
import { formatDate } from "./utils/date";
import { downloadCsv, downloadExcel, downloadPdf, printReport } from "./utils/exporters";

type Page =
  | "login"
  | "overview"
  | "processes"
  | "timeline"
  | "kanban"
  | "blockers"
  | "activity"
  | "reports"
  | "upload"
  | "settings"
  | "not-found";

const navItems: { page: Page; label: string; icon: typeof Home }[] = [
  { page: "overview", label: "Overview", icon: LayoutDashboard },
  { page: "processes", label: "Processes", icon: ClipboardList },
  { page: "timeline", label: "Timeline", icon: CalendarDays },
  { page: "kanban", label: "Kanban", icon: KanbanSquare },
  { page: "blockers", label: "Blockers", icon: ShieldAlert },
  { page: "activity", label: "Activity Log", icon: List },
  { page: "reports", label: "Reports", icon: BarChart3 },
  { page: "upload", label: "Upload Excel", icon: Upload },
  { page: "settings", label: "Settings", icon: Settings },
];

function display(value?: string | number | boolean) {
  if (value === undefined || value === null || value === "") return "Not provided";
  if (typeof value === "boolean") return value ? "Yes" : "No";
  return String(value);
}

function processLabel(process: Process) {
  return display(process.processName);
}

function pct(value: number) {
  return `${Math.round(value * 100)}%`;
}

export function App() {
  const [data, setData] = useState<ProjectData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [page, setPage] = useState<Page>("overview");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [theme, setTheme] = useState<"light" | "dark">(() => (localStorage.getItem("rpa-theme") as "light" | "dark") || "light");
  const [filters, setFilters] = useState<FilterState>(() => {
    const saved = localStorage.getItem("rpa-last-filters");
    return saved ? { ...defaultFilters, ...JSON.parse(saved) } : defaultFilters;
  });
  const [viewMode, setViewMode] = useState<"cards" | "list">(() => (sessionStorage.getItem("rpa-view") as "cards" | "list") || "cards");
  const [cardSize, setCardSize] = useState<"comfortable" | "compact">("comfortable");
  const [selectedProcessId, setSelectedProcessId] = useState<string>("");

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem("rpa-theme", theme);
  }, [theme]);

  useEffect(() => {
    localStorage.setItem("rpa-last-filters", JSON.stringify(filters));
  }, [filters]);

  useEffect(() => {
    sessionStorage.setItem("rpa-view", viewMode);
  }, [viewMode]);

  useEffect(() => {
    async function load() {
      try {
        const stored = await getActiveProject();
        if (stored && ![DEFAULT_TRACKER_FILE, DEFAULT_TRACKER_LABEL].includes(stored.uploadedFileName)) {
          setData(stored);
        } else {
          const initial = await importDefaultExcel();
          await saveActiveProject(initial, stored);
          setData(initial);
        }
      } catch (loadError) {
        setError(loadError instanceof Error ? loadError.message : "The project data could not be loaded.");
      } finally {
        setLoading(false);
      }
    }
    load();
  }, []);

  const processes = data?.processes || [];
  const filtered = useMemo(() => applyFilters(processes, filters), [processes, filters]);
  const selectedProcess = processes.find((process) => process.processId === selectedProcessId) || filtered[0] || processes[0];

  function applyChartFilter(key: keyof FilterState, value: string) {
    setFilters({ ...filters, [key]: [value] });
    setPage("processes");
  }

  function openProcess(process: Process) {
    setSelectedProcessId(process.processId);
    setPage("processes");
    window.location.hash = `process/${process.processId}`;
  }

  useEffect(() => {
    const handleHash = () => {
      const hash = window.location.hash.replace("#", "");
      if (hash.startsWith("process/")) {
        setSelectedProcessId(hash.split("/")[1]);
        setPage("processes");
      }
    };
    handleHash();
    window.addEventListener("hashchange", handleHash);
    return () => window.removeEventListener("hashchange", handleHash);
  }, []);

  if (loading) return <LoadingState />;
  if (error) return <ErrorState message={error} />;

  return (
    <div className="app-shell">
      <aside className={`sidebar ${sidebarCollapsed ? "is-collapsed" : ""}`}>
        <div className="brand">
          <div className="brand-mark"><Gauge size={20} /></div>
          {!sidebarCollapsed && <div><strong>RPA Control</strong><span>Project center</span></div>}
        </div>
        <button className="icon-button sidebar-toggle" onClick={() => setSidebarCollapsed(!sidebarCollapsed)} aria-label="Toggle sidebar">
          {sidebarCollapsed ? <PanelLeftOpen size={18} /> : <PanelLeftClose size={18} />}
        </button>
        <nav>
          {navItems.map((item) => (
            <button key={item.page} className={`nav-item ${page === item.page ? "active" : ""}`} onClick={() => setPage(item.page)}>
              <item.icon size={18} />
              {!sidebarCollapsed && <span>{item.label}</span>}
            </button>
          ))}
        </nav>
      </aside>

      <main className="main">
        <Topbar
          data={data}
          theme={theme}
          setTheme={setTheme}
          search={filters.search}
          setSearch={(search) => setFilters({ ...filters, search })}
        />

        {!data ? (
          <EmptyState title="No project data has been uploaded yet." description="Upload an Excel tracker to activate the dashboard." icon={FileSpreadsheet} />
        ) : (
          <>
            {page === "login" && <LoginPlaceholder />}
            {page === "overview" && <OverviewPage data={data} filters={filters} setFilters={setFilters} openProcess={openProcess} applyChartFilter={applyChartFilter} />}
            {page === "processes" && (
              <ProcessesPage
                processes={filtered}
                allProcesses={processes}
                filters={filters}
                setFilters={setFilters}
                viewMode={viewMode}
                setViewMode={setViewMode}
                cardSize={cardSize}
                setCardSize={setCardSize}
                openProcess={openProcess}
                selectedProcess={selectedProcess}
                activities={data.activities}
              />
            )}
            {page === "timeline" && <TimelinePage processes={filtered} openProcess={openProcess} />}
            {page === "kanban" && <KanbanPage processes={filtered} openProcess={openProcess} />}
            {page === "blockers" && <BlockersPage processes={filtered} openProcess={openProcess} />}
            {page === "activity" && <ActivityPage activities={data.activities} processes={processes} />}
            {page === "reports" && <ReportsPage processes={filtered} allProcesses={processes} activities={data.activities} overallProgress={data.overallProgress} />}
            {page === "upload" && <UploadPage current={data} onImported={setData} />}
            {page === "settings" && <SettingsPage data={data} reset={async () => { await clearActiveProject(); window.location.reload(); }} />}
            {page === "not-found" && <NotFound />}
          </>
        )}
      </main>
    </div>
  );
}

function Topbar({ data, theme, setTheme, search, setSearch }: { data: ProjectData | null; theme: "light" | "dark"; setTheme: (theme: "light" | "dark") => void; search: string; setSearch: (value: string) => void }) {
  return (
    <header className="topbar">
      <div>
        <p className="eyebrow">RPA Project Management</p>
        <h1>RPA Project Control Center</h1>
      </div>
      <label className="global-search">
        <Search size={18} />
        <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search processes, owners, blockers, next actions" />
      </label>
      <div className="topbar-meta">
        <span title="Current data file"><FileSpreadsheet size={16} /> {data?.uploadedFileName || "No file"}</span>
        <span>{data ? formatDate(data.uploadedAt, true) : "Not uploaded"}</span>
        <button className="icon-button" onClick={() => setTheme(theme === "light" ? "dark" : "light")} aria-label="Toggle theme">
          {theme === "light" ? <Moon size={18} /> : <Sun size={18} />}
        </button>
        <div className="profile"><UserCircle size={24} /><span>Project Manager</span></div>
      </div>
    </header>
  );
}

function OverviewPage({ data, filters, setFilters, openProcess, applyChartFilter }: { data: ProjectData; filters: FilterState; setFilters: (filters: FilterState) => void; openProcess: (process: Process) => void; applyChartFilter: (key: keyof FilterState, value: string) => void }) {
  const metrics = calculateMetrics(data.processes, data.activities);
  const attention = data.processes.filter(needsAttention);
  const infrastructureActivities = data.activities.filter((activity) => activity.processId === "INFRA");
  const overallProgress = data.overallProgress ?? metrics.completion;
  const bottle = bottlenecks(data.processes);
  const insights = buildInsights(data.processes);
  return (
    <div className="page stack">
      <section className="kpi-grid">
        <Kpi title="Total Processes" value={metrics.total} icon={Database} onClick={() => setFilters({ ...defaultFilters })} />
        <Kpi title="Completed" value={metrics.completed} percent={metrics.total ? metrics.completed / metrics.total : 0} icon={CheckCircle2} onClick={() => applyChartFilter("statuses", "Completed")} />
        <Kpi title="In Progress" value={metrics.inProgress} icon={RefreshCcw} onClick={() => applyChartFilter("statuses", "In Progress")} />
        <Kpi title="Not Started" value={metrics.notStarted} icon={ClipboardList} onClick={() => applyChartFilter("statuses", "Not Started")} />
        <Kpi title="Delayed" value={metrics.delayed} icon={AlertTriangle} tone="red" onClick={() => setFilters({ ...filters, delayed: "yes" })} />
        <Kpi title="Blocked" value={metrics.blocked} icon={Lock} tone="red" onClick={() => setFilters({ ...filters, blocked: "yes" })} />
        <Kpi title="At Risk" value={metrics.atRisk} icon={ShieldAlert} tone="amber" onClick={() => applyChartFilter("health", "Amber")} />
        <Kpi title="Overall Completion" value={pct(overallProgress)} icon={Gauge} />
        <Kpi title="Infra Activities" value={infrastructureActivities.length} icon={Server} />
      </section>

      <section className="split-grid">
        <Panel title="Needs Attention" subtitle="Automatically detected from Excel fields and date rules.">
          {attention.length ? (
            <div className="attention-list">
              {attention.slice(0, 8).map((process) => (
                <button className="attention-item" key={process.processId} onClick={() => openProcess(process)}>
                  <div><strong>{processLabel(process)}</strong><span>{display(process.department)}</span></div>
                  <Badge value={process.priority} type="priority" />
                  <span>{process.currentStage}</span>
                  <span>{display(process.waitingFor)}</span>
                  <strong>{process.varianceDays}d</strong>
                  <span>{display(process.nextAction)}</span>
                </button>
              ))}
            </div>
          ) : (
            <EmptyState title="All active processes are currently within their planned timelines." icon={CheckCircle2} compact />
          )}
        </Panel>
        <Panel title="AI Insights" subtitle="Deterministic insights only; no unsupported causes are invented.">
          <div className="insight-list">
            {insights.map((insight) => (
              <div className={`insight ${insight.type.toLowerCase()}`} key={insight.title}>
                <span>{insight.type}</span>
                <strong>{insight.title}</strong>
                <p>{insight.evidence}</p>
                {!!insight.processIds?.length && <small>{data.processes.filter((process) => insight.processIds?.includes(process.processId)).slice(0, 8).map(processLabel).join(", ")}</small>}
              </div>
            ))}
          </div>
        </Panel>
      </section>

      <Panel title="Project Bottlenecks" subtitle="Calculated from normalized uploaded Excel data.">
        <div className="bottleneck-grid">
          <Fact label="Highest phase volume" value={bottle.topStage ? `${bottle.topStage.name} (${bottle.topStage.value})` : "Insufficient data"} />
          <Fact label="Highest pending party" value={bottle.topWaiting ? `${bottle.topWaiting.name} (${bottle.topWaiting.value})` : "Insufficient data"} />
          <Fact label="Common delay reason" value={bottle.topDelayReason ? `${bottle.topDelayReason.name} (${bottle.topDelayReason.value})` : "Insufficient data"} />
          <Fact label="Oldest active blocker" value={bottle.oldestBlocker ? `${processLabel(bottle.oldestBlocker)} · ${formatDate(bottle.oldestBlocker.lastUpdated, true)}` : "No active blockers"} />
          <Fact label="Longest delay" value={bottle.longestDelay ? `${processLabel(bottle.longestDelay)} · ${bottle.longestDelay.varianceDays} days` : "Insufficient data"} />
          <Fact label="Recently stale" value={bottle.stale.length ? `${bottle.stale.length} processes` : "None"} />
        </div>
      </Panel>
    </div>
  );
}

function ProcessesPage(props: { processes: Process[]; allProcesses: Process[]; filters: FilterState; setFilters: (filters: FilterState) => void; viewMode: "cards" | "list"; setViewMode: (mode: "cards" | "list") => void; cardSize: "comfortable" | "compact"; setCardSize: (size: "comfortable" | "compact") => void; openProcess: (process: Process) => void; selectedProcess?: Process; activities: Activity[] }) {
  return (
    <div className="page stack">
      <Toolbar title="Processes" count={props.processes.length}>
        <Segmented value={props.viewMode} options={[["cards", <Columns3 size={16} />, "Card View"], ["list", <Table2 size={16} />, "List View"]]} onChange={props.setViewMode} />
        {props.viewMode === "cards" && <Segmented value={props.cardSize} options={[["comfortable", null, "Comfortable"], ["compact", null, "Compact"]]} onChange={props.setCardSize} />}
        <button className="secondary-button" onClick={() => downloadCsv(props.processes)}><Download size={16} /> Export CSV</button>
      </Toolbar>
      <FiltersBar filters={props.filters} setFilters={props.setFilters} processes={props.allProcesses} />
      {props.processes.length === 0 ? (
        <EmptyState title="No processes match the selected filters." description="Clear filters or upload a tracker with matching records." icon={Filter} />
      ) : props.viewMode === "cards" ? (
        <div className={`process-grid ${props.cardSize}`}>
          {props.processes.map((process) => <ProcessCard key={process.processId} process={process} onClick={() => props.openProcess(process)} />)}
        </div>
      ) : (
        <ProcessTable processes={props.processes} openProcess={props.openProcess} />
      )}
      {props.selectedProcess && <ProcessDetails process={props.selectedProcess} activities={props.activities} />}
    </div>
  );
}

function ProcessCard({ process, onClick }: { process: Process; onClick: () => void }) {
  const statusSummary = `${processLabel(process)}: ${process.overallStatus}, ${pct(process.progress)}, ${process.health}, due ${formatDate(process.dueDate)}. Next action: ${display(process.nextAction)}.`;
  return (
    <motion.article className={`process-card ${needsAttention(process) ? "needs-attention" : ""}`} whileHover={{ y: -3 }} onClick={onClick}>
      <div className="card-top">
        <div><strong>Process</strong><h3>{processLabel(process)}</h3><p>{display(process.department)}</p></div>
        <Badge value={process.health || "System warning"} type="health" />
      </div>
      <div className="progress-row"><span>{process.currentStage}</span><strong>{pct(process.progress)}</strong></div>
      <div className="progress-track"><span style={{ width: pct(process.progress) }} /></div>
      <div className="badge-row">
        <Badge value={process.overallStatus} type="status" />
        <Badge value={process.priority} type="priority" />
        <Badge value={process.blocked ? "Blocked" : "Not blocked"} type="blocked" />
      </div>
      <dl className="card-meta">
        <dt>Owner</dt><dd>{display(process.currentOwner)}</dd>
        <dt>Waiting For</dt><dd>{display(process.waitingFor)}</dd>
        <dt>Delivery</dt><dd>{process.durationDays} days</dd>
        <dt>Due</dt><dd>{formatDate(process.dueDate)}</dd>
        <dt>Delay</dt><dd>{process.varianceDays} days</dd>
        <dt>Next Action</dt><dd>{display(process.nextAction)}</dd>
        <dt>Last Updated</dt><dd>{formatDate(process.lastUpdated, true)}</dd>
      </dl>
      <div className="quick-actions">
        <button onClick={(event) => { event.stopPropagation(); onClick(); }}>View Details</button>
        <button onClick={(event) => { event.stopPropagation(); document.querySelector("#timeline")?.scrollIntoView(); }}>View Timeline</button>
        <button onClick={(event) => { event.stopPropagation(); navigator.clipboard.writeText(statusSummary); }}>Copy Status Summary</button>
      </div>
    </motion.article>
  );
}

function ProcessTable({ processes, openProcess }: { processes: Process[]; openProcess: (process: Process) => void }) {
  const [sorting, setSorting] = useState<SortingState>([]);
  const [columnVisibility, setColumnVisibility] = useState<VisibilityState>({});
  const columns = useMemo<ColumnDef<Process>[]>(() => [
    { accessorKey: "processName", header: "Process", cell: (info) => display(info.getValue<string>()) },
    { accessorKey: "department", header: "Department", cell: (info) => display(info.getValue<string>()) },
    { accessorKey: "currentStage", header: "Current Stage" },
    { accessorKey: "overallStatus", header: "Status", cell: (info) => <Badge value={info.getValue<string>()} type="status" /> },
    { accessorKey: "progress", header: "Progress", cell: (info) => pct(info.getValue<number>()) },
    { accessorKey: "health", header: "Health", cell: (info) => <Badge value={info.getValue<string>()} type="health" /> },
    { accessorKey: "priority", header: "Priority", cell: (info) => <Badge value={info.getValue<string>()} type="priority" /> },
    { accessorKey: "currentOwner", header: "Current Owner", cell: (info) => display(info.getValue<string>()) },
    { accessorKey: "waitingFor", header: "Waiting For", cell: (info) => display(info.getValue<string>()) },
    { accessorKey: "blocked", header: "Blocked", cell: (info) => display(info.getValue<boolean>()) },
    { accessorKey: "durationDays", header: "Delivery Days" },
    { accessorKey: "dueDate", header: "Due Date", cell: (info) => formatDate(info.getValue<string>()) },
    { accessorKey: "varianceDays", header: "Delay Days" },
    { accessorKey: "nextAction", header: "Next Action", cell: (info) => display(info.getValue<string>()) },
    { accessorKey: "lastUpdated", header: "Last Updated", cell: (info) => formatDate(info.getValue<string>(), true) },
    { accessorKey: "businessOwner", header: "Business Owner", cell: (info) => display(info.getValue<string>()) },
    { accessorKey: "delayReason", header: "Delay Reason", cell: (info) => display(info.getValue<string>()) },
    { accessorKey: "blockerDescription", header: "Blocker", cell: (info) => display(info.getValue<string>()) },
  ], []);
  const table = useReactTable({ data: processes, columns, state: { sorting, columnVisibility }, onSortingChange: setSorting, onColumnVisibilityChange: setColumnVisibility, getCoreRowModel: getCoreRowModel(), getSortedRowModel: getSortedRowModel(), getPaginationRowModel: getPaginationRowModel(), initialState: { pagination: { pageSize: 12 } } });
  return (
    <Panel title="List View" subtitle="Sortable, filterable, paginated, and export-ready.">
      <details className="column-menu">
        <summary>Column visibility</summary>
        <div>{table.getAllLeafColumns().map((column) => <label key={column.id}><input type="checkbox" checked={column.getIsVisible()} onChange={column.getToggleVisibilityHandler()} /> {column.id}</label>)}</div>
      </details>
      <div className="table-wrap">
        <table className="data-table">
          <thead>{table.getHeaderGroups().map((group) => <tr key={group.id}>{group.headers.map((header) => <th key={header.id} onClick={header.column.getToggleSortingHandler()}>{flexRender(header.column.columnDef.header, header.getContext())}<span className="resize-handle" /></th>)}</tr>)}</thead>
          <tbody>{table.getRowModel().rows.map((row) => <tr key={row.id} onClick={() => openProcess(row.original)}>{row.getVisibleCells().map((cell) => <td key={cell.id}>{flexRender(cell.column.columnDef.cell, cell.getContext())}</td>)}</tr>)}</tbody>
        </table>
      </div>
      <div className="pagination">
        <button className="icon-button" onClick={() => table.previousPage()} disabled={!table.getCanPreviousPage()}><ChevronLeft size={16} /></button>
        <span>Page {table.getState().pagination.pageIndex + 1} of {table.getPageCount()}</span>
        <button className="icon-button" onClick={() => table.nextPage()} disabled={!table.getCanNextPage()}><ChevronRight size={16} /></button>
      </div>
    </Panel>
  );
}

function ProcessDetails({ process, activities }: { process: Process; activities: Activity[] }) {
  const relatedActivities = activities.filter((activity) => activity.processId === process.processId).sort((a, b) => new Date(b.updateDate || 0).getTime() - new Date(a.updateDate || 0).getTime());
  return (
    <Panel title={processLabel(process)} subtitle="Process detail page">
      <div className="details-header">
        <Badge value={process.priority} type="priority" /><Badge value={process.overallStatus} type="status" /><Badge value={process.health || "System warning"} type="health" /><Badge value={process.blocked ? "Blocked" : "Not blocked"} type="blocked" />
        <strong>{pct(process.progress)}</strong><span>{process.currentStage}</span><span>{display(process.currentOwner)}</span><span>{display(process.waitingFor)}</span><span>{process.durationDays} delivery days</span><span>{formatDate(process.dueDate)}</span><span>{process.varianceDays} delay days</span>
      </div>
      <div className="detail-grid">
        <DetailSection title="Executive Summary" items={{ "Current status": process.overallStatus, "Current stage": process.currentStage, "Current owner": process.currentOwner, "Waiting for": process.waitingFor, "Reason for delay": process.delayReason, Blocker: process.blockerDescription, "Next action": process.nextAction, "Latest update": formatDate(process.lastUpdated, true), "Business owner": process.businessOwner }} />
        <DetailSection title="Dates and Performance" items={{ "Planned Start": formatDate(process.plannedStart), "Planned Finish": formatDate(process.plannedFinish), "Delivery Days": `${process.durationDays} days`, "Actual Start": formatDate(process.actualStart), "Actual Finish": formatDate(process.actualFinish), Variance: `${process.varianceDays} days`, Completion: pct(process.progress) }} />
        <DetailSection title="Ownership" items={{ "Business Owner": process.businessOwner, "Current Owner": process.currentOwner, Responsibility: process.responsibility, "Waiting For": process.waitingFor, Department: process.department, "Updated By": process.updatedBy }} />
        <DetailSection title="Blockers and Risks" items={{ Blocked: process.blocked ? "Yes" : "No", "Blocker description": process.blockerDescription, "Delay reason": process.delayReason, Health: process.health, Priority: process.priority, "Age of blocker": process.blocked ? `${staleDays(process) ?? "Not available"} days` : "Not provided" }} />
      </div>
      <section className="next-action"><strong>Next Action</strong><p>{display(process.nextAction)}</p></section>
      <section id="timeline" className="phase-timeline">
        {process.phases.map((phase) => <PhaseStep key={phase.phaseName} phase={phase} />)}
      </section>
      <section>
        <h3>Activity Log</h3>
        {relatedActivities.length ? relatedActivities.map((activity) => <ActivityItem key={activity.activityId} activity={activity} />) : <EmptyState title="No activity history is available in the uploaded workbook." compact icon={List} />}
      </section>
    </Panel>
  );
}

function TimelinePage({ processes, openProcess }: { processes: Process[]; openProcess: (process: Process) => void }) {
  const [mode, setMode] = useState<"process" | "phase">("process");
  const [scale, setScale] = useState<"week" | "month" | "quarter">("month");
  const allDates = processes.flatMap((process) => [process.plannedStart, process.plannedFinish]).filter(Boolean) as string[];
  const min = allDates.length ? Math.min(...allDates.map((date) => new Date(date).getTime())) : Date.now();
  const max = allDates.length ? Math.max(...allDates.map((date) => new Date(date).getTime())) : Date.now() + 1;
  const span = Math.max(max - min, 1);
  return (
    <div className="page stack">
      <Toolbar title="Project Timeline" count={processes.length}>
        <Segmented value={mode} options={[["process", null, "Process Timeline"], ["phase", null, "Phase Timeline"]]} onChange={setMode} />
        <Segmented value={scale} options={[["week", null, "Week"], ["month", null, "Month"], ["quarter", null, "Quarter"]]} onChange={setScale} />
      </Toolbar>
      <Panel title={`${mode === "process" ? "Process" : "Phase"} Timeline`} subtitle={`Viewing by ${scale}. Today is marked visually when it falls within the range.`}>
        <div className="gantt">
          {processes.map((process) => {
            const rows = mode === "process" ? [{ name: processLabel(process), start: process.plannedStart, finish: process.plannedFinish, health: process.health, blocked: process.blocked, status: process.overallStatus }] : process.phases.map((phase) => ({ name: `${processLabel(process)} · ${phase.phaseName}`, start: phase.plannedStart, finish: phase.plannedFinish, health: phase.health, blocked: phase.blocked, status: phase.status }));
            return rows.map((row) => {
              const left = row.start ? ((new Date(row.start).getTime() - min) / span) * 100 : 0;
              const width = row.finish && row.start ? Math.max(((new Date(row.finish).getTime() - new Date(row.start).getTime()) / span) * 100, 2) : 2;
              return (
                <button key={`${process.processId}-${row.name}`} className="gantt-row" onClick={() => openProcess(process)}>
                  <span>{row.name}</span>
                  <div className="gantt-track"><i className={`gantt-bar ${row.blocked ? "blocked" : row.health?.toLowerCase()}`} style={{ left: `${left}%`, width: `${width}%` }}>{row.status}</i></div>
                </button>
              );
            });
          })}
        </div>
      </Panel>
    </div>
  );
}

function KanbanPage({ processes, openProcess }: { processes: Process[]; openProcess: (process: Process) => void }) {
  const stages = ["Assessment", "PDD Share", "PDD Approval", "Development", "UAT", "Go Live", "Completed"];
  return (
    <div className="page stack">
      <Toolbar title="Kanban Board" count={processes.length}><span className="muted">Drag-and-drop is disabled in v1 because Excel remains the official source.</span></Toolbar>
      <div className="kanban">
        {stages.map((stage) => (
          <section className="kanban-col" key={stage}>
            <h3>{stage}<span>{processes.filter((p) => p.currentStage === stage || (stage === "Completed" && p.overallStatus === "Completed")).length}</span></h3>
            {processes.filter((p) => p.currentStage === stage || (stage === "Completed" && p.overallStatus === "Completed")).map((process) => (
              <button className="kanban-card" key={process.processId} onClick={() => openProcess(process)}>
                <strong>{processLabel(process)}</strong><Badge value={process.health} type="health" /><small>{display(process.currentOwner)} · {formatDate(process.dueDate)}</small>{process.blocked && <em><Lock size={13} /> Blocked</em>}
              </button>
            ))}
          </section>
        ))}
      </div>
    </div>
  );
}

function BlockersPage({ processes, openProcess }: { processes: Process[]; openProcess: (process: Process) => void }) {
  const blocked = processes.flatMap((process) => {
    const phaseBlockers = process.phases.filter((phase) => phase.blocked).map((phase) => ({ process, phase: phase.phaseName, description: phase.blockerDescription, waitingFor: phase.waitingFor, dueDate: phase.plannedFinish, delay: phase.varianceDays, owner: phase.currentOwner, priority: process.priority, nextAction: phase.nextAction, lastUpdated: phase.lastUpdated }));
    return process.blocked && phaseBlockers.length === 0 ? [{ process, phase: process.currentStage, description: process.blockerDescription, waitingFor: process.waitingFor, dueDate: process.dueDate, delay: process.varianceDays, owner: process.currentOwner, priority: process.priority, nextAction: process.nextAction, lastUpdated: process.lastUpdated }] : phaseBlockers;
  });
  if (!blocked.length) return <div className="page"><EmptyState title="No active blockers were found." description="Blocked = Yes was not found in the active uploaded Excel data." icon={CheckCircle2} /></div>;
  return <div className="page stack"><Toolbar title="Blockers" count={blocked.length} /><div className="blocker-grid">{blocked.map((item) => <button className="blocker-card" key={`${item.process.processId}-${item.phase}`} onClick={() => openProcess(item.process)}><Badge value={item.priority} type="priority" /><h3>{processLabel(item.process)}</h3><p>{display(item.description)}</p><dl><dt>Phase</dt><dd>{item.phase}</dd><dt>Owner</dt><dd>{display(item.owner)}</dd><dt>Waiting For</dt><dd>{display(item.waitingFor)}</dd><dt>Due</dt><dd>{formatDate(item.dueDate)}</dd><dt>Delay</dt><dd>{item.delay} days</dd><dt>Last Updated</dt><dd>{formatDate(item.lastUpdated, true)}</dd></dl><strong>{display(item.nextAction)}</strong></button>)}</div></div>;
}

function ActivityPage({ activities, processes }: { activities: Activity[]; processes: Process[] }) {
  const [tableView, setTableView] = useState<"timeline" | "table">("timeline");
  const [query, setQuery] = useState("");
  const filtered = activities.filter((activity) => JSON.stringify(activity).toLowerCase().includes(query.toLowerCase()));
  return (
    <div className="page stack">
      <Toolbar title="Activity Log" count={filtered.length}><Segmented value={tableView} options={[["timeline", null, "Timeline"], ["table", null, "Table"]]} onChange={setTableView} /><label className="mini-search"><Search size={15} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Filter activity" /></label></Toolbar>
      {!activities.length ? <EmptyState title="No activity history is currently available. Activity records can be added to the Excel Activity Log sheet and uploaded again." icon={List} /> : tableView === "timeline" ? <Panel title="Timeline View">{filtered.map((activity) => <ActivityItem key={activity.activityId} activity={activity} />)}</Panel> : <Panel title="Table View"><div className="table-wrap"><table className="data-table"><thead><tr>{["Update Date", "Process", "Phase", "Updated By", "Previous", "New", "Next Action", "Waiting For", "Blocker", "Due Date"].map((header) => <th key={header}>{header}</th>)}</tr></thead><tbody>{filtered.map((activity) => <tr key={activity.activityId}><td>{formatDate(activity.updateDate, true)}</td><td>{activity.processName || processes.find((p) => p.processId === activity.processId)?.processName || "Not provided"}</td><td>{activity.phase}</td><td>{activity.updatedBy}</td><td>{activity.previousStatus}</td><td>{activity.newStatus}</td><td>{activity.nextAction}</td><td>{activity.waitingFor}</td><td>{activity.blocker}</td><td>{formatDate(activity.dueDate)}</td></tr>)}</tbody></table></div></Panel>}
    </div>
  );
}

function ReportsPage({ processes, allProcesses, activities, overallProgress }: { processes: Process[]; allProcesses: Process[]; activities: Activity[]; overallProgress?: number }) {
  const metrics = calculateMetrics(allProcesses, activities);
  const projectProgress = overallProgress ?? metrics.completion;
  const bottle = bottlenecks(allProcesses);
  const delayed = allProcesses.filter(isDelayed);
  const blocked = allProcesses.filter((process) => process.blocked);
  return (
    <div className="page stack report-page">
      <Toolbar title="Reports" count={processes.length}>
        <button className="secondary-button" onClick={() => downloadPdf(processes)}><Download size={16} /> PDF</button>
        <button className="secondary-button" onClick={() => downloadExcel(processes)}><FileSpreadsheet size={16} /> Excel</button>
        <button className="secondary-button" onClick={() => downloadCsv(processes)}><Download size={16} /> CSV</button>
        <button className="secondary-button" onClick={printReport}><Printer size={16} /> Print</button>
      </Toolbar>
      {["Executive Summary", "Weekly Status Report", "Delayed Processes Report", "Blockers Report", "Owner Workload Report", "Stage Bottleneck Report", "Process Health Report", "Department Progress Report"].map((title) => (
        <Panel key={title} title={title}>
          {title === "Executive Summary" ? <p>Total number of processes: {metrics.total}. Overall project completion from the uploaded workbook: {pct(projectProgress)}. Completed processes: {metrics.completed}. In progress: {metrics.inProgress}. Delayed: {metrics.delayed}. Blocked: {metrics.blocked}. Main bottleneck: {bottle.topStage ? bottle.topStage.name : "insufficient data"}. Most common waiting party: {bottle.topWaiting ? bottle.topWaiting.name : "insufficient data"}. Highest-risk processes: {allProcesses.filter(needsAttention).slice(0, 5).map(processLabel).join(", ") || "not available"}.</p> : title.includes("Delayed") ? <ProcessMiniList processes={delayed} /> : title.includes("Blockers") ? <ProcessMiniList processes={blocked} /> : <p>This report is generated only from the uploaded Excel data. Where source fields are blank, the report displays Not provided rather than inventing information.</p>}
        </Panel>
      ))}
    </div>
  );
}

function UploadPage({ current, onImported }: { current: ProjectData; onImported: (data: ProjectData) => void }) {
  const [dragging, setDragging] = useState(false);
  const [validation, setValidation] = useState<ValidationResult | null>(null);
  const [pendingData, setPendingData] = useState<ProjectData | null>(null);
  const [message, setMessage] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);
  async function handleFile(file?: File) {
    if (!file) return;
    setMessage("");
    setPendingData(null);
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      setValidation({ ok: false, errors: [{ code: "unsupported-file", message: "Only .xlsx and .xls files are accepted." }], warnings: [], processCount: 0, phaseCount: 0, activityCount: 0, fileName: file.name, uploadedAt: new Date().toISOString() });
      return;
    }
    try {
      const workbook = await readWorkbook(file);
      const preview = validateWorkbook(workbook, file.name);
      setValidation(preview);
      if (!preview.ok) return;
      const result = await importExcel(file);
      setPendingData(result.data || null);
    } catch (uploadError) {
      setValidation({ ok: false, errors: [{ code: "import-failure", message: uploadError instanceof Error ? uploadError.message : "The Excel file could not be imported." }], warnings: [], processCount: 0, phaseCount: 0, activityCount: 0, fileName: file.name, uploadedAt: new Date().toISOString() });
    }
  }
  async function confirm() {
    if (!pendingData) return;
    await saveActiveProject(pendingData, current);
    onImported(pendingData);
    setMessage("Upload confirmed. The active dashboard has been refreshed from the new Excel file.");
    setPendingData(null);
  }
  return (
    <div className="page stack">
      <Toolbar title="Upload Excel" />
      <Panel title="Replace Active Data Source" subtitle="Step 1: select a workbook. Step 2: validate. Step 3: confirm replacement.">
        <div className={`dropzone ${dragging ? "dragging" : ""}`} onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(event) => { event.preventDefault(); setDragging(false); handleFile(event.dataTransfer.files[0]); }}>
          <Upload size={36} /><h3>Select or drag the latest Excel tracker</h3><p>Accepted formats: .xlsx and .xls. A failed validation will not replace the active data.</p>
          <button className="primary-button" onClick={() => inputRef.current?.click()}>Choose File</button>
          <input ref={inputRef} hidden type="file" accept=".xlsx,.xls" onChange={(event) => handleFile(event.target.files?.[0])} />
        </div>
      </Panel>
      {validation && <Panel title="Validation Preview"><ValidationPreview validation={validation} />{pendingData && <div className="confirm-box"><p>Uploading this file will replace the currently displayed project data. The previous data will no longer be active.</p><button className="primary-button" onClick={confirm}>Confirm Replacement</button></div>}</Panel>}
      {message && <div className="success-banner">{message}</div>}
    </div>
  );
}

function SettingsPage({ data, reset }: { data: ProjectData; reset: () => void }) {
  return (
    <div className="page stack">
      <Toolbar title="Settings" />
      <Panel title="Data Architecture">
        <div className="settings-grid">
          <Fact label="Implemented storage" value="Browser IndexedDB for normalized JSON, with a default Excel bootstrap file." />
          <Fact label="No database" value="The first version does not require a database or server write access." />
          <Fact label="Active source" value={data.uploadedFileName} />
          <Fact label="Uploaded at" value={formatDate(data.uploadedAt, true)} />
          <Fact label="Future database path" value="Replace ExcelDataSource with DatabaseDataSource behind the same DataSource interface." />
          <Fact label="Unmapped fields" value="Blank Excel fields remain Not provided; no missing owners, departments, blockers, or causes are invented." />
        </div>
        <button className="danger-button" onClick={reset}>Reset local active data</button>
      </Panel>
    </div>
  );
}

function FiltersBar({ filters, setFilters, processes }: { filters: FilterState; setFilters: (filters: FilterState) => void; processes: Process[] }) {
  const options = {
    statuses: uniqueValues(processes, (p) => p.overallStatus),
    stages: uniqueValues(processes, (p) => p.currentStage),
    health: uniqueValues(processes, (p) => p.health),
    priorities: uniqueValues(processes, (p) => p.priority),
    departments: uniqueValues(processes, (p) => p.department),
    owners: uniqueValues(processes, (p) => p.currentOwner),
    waitingFor: uniqueValues(processes, (p) => p.waitingFor),
  };
  return (
    <Panel title="Filters" className="filters-panel">
      <div className="filters-grid">
        {Object.entries(options).map(([key, values]) => <MultiSelect key={key} label={key} values={values} selected={filters[key as keyof typeof options]} onChange={(selected) => setFilters({ ...filters, [key]: selected })} />)}
        <label><span>Blocked</span><select value={filters.blocked} onChange={(e) => setFilters({ ...filters, blocked: e.target.value as FilterState["blocked"] })}><option value="all">All</option><option value="yes">Yes</option><option value="no">No</option></select></label>
        <label><span>Delayed</span><select value={filters.delayed} onChange={(e) => setFilters({ ...filters, delayed: e.target.value as FilterState["delayed"] })}><option value="all">All</option><option value="yes">Yes</option><option value="no">No</option></select></label>
        <label><span>Due From</span><input type="date" value={filters.dueFrom} onChange={(e) => setFilters({ ...filters, dueFrom: e.target.value })} /></label>
        <label><span>Due To</span><input type="date" value={filters.dueTo} onChange={(e) => setFilters({ ...filters, dueTo: e.target.value })} /></label>
      </div>
      <div className="filter-actions"><button className="secondary-button" onClick={() => setFilters(defaultFilters)}><X size={15} /> Clear All Filters</button><button className="secondary-button" onClick={() => localStorage.setItem(`rpa-saved-filter-${Date.now()}`, JSON.stringify(filters))}>Save Filter View</button><span>Recently used filters are stored locally in this browser.</span></div>
    </Panel>
  );
}

function MultiSelect({ label, values, selected, onChange }: { label: string; values: string[]; selected: string[]; onChange: (values: string[]) => void }) {
  return <label><span>{label}</span><select value={selected[0] || ""} onChange={(e) => onChange(e.target.value ? [e.target.value] : [])}><option value="">All</option>{values.map((value) => <option key={value} value={value}>{value}</option>)}</select></label>;
}

function Kpi({ title, value, percent, icon: Icon, tone = "blue", onClick }: { title: string; value: string | number; percent?: number; icon: typeof Home; tone?: string; onClick?: () => void }) {
  return <button className={`kpi ${tone}`} onClick={onClick}><Icon size={20} /><span>{title}</span><strong>{value}</strong>{percent !== undefined && <small>{pct(percent)}</small>}</button>;
}

function Badge({ value, type }: { value: string; type: "status" | "health" | "priority" | "blocked" }) {
  return <span className={`badge ${type} ${String(value || "na").toLowerCase().replace(/\s+/g, "-")}`}>{display(value)}</span>;
}

function Panel({ title, subtitle, children, className = "" }: { title: string; subtitle?: string; children: React.ReactNode; className?: string }) {
  return <section className={`panel ${className}`}><div className="panel-head"><div><h2>{title}</h2>{subtitle && <p>{subtitle}</p>}</div></div>{children}</section>;
}

function Toolbar({ title, count, children }: { title: string; count?: number; children?: React.ReactNode }) {
  return <div className="toolbar"><div><p className="eyebrow">Workspace</p><h2>{title}{count !== undefined && <span>{count}</span>}</h2></div><div className="toolbar-actions">{children}</div></div>;
}

function Segmented<T extends string>({ value, options, onChange }: { value: T; options: [T, React.ReactNode, string][]; onChange: (value: T) => void }) {
  return <div className="segmented">{options.map(([key, icon, label]) => <button key={key} className={value === key ? "active" : ""} onClick={() => onChange(key)}>{icon}{label}</button>)}</div>;
}

function Fact({ label, value }: { label: string; value: string }) {
  return <div className="fact"><span>{label}</span><strong>{value}</strong></div>;
}

function DetailSection({ title, items }: { title: string; items: Record<string, string | number | boolean | undefined> }) {
  return <section className="detail-section"><h3>{title}</h3>{Object.entries(items).map(([key, value]) => <div key={key}><span>{key}</span><strong>{display(value)}</strong></div>)}</section>;
}

function PhaseStep({ phase }: { phase: Process["phases"][number] }) {
  const Icon = phase.status === "Completed" ? CheckCircle2 : phase.blocked ? Lock : phase.health === "Red" ? AlertTriangle : phase.status === "In Progress" ? RefreshCcw : Gauge;
  return <article className="phase-step"><Icon size={18} /><div><h4>{phase.phaseName}</h4><p>{phase.status} · {phase.responsibility}</p><dl><dt>Planned</dt><dd>{formatDate(phase.plannedStart)} - {formatDate(phase.plannedFinish)}</dd><dt>Days</dt><dd>{phase.durationDays} days</dd><dt>Actual</dt><dd>{formatDate(phase.actualStart)} - {formatDate(phase.actualFinish)}</dd><dt>Owner</dt><dd>{display(phase.currentOwner)}</dd><dt>Waiting</dt><dd>{display(phase.waitingFor)}</dd><dt>Variance</dt><dd>{phase.varianceDays} days</dd><dt>Blocked</dt><dd>{phase.blocked ? "Yes" : "No"}</dd><dt>Blocker</dt><dd>{display(phase.blockerDescription)}</dd><dt>Delay</dt><dd>{display(phase.delayReason)}</dd><dt>Next</dt><dd>{display(phase.nextAction)}</dd><dt>Updated</dt><dd>{formatDate(phase.lastUpdated, true)}</dd></dl></div></article>;
}

function ActivityItem({ activity }: { activity: Activity }) {
  return <article className="activity-item"><span>{formatDate(activity.updateDate, true)}</span><strong>{display(activity.processName)}</strong><p>{activity.updateDescription || "Not provided"}</p><small>{display(activity.updatedBy)} · {activity.previousStatus || "Not provided"} → {activity.newStatus || "Not provided"} · Waiting for {display(activity.waitingFor)}</small></article>;
}

function ProcessMiniList({ processes }: { processes: Process[] }) {
  return processes.length ? <ul className="mini-list">{processes.slice(0, 10).map((process) => <li key={process.processId}><strong>{processLabel(process)}</strong><Badge value={process.health} type="health" /></li>)}</ul> : <p>Insufficient matching data in the uploaded workbook.</p>;
}

function ValidationPreview({ validation }: { validation: ValidationResult }) {
  return <div className="validation"><div className={`validation-status ${validation.ok ? "ok" : "bad"}`}>{validation.ok ? "Workbook validation passed" : "Workbook validation failed"}</div><div className="validation-grid"><Fact label="File name" value={validation.fileName} /><Fact label="Upload date" value={formatDate(validation.uploadedAt, true)} /><Fact label="Processes" value={String(validation.processCount)} /><Fact label="Phase records" value={String(validation.phaseCount)} /><Fact label="Activity records" value={String(validation.activityCount)} /><Fact label="Warnings" value={String(validation.warnings.length)} /><Fact label="Errors" value={String(validation.errors.length)} /></div>{validation.errors.map((error) => <p className="validation-error" key={error.code + error.message}>{error.message}</p>)}{validation.warnings.map((warning) => <p className="validation-warning" key={warning.code + warning.message}>{warning.message}</p>)}</div>;
}

function EmptyState({ title, description, icon: Icon, compact = false }: { title: string; description?: string; icon: typeof Home; compact?: boolean }) {
  return <div className={`empty-state ${compact ? "compact" : ""}`}><Icon size={compact ? 22 : 34} /><h2>{title}</h2>{description && <p>{description}</p>}</div>;
}

function LoadingState() {
  return <div className="full-state"><Sparkles size={34} /><h1>Loading project control center</h1><p>Reading the active Excel data source.</p></div>;
}

function ErrorState({ message }: { message: string }) {
  return <div className="full-state error"><AlertTriangle size={34} /><h1>The Excel file could not be imported.</h1><p>{message}</p></div>;
}

function LoginPlaceholder() {
  return <div className="page"><Panel title="Login Placeholder" subtitle="Authentication can be connected later without changing the dashboard data model."><button className="primary-button">Continue as Project Manager</button></Panel></div>;
}

function NotFound() {
  return <div className="page"><EmptyState title="404 page" description="The requested application page does not exist." icon={AlertTriangle} /></div>;
}
