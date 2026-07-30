# RPA Project Control Center

Production-quality local web application for reading the `RPA_Project_Tracker_Web_Ready.xlsx` workbook and turning it into an executive RPA project control system.

## Implemented Stack

- React + TypeScript + Vite
- SheetJS (`xlsx`) for Excel parsing
- IndexedDB for active normalized JSON storage, without a database
- Recharts for dashboard charts
- TanStack Table for process list sorting, pagination, and column visibility
- Lucide icons
- Framer Motion micro-interactions

## Run Locally

This workspace uses the bundled Codex Node runtime:

```bash
PATH="/Users/macbookpro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin:$PATH" \
/Users/macbookpro/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/pnpm dev
```

Then open the local URL shown by Vite.

Build for production:

```bash
PATH="/Users/macbookpro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin:$PATH" \
/Users/macbookpro/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/pnpm build
```

## Excel Data Source

The default workbook is served from:

`public/data/RPA_Project_Tracker_Web_Ready.xlsx`

The application does not hardcode process records. It parses the workbook at startup, normalizes it into application objects, and calculates dashboards from that normalized data.

## Upload Workflow

1. Go to **Upload Excel**.
2. Drag and drop or select a `.xlsx` or `.xls` workbook.
3. The app validates required sheets and columns.
4. A validation preview shows process, phase, activity, warning, and error counts.
5. Confirm replacement.
6. The previous normalized active data is saved as a local backup in IndexedDB, and the new valid normalized JSON becomes active.

Invalid uploads do not replace the active data.

## Excel Column Mapping

`Project Tracker` is normalized into:

- Process summary rows where `Process / Phase = Process Summary`
- Phase rows where `Process / Phase` is one of `Assessment`, `PDD Share`, `PDD Approval`, `Development`, `UAT`, or `Go Live`

Key mappings:

- `Process ID` → `process.processId`
- `Process` → `process.processName`
- `Process / Phase` → process summary marker or `phase.phaseName`
- `Status` → status fields
- `Progress` → completion percentage
- `Start` and `Finish` → planned dates
- `Actual Start` and `Actual Finish` → actual dates
- `Variance Days` → delay days
- `Health` → health badge
- `Responsibility` → team or role
- `Next Action` → next action
- `Current Owner`, `Waiting For`, `Blocked`, `Blocker Description`, `Delay Reason`, `Priority`, `Department`, `Business Owner`, `Last Updated`, `Updated By` → operational control fields

`Activity Log` maps to activity timeline records and ignores the `EXAMPLE-DELETE` row.

## Data Storage Without a Database

This first version stores normalized JSON in browser IndexedDB:

- Active project: latest valid uploaded workbook
- Backup project: previous active normalized JSON when replacing data

Raw uploaded files are not exposed publicly by the app. The default workbook is included only as a starter local data source.

## Future Database Readiness

The UI is designed around a `DataSource` interface:

- `getProcesses()`
- `getProcessById()`
- `getActivities()`
- `importExcel()`
- `getDashboardMetrics()`

The current implementation is `ExcelDataSource`. A future `DatabaseDataSource` can replace it without rebuilding the UI.

## Included Pages

- Login placeholder
- Executive Dashboard
- Processes card view
- Processes list view
- Process details
- Project timeline
- Kanban board
- Blockers
- Activity log
- Reports
- Excel upload
- Settings
- Empty, loading, error, and 404 states

## Unmapped or Blank Fields

Fields blank in Excel are displayed as `Not provided` or left visually empty where appropriate. The app does not invent missing current owners, departments, business owners, waiting parties, blocker causes, delay reasons, or historical trends.

## Preview

Run the app locally and open the Vite URL. The landing page is the executive dashboard.
