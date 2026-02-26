# Honeywell Experion Migration Project Scheduler

A Blazor Server desktop application for configuring Honeywell Experion migration schedules and exporting Microsoft Project-compatible XML files.

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- [WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) (required for the desktop host; pre-installed on Windows 10/11 with Edge)
- Windows 10 or later (for the WinForms host)

## Project Structure

```
MigrationScheduler/
├── MigrationScheduler.sln                   # Solution file
├── MigrationScheduler.Host/                 # WinForms + WebView2 desktop shell
│   ├── Program.cs
│   ├── MainForm.cs
│   └── MigrationScheduler.Host.csproj
├── MigrationScheduler.Blazor/               # Blazor Server application
│   ├── Program.cs
│   ├── App.razor
│   ├── Pages/
│   │   ├── ProjectSetup.razor               # Step 1: Project setup wizard
│   │   ├── ActivityEditor.razor             # Step 2: Task editor grid
│   │   ├── ConstraintManager.razor          # Step 3: Dependencies & constraints
│   │   └── ExportPage.razor                 # Step 4: XML export & Gantt preview
│   ├── Components/
│   │   ├── TaskGrid.razor                   # Reusable task table
│   │   ├── PrerequisiteWarning.razor        # Prerequisite error banner
│   │   ├── GanttPreview.razor               # Read-only HTML Gantt chart
│   │   └── MigrationTypeSelector.razor      # Migration type checkboxes
│   ├── Models/
│   │   ├── ProjectModel.cs                  # Root project model
│   │   ├── TaskModel.cs                     # Task/activity model
│   │   ├── PredecessorLink.cs               # Dependency link model
│   │   ├── TaskConstraint.cs                # Constraint type enum
│   │   └── MigrationType.cs                 # Migration category enum
│   ├── Services/
│   │   ├── ProjectStateService.cs           # Singleton state management
│   │   ├── XmlExportService.cs              # MSP XML generation
│   │   ├── DraftService.cs                  # JSON save/load
│   │   ├── PrerequisiteValidator.cs         # Prerequisite rules engine
│   │   └── TaskTemplateService.cs           # Default task templates
│   ├── wwwroot/
│   │   ├── css/app.css
│   │   └── js/interop.js                   # JS interop for file save/load
│   └── MigrationScheduler.Blazor.csproj
└── README.md
```

## Build & Run

### Option 1: Blazor Server (Browser)

Run the Blazor application directly in a browser — no WebView2 required:

```bash
cd MigrationScheduler
dotnet restore
dotnet build
dotnet run --project MigrationScheduler.Blazor
```

Then open `http://localhost:5199` in your browser.

### Option 2: Desktop Application (WebView2)

Run the WinForms host which embeds the Blazor app in a native window:

```bash
cd MigrationScheduler
dotnet restore
dotnet build
dotnet run --project MigrationScheduler.Host
```

> **Note:** The Host project requires Windows and the WebView2 Runtime.

### Publish as a Self-Contained Executable

```bash
dotnet publish MigrationScheduler.Host -c Release -r win-x64 --self-contained
```

The output will be in `MigrationScheduler.Host/bin/Release/net8.0-windows/win-x64/publish/`.

## Usage Workflow

1. **Project Setup** — Enter project details (name, USP number, dates, PM), select migration types, and validate prerequisites.
2. **Activity Editor** — Review and edit the auto-generated task list. Add, delete, or reorder tasks.
3. **Constraints** — Define predecessor relationships (FS, SS, FF, SF with lag) and task constraints (ASAP, MSO, SNET, etc.).
4. **Export** — Preview the Gantt chart, validate the schedule, and export as MSP-compatible XML.

## Importing into Microsoft Project

1. Open Microsoft Project (2007 or later).
2. Go to **File > Open**.
3. Change the file type filter to **XML Format (*.xml)**.
4. Select the exported `{USP}_{ProjectName}_Schedule.xml` file.
5. Microsoft Project will import all tasks, dependencies, constraints, resources, and assignments.

## Migration Types Supported

| Code | Full Name | Description |
|------|-----------|-------------|
| ELCN | Enhanced Local Control Network | LCN to ELCN / FTE network migration |
| EUCN | Enhanced Universal Control Network | xPM to EHPM migration via NIM to ENIM/ENB |
| AMT | Advanced Migration Technology | ControlEdge PLC / RTU migrations |
| Virtualization | Experion Virtualization | Physical to VMware vSphere migration |
| Safety Manager | Safety Manager / TCMI | Triconex Communication Module Interface upgrade |
| UOC | Universal Operations Controller | Migration to ControlEdge UOC |
| Hardware Refresh | Hardware Refresh | Server/workstation/switch replacement |

## Prerequisite Rules

- **EUCN** requires ELCN
- **AMT** requires EUCN
- **Safety Manager** requires both EUCN and ELCN
- **UOC** requires ELCN
- **Virtualization** requires Hardware Refresh

Each prerequisite can be bypassed by marking it as "Already Completed (Pre-existing)" for brownfield projects.

## Save & Resume

Use the **Save Draft** button to export the current project state as a JSON file. Use **Load Draft** on the Setup page to resume a previous session.

## MSP XML Compatibility

The generated XML conforms to the Microsoft Project XML Schema (`xmlns="http://schemas.microsoft.com/project"`) and is compatible with:

- Microsoft Project 2007, 2010, 2013, 2016, 2019, 2021, and Project for Microsoft 365
- Duration format: ISO 8601 (`PT{hours}H0M0S`)
- Predecessor lag: stored as `days * 4800` (tenths of minutes)
- Standard calendar: 5-day work week, Mon-Fri, 08:00-17:00
