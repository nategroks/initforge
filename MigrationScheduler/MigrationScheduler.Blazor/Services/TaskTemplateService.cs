using MigrationScheduler.Blazor.Models;

namespace MigrationScheduler.Blazor.Services;

/// <summary>
/// Provides default task templates for each migration type.
/// </summary>
public class TaskTemplateService
{
    /// <summary>
    /// Generates the full default task list based on selected migration types.
    /// General tasks are always included.
    /// </summary>
    public List<TaskModel> GenerateDefaultTasks(
        HashSet<MigrationType> selectedTypes,
        string projectUspNumber,
        DateTime? projectStartDate)
    {
        var tasks = new List<TaskModel>();
        int taskId = 1;
        int sectionIndex = 1;

        // General tasks are always included
        AddSection(tasks, ref taskId, ref sectionIndex, MigrationType.General, projectUspNumber, GetGeneralTasks());

        // Add tasks for each selected type in a deterministic order
        MigrationType[] orderedTypes =
        [
            MigrationType.ELCN,
            MigrationType.EUCN,
            MigrationType.AMT,
            MigrationType.Virtualization,
            MigrationType.SafetyManager,
            MigrationType.UOC,
            MigrationType.HardwareRefresh
        ];

        foreach (var type in orderedTypes)
        {
            if (selectedTypes.Contains(type))
            {
                var templateTasks = GetTasksForType(type);
                AddSection(tasks, ref taskId, ref sectionIndex, type, projectUspNumber, templateTasks);
            }
        }

        return tasks;
    }

    private static void AddSection(
        List<TaskModel> tasks,
        ref int taskId,
        ref int sectionIndex,
        MigrationType type,
        string uspNumber,
        List<string> taskNames)
    {
        // Add summary task for the section
        tasks.Add(new TaskModel
        {
            TaskId = taskId++,
            WbsCode = $"{sectionIndex}",
            TaskName = type == MigrationType.General ? "General Project Tasks" : $"{type.GetShortName()} Migration Tasks",
            MigrationTypeTag = type,
            UspNumber = uspNumber,
            DurationDays = 0,
            IsSummaryTask = true
        });

        int childIndex = 1;
        foreach (var name in taskNames)
        {
            tasks.Add(new TaskModel
            {
                TaskId = taskId++,
                WbsCode = $"{sectionIndex}.{childIndex}",
                TaskName = name,
                MigrationTypeTag = type,
                UspNumber = uspNumber,
                DurationDays = 5 // Default 5-day duration
            });
            childIndex++;
        }

        sectionIndex++;
    }

    private static List<string> GetGeneralTasks() =>
    [
        "Project Kick-Off / Mobilization",
        "Site Readiness Assessment (SB-MIG30001 Checklist)",
        "Scope Definition Form (SB-MIG30002) Submission to COE",
        "Migration COE Review & Approval",
        "Engineering Design Review",
        "Factory Acceptance Test (FAT)",
        "Site Acceptance Test (SAT)",
        "Project Close-Out & Documentation Handover"
    ];

    private static List<string> GetTasksForType(MigrationType type) => type switch
    {
        MigrationType.ELCN => GetElcnTasks(),
        MigrationType.EUCN => GetEucnTasks(),
        MigrationType.AMT => GetAmtTasks(),
        MigrationType.Virtualization => GetVirtualizationTasks(),
        MigrationType.SafetyManager => GetSafetyManagerTasks(),
        MigrationType.UOC => GetUocTasks(),
        MigrationType.HardwareRefresh => GetHardwareRefreshTasks(),
        _ => []
    };

    private static List<string> GetElcnTasks() =>
    [
        "LCN Network Audit & Documentation",
        "FTE Network Design & Cable Schedule",
        "ENIM / ENB Hardware Procurement",
        "FTE Switch Installation & Configuration",
        "ELCN Node Migration (per LCN)",
        "ELCN Commissioning & Verification",
        "LCN Decommission"
    ];

    private static List<string> GetEucnTasks() =>
    [
        "UCN P2P Data Collection (24hr trending per UCN)",
        "P2P Analysis Submission to COE via EMA",
        "P2P Report Review & Approval",
        "HPM/APM/PM Resource Utilization Snapshot",
        "LVRLOG / Audit File Collection",
        "EHPM Hardware Procurement",
        "NIM → ENIM/ENB Migration (per UCN)",
        "xPM → EHPM Migration (per UCN, per HPM)",
        "UCN Commissioning & Loop Check",
        "NIM Decommission"
    ];

    private static List<string> GetAmtTasks() =>
    [
        "ControlEdge PLC/RTU Hardware Procurement",
        "Existing LM / LEPIU Cabinet Survey",
        "IO Type & Signal Survey",
        "ControlEdge PLC Rack Installation",
        "Field Wiring Reconnection (if applicable)",
        "Ladder Logic / Control Strategy Migration",
        "Off-Process Commissioning & Test",
        "Cutover Execution (Shutdown window)",
        "LM / LEPIU Decommission"
    ];

    private static List<string> GetVirtualizationTasks() =>
    [
        "VMware vSphere Infrastructure Assessment",
        "Virtual Machine Sizing & Design",
        "ESXi Host Hardware Procurement",
        "VMware Installation & Configuration",
        "Experion Server VM Build & Configuration",
        "VM Migration / P2V Conversion",
        "Virtual Infrastructure Commissioning",
        "Physical Server Decommission"
    ];

    private static List<string> GetSafetyManagerTasks() =>
    [
        "Triconex SMM Release Verification",
        "TCMI Hardware Procurement",
        "Safety System Backup",
        "TCMI Installation & Wiring",
        "Safety Manager Configuration Upload",
        "Proof Test Execution",
        "Safety System Recommission"
    ];

    private static List<string> GetUocTasks() =>
    [
        "UOC Hardware Procurement",
        "Profibus IO Assessment (if applicable)",
        "UOC Cabinet Installation",
        "IO Wiring & Connection",
        "UOC Control Strategy Configuration",
        "UOC Commissioning & Loop Test",
        "Legacy Controller Decommission"
    ];

    private static List<string> GetHardwareRefreshTasks() =>
    [
        "Hardware Inventory & End-of-Life Assessment",
        "New Server / Workstation Procurement",
        "Network Switch Procurement",
        "Hardware Installation & Rack Build",
        "OS Build & Patch (Windows Server / Windows 10+)",
        "Experion Software Re-installation",
        "Data Migration from Old Hardware",
        "Old Hardware Decommission"
    ];
}
