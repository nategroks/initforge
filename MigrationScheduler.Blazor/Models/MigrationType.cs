namespace MigrationScheduler.Blazor.Models;

/// <summary>
/// Defines the available Honeywell Experion migration categories.
/// </summary>
public enum MigrationType
{
    General,
    ELCN,
    EUCN,
    AMT,
    Virtualization,
    SafetyManager,
    UOC,
    HardwareRefresh
}

public static class MigrationTypeExtensions
{
    public static string GetDisplayName(this MigrationType type) => type switch
    {
        MigrationType.General => "General",
        MigrationType.ELCN => "ELCN — Enhanced Local Control Network",
        MigrationType.EUCN => "EUCN — Enhanced Universal Control Network",
        MigrationType.AMT => "AMT — Advanced Migration Technology",
        MigrationType.Virtualization => "Virtualization — Experion Virtualization",
        MigrationType.SafetyManager => "Safety Manager — Safety Manager / TCMI",
        MigrationType.UOC => "UOC — Universal Operations Controller",
        MigrationType.HardwareRefresh => "Hardware Refresh — Hardware Refresh",
        _ => type.ToString()
    };

    public static string GetShortName(this MigrationType type) => type switch
    {
        MigrationType.General => "General",
        MigrationType.ELCN => "ELCN",
        MigrationType.EUCN => "EUCN",
        MigrationType.AMT => "AMT",
        MigrationType.Virtualization => "Virtualization",
        MigrationType.SafetyManager => "Safety Manager",
        MigrationType.UOC => "UOC",
        MigrationType.HardwareRefresh => "Hardware Refresh",
        _ => type.ToString()
    };

    public static string GetDescription(this MigrationType type) => type switch
    {
        MigrationType.ELCN => "LCN to ELCN / FTE network migration",
        MigrationType.EUCN => "xPM (HPM/APM/PM) to EHPM migration via NIM→ENIM/ENB",
        MigrationType.AMT => "ControlEdge PLC / RTU migrations (LM, LEPIU, ELMM)",
        MigrationType.Virtualization => "Migration of physical servers to VMware vSphere virtual machines",
        MigrationType.SafetyManager => "Triconex Communication Module Interface upgrade",
        MigrationType.UOC => "Migration to ControlEdge UOC",
        MigrationType.HardwareRefresh => "Server/workstation/switch hardware replacement",
        _ => string.Empty
    };

    public static string GetCode(this MigrationType type) => type switch
    {
        MigrationType.General => "GEN",
        MigrationType.ELCN => "ELCN",
        MigrationType.EUCN => "EUCN",
        MigrationType.AMT => "AMT",
        MigrationType.Virtualization => "VIRT",
        MigrationType.SafetyManager => "SM",
        MigrationType.UOC => "UOC",
        MigrationType.HardwareRefresh => "HW",
        _ => "GEN"
    };
}
