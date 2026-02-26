namespace MigrationScheduler.Blazor.Models;

/// <summary>
/// Dependency type codes matching Microsoft Project XML schema.
/// </summary>
public enum DependencyType
{
    FF = 0,
    FS = 1,
    SF = 2,
    SS = 3
}

public static class DependencyTypeExtensions
{
    public static string GetDisplayName(this DependencyType type) => type switch
    {
        DependencyType.FF => "Finish-to-Finish (FF)",
        DependencyType.FS => "Finish-to-Start (FS)",
        DependencyType.SF => "Start-to-Finish (SF)",
        DependencyType.SS => "Start-to-Start (SS)",
        _ => type.ToString()
    };
}

/// <summary>
/// Represents a predecessor link between two tasks.
/// </summary>
public class PredecessorLinkModel
{
    public int PredecessorTaskId { get; set; }
    public DependencyType Type { get; set; } = DependencyType.FS;
    public int LagDays { get; set; }

    public PredecessorLinkModel() { }

    public PredecessorLinkModel(int predecessorTaskId, DependencyType type = DependencyType.FS, int lagDays = 0)
    {
        PredecessorTaskId = predecessorTaskId;
        Type = type;
        LagDays = lagDays;
    }

    public PredecessorLinkModel Clone() => new()
    {
        PredecessorTaskId = PredecessorTaskId,
        Type = Type,
        LagDays = LagDays
    };
}
