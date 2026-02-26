namespace MigrationScheduler.Blazor.Models;

/// <summary>
/// Represents a single project task/activity in the migration schedule.
/// </summary>
public class TaskModel
{
    public int TaskId { get; set; }
    public string WbsCode { get; set; } = string.Empty;
    public string TaskName { get; set; } = string.Empty;
    public MigrationType MigrationTypeTag { get; set; } = MigrationType.General;
    public string? UspNumber { get; set; }
    public int DurationDays { get; set; } = 1;
    public DateTime? StartDate { get; set; }
    public DateTime? FinishDate { get; set; }
    public string? ResourceOwner { get; set; }
    public string? Notes { get; set; }
    public int PercentComplete { get; set; }

    // Constraint fields
    public ConstraintType ConstraintType { get; set; } = ConstraintType.ASAP;
    public DateTime? ConstraintDate { get; set; }
    public DateTime? Deadline { get; set; }

    // Predecessor links
    public List<PredecessorLinkModel> Predecessors { get; set; } = [];

    // Computed
    public int OutlineLevel => string.IsNullOrEmpty(WbsCode)
        ? 1
        : WbsCode.Split('.').Length;

    public bool IsSummaryTask { get; set; }

    /// <summary>
    /// Calculates finish date from start date and duration (business days).
    /// </summary>
    public void CalculateFinishDate()
    {
        if (StartDate.HasValue && DurationDays > 0)
        {
            var date = StartDate.Value;
            var remaining = DurationDays - 1; // Start date counts as day 1
            while (remaining > 0)
            {
                date = date.AddDays(1);
                if (date.DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday))
                    remaining--;
            }
            FinishDate = date;
        }
    }

    public TaskModel Clone() => new()
    {
        TaskId = TaskId,
        WbsCode = WbsCode,
        TaskName = TaskName,
        MigrationTypeTag = MigrationTypeTag,
        UspNumber = UspNumber,
        DurationDays = DurationDays,
        StartDate = StartDate,
        FinishDate = FinishDate,
        ResourceOwner = ResourceOwner,
        Notes = Notes,
        PercentComplete = PercentComplete,
        ConstraintType = ConstraintType,
        ConstraintDate = ConstraintDate,
        Deadline = Deadline,
        Predecessors = Predecessors.Select(p => p.Clone()).ToList(),
        IsSummaryTask = IsSummaryTask
    };
}
