using System.ComponentModel.DataAnnotations;

namespace MigrationScheduler.Blazor.Models;

/// <summary>
/// Root model representing the entire migration project.
/// </summary>
public class ProjectModel
{
    [Required(ErrorMessage = "Project Name is required.")]
    public string ProjectName { get; set; } = string.Empty;

    [Required(ErrorMessage = "USP Number is required.")]
    [RegularExpression(@"^USP-\d{6}$", ErrorMessage = "USP Number must be in format USP-XXXXXX (6 digits).")]
    public string UspNumber { get; set; } = string.Empty;

    [Required(ErrorMessage = "Customer / Site Name is required.")]
    public string CustomerSiteName { get; set; } = string.Empty;

    [Required(ErrorMessage = "Project Manager is required.")]
    public string ProjectManager { get; set; } = string.Empty;

    [Required(ErrorMessage = "Project Start Date is required.")]
    public DateTime? ProjectStartDate { get; set; }

    [Required(ErrorMessage = "Target Completion Date is required.")]
    public DateTime? TargetCompletionDate { get; set; }

    public string? ProjectDescription { get; set; }

    /// <summary>
    /// Migration types selected for this project.
    /// </summary>
    public HashSet<MigrationType> SelectedTypes { get; set; } = [];

    /// <summary>
    /// Migration types that have been previously completed (brownfield).
    /// </summary>
    public HashSet<MigrationType> PreExistingTypes { get; set; } = [];

    /// <summary>
    /// All tasks in the project.
    /// </summary>
    public List<TaskModel> Tasks { get; set; } = [];

    /// <summary>
    /// Tracks whether the setup wizard has been completed.
    /// </summary>
    public bool SetupComplete { get; set; }

    /// <summary>
    /// Tracks whether the activity editor has been reviewed.
    /// </summary>
    public bool TasksComplete { get; set; }

    /// <summary>
    /// Tracks whether constraints have been reviewed.
    /// </summary>
    public bool ConstraintsComplete { get; set; }

    /// <summary>
    /// Gets the next available Task ID.
    /// </summary>
    public int GetNextTaskId() => Tasks.Count > 0 ? Tasks.Max(t => t.TaskId) + 1 : 1;
}
