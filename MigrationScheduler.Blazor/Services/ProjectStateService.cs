using MigrationScheduler.Blazor.Models;

namespace MigrationScheduler.Blazor.Services;

/// <summary>
/// Singleton service that holds the current project state across Blazor navigation.
/// </summary>
public class ProjectStateService
{
    public ProjectModel Project { get; set; } = new();

    public event Action? OnChange;

    public void NotifyStateChanged() => OnChange?.Invoke();

    /// <summary>
    /// Resets the project to a fresh state.
    /// </summary>
    public void Reset()
    {
        Project = new ProjectModel();
        NotifyStateChanged();
    }

    /// <summary>
    /// Replaces the current project with a loaded one (from draft).
    /// </summary>
    public void LoadProject(ProjectModel project)
    {
        Project = project;
        NotifyStateChanged();
    }

    /// <summary>
    /// Checks whether the project header fields are valid.
    /// </summary>
    public bool IsHeaderValid()
    {
        var p = Project;
        return !string.IsNullOrWhiteSpace(p.ProjectName)
            && !string.IsNullOrWhiteSpace(p.UspNumber)
            && System.Text.RegularExpressions.Regex.IsMatch(p.UspNumber, @"^USP-\d{6}$")
            && !string.IsNullOrWhiteSpace(p.CustomerSiteName)
            && !string.IsNullOrWhiteSpace(p.ProjectManager)
            && p.ProjectStartDate.HasValue
            && p.TargetCompletionDate.HasValue
            && p.TargetCompletionDate > p.ProjectStartDate;
    }

    /// <summary>
    /// Checks if at least one migration type is selected.
    /// </summary>
    public bool HasSelectedTypes() => Project.SelectedTypes.Count > 0;
}
