using MigrationScheduler.Blazor.Models;

namespace MigrationScheduler.Blazor.Services;

/// <summary>
/// Validates prerequisite dependency rules between migration types.
/// </summary>
public class PrerequisiteValidator
{
    /// <summary>
    /// Returns a list of validation error messages for unmet prerequisites.
    /// </summary>
    public List<string> Validate(ProjectModel project)
    {
        var errors = new List<string>();
        var selected = project.SelectedTypes;
        var preExisting = project.PreExistingTypes;

        bool Has(MigrationType t) => selected.Contains(t) || preExisting.Contains(t);

        if (selected.Contains(MigrationType.EUCN) && !Has(MigrationType.ELCN))
            errors.Add("EUCN requires ELCN to be selected or marked as pre-existing.");

        if (selected.Contains(MigrationType.AMT) && !Has(MigrationType.EUCN))
            errors.Add("AMT requires EUCN to be selected or marked as pre-existing.");

        if (selected.Contains(MigrationType.SafetyManager))
        {
            if (!Has(MigrationType.EUCN))
                errors.Add("Safety Manager requires EUCN to be selected or marked as pre-existing.");
            if (!Has(MigrationType.ELCN))
                errors.Add("Safety Manager requires ELCN to be selected or marked as pre-existing.");
        }

        if (selected.Contains(MigrationType.UOC) && !Has(MigrationType.ELCN))
            errors.Add("UOC requires ELCN to be selected or marked as pre-existing.");

        if (selected.Contains(MigrationType.Virtualization) && !Has(MigrationType.HardwareRefresh))
            errors.Add("Virtualization requires Hardware Refresh to be selected or marked as pre-existing.");

        return errors;
    }
}
