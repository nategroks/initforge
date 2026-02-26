namespace MigrationScheduler.Blazor.Models;

/// <summary>
/// Microsoft Project constraint types with their MSP XML numeric codes.
/// </summary>
public enum ConstraintType
{
    ASAP = 0,
    ALAP = 1,
    MSO = 2,
    MFO = 3,
    FNET = 4,
    FNLT = 5,
    SNET = 6,
    SNLT = 7
}

public static class ConstraintTypeExtensions
{
    public static string GetDisplayName(this ConstraintType type) => type switch
    {
        ConstraintType.ASAP => "As Soon As Possible (ASAP)",
        ConstraintType.ALAP => "As Late As Possible (ALAP)",
        ConstraintType.MSO => "Must Start On (MSO)",
        ConstraintType.MFO => "Must Finish On (MFO)",
        ConstraintType.FNET => "Finish No Earlier Than (FNET)",
        ConstraintType.FNLT => "Finish No Later Than (FNLT)",
        ConstraintType.SNET => "Start No Earlier Than (SNET)",
        ConstraintType.SNLT => "Start No Later Than (SNLT)",
        _ => type.ToString()
    };

    public static bool RequiresDate(this ConstraintType type) =>
        type is not (ConstraintType.ASAP or ConstraintType.ALAP);
}
