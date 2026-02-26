using System.Text;
using System.Xml;
using MigrationScheduler.Blazor.Models;

namespace MigrationScheduler.Blazor.Services;

/// <summary>
/// Generates Microsoft Project-compatible XML from the project model.
/// Conforms to the MSP XML schema (xmlns="http://schemas.microsoft.com/project").
/// Compatible with MSP 2007/2010/2013/2016/2019/2021.
/// </summary>
public class XmlExportService
{
    private const string MspNamespace = "http://schemas.microsoft.com/project";

    /// <summary>
    /// Validates the project and returns any error messages.
    /// </summary>
    public List<string> ValidateForExport(ProjectModel project)
    {
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(project.ProjectName))
            errors.Add("Project Name is required.");
        if (string.IsNullOrWhiteSpace(project.UspNumber))
            errors.Add("USP Number is required.");
        if (string.IsNullOrWhiteSpace(project.ProjectManager))
            errors.Add("Project Manager is required.");
        if (!project.ProjectStartDate.HasValue)
            errors.Add("Project Start Date is required.");
        if (!project.TargetCompletionDate.HasValue)
            errors.Add("Target Completion Date is required.");
        if (project.Tasks.Count == 0)
            errors.Add("Project must have at least one task.");

        foreach (var task in project.Tasks.Where(t => !t.IsSummaryTask))
        {
            if (string.IsNullOrWhiteSpace(task.TaskName))
                errors.Add($"Task {task.TaskId}: Task Name is required.");
            if (string.IsNullOrWhiteSpace(task.WbsCode))
                errors.Add($"Task {task.TaskId}: WBS Code is required.");
            if (task.DurationDays <= 0)
                errors.Add($"Task {task.TaskId} ({task.TaskName}): Duration must be greater than 0.");
        }

        // Check for circular dependencies
        var circularErrors = DetectCircularDependencies(project.Tasks);
        errors.AddRange(circularErrors);

        return errors;
    }

    /// <summary>
    /// Detects circular dependencies in the task list.
    /// </summary>
    private static List<string> DetectCircularDependencies(List<TaskModel> tasks)
    {
        var errors = new List<string>();
        var taskMap = tasks.ToDictionary(t => t.TaskId);

        foreach (var task in tasks)
        {
            var visited = new HashSet<int>();
            if (HasCycle(task.TaskId, taskMap, visited))
            {
                errors.Add($"Circular dependency detected involving Task {task.TaskId} ({task.TaskName}).");
                break; // Report once
            }
        }

        return errors;
    }

    private static bool HasCycle(int taskId, Dictionary<int, TaskModel> taskMap, HashSet<int> visited)
    {
        if (!visited.Add(taskId))
            return true;

        if (taskMap.TryGetValue(taskId, out var task))
        {
            foreach (var pred in task.Predecessors)
            {
                if (HasCycle(pred.PredecessorTaskId, taskMap, new HashSet<int>(visited)))
                    return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Generates the MSP-compatible XML document as a string.
    /// </summary>
    public string GenerateXml(ProjectModel project)
    {
        var settings = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            Encoding = Encoding.UTF8,
            OmitXmlDeclaration = false
        };

        using var stream = new MemoryStream();
        using (var writer = XmlWriter.Create(stream, settings))
        {
            writer.WriteStartDocument(true);
            writer.WriteStartElement("Project", MspNamespace);

            WriteProjectHeader(writer, project);
            WriteCalendars(writer);
            WriteExtendedAttributeDefinitions(writer);
            WriteTasks(writer, project);
            WriteResources(writer, project);
            WriteAssignments(writer, project);

            writer.WriteEndElement(); // Project
            writer.WriteEndDocument();
        }

        return Encoding.UTF8.GetString(stream.ToArray());
    }

    /// <summary>
    /// Gets the default export filename.
    /// </summary>
    public static string GetDefaultExportFileName(ProjectModel project)
    {
        var usp = string.IsNullOrWhiteSpace(project.UspNumber) ? "PROJECT" : project.UspNumber;
        var name = string.IsNullOrWhiteSpace(project.ProjectName) ? "Schedule" : project.ProjectName.Replace(" ", "_");
        return $"{usp}_{name}_Schedule.xml";
    }

    private static void WriteProjectHeader(XmlWriter writer, ProjectModel project)
    {
        writer.WriteElementString("Name", project.ProjectName);
        writer.WriteElementString("Title", project.ProjectName);
        writer.WriteElementString("Subject", project.UspNumber);
        writer.WriteElementString("Author", project.ProjectManager);
        writer.WriteElementString("Company", "Honeywell");
        writer.WriteElementString("Manager", project.ProjectManager);

        if (!string.IsNullOrWhiteSpace(project.ProjectDescription))
            writer.WriteElementString("Notes", project.ProjectDescription);

        writer.WriteElementString("CreationDate", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"));
        writer.WriteElementString("StartDate", FormatDate(project.ProjectStartDate, "08:00:00"));
        writer.WriteElementString("FinishDate", FormatDate(project.TargetCompletionDate, "17:00:00"));
        writer.WriteElementString("ScheduleFromStart", "1");
        writer.WriteElementString("CalendarUID", "1");
        writer.WriteElementString("DefaultStartTime", "08:00:00");
        writer.WriteElementString("DefaultFinishTime", "17:00:00");
        writer.WriteElementString("MinutesPerDay", "480");
        writer.WriteElementString("MinutesPerWeek", "2400");
        writer.WriteElementString("DaysPerMonth", "20");
        writer.WriteElementString("CurrencySymbol", "$");
        writer.WriteElementString("CurrencyCode", "USD");
    }

    private static void WriteCalendars(XmlWriter writer)
    {
        writer.WriteStartElement("Calendars");
        writer.WriteStartElement("Calendar");

        writer.WriteElementString("UID", "1");
        writer.WriteElementString("Name", "Standard");
        writer.WriteElementString("IsBaseCalendar", "1");
        writer.WriteElementString("BaseCalendarUID", "-1");

        writer.WriteStartElement("WeekDays");

        // Sunday - non-working
        WriteWeekDay(writer, 1, false);
        // Monday through Friday - working
        for (int day = 2; day <= 6; day++)
            WriteWeekDay(writer, day, true);
        // Saturday - non-working
        WriteWeekDay(writer, 7, false);

        writer.WriteEndElement(); // WeekDays
        writer.WriteEndElement(); // Calendar
        writer.WriteEndElement(); // Calendars
    }

    private static void WriteWeekDay(XmlWriter writer, int dayType, bool isWorking)
    {
        writer.WriteStartElement("WeekDay");
        writer.WriteElementString("DayType", dayType.ToString());
        writer.WriteElementString("DayWorking", isWorking ? "1" : "0");

        if (isWorking)
        {
            writer.WriteStartElement("WorkingTimes");

            writer.WriteStartElement("WorkingTime");
            writer.WriteElementString("FromTime", "08:00:00");
            writer.WriteElementString("ToTime", "12:00:00");
            writer.WriteEndElement(); // WorkingTime

            writer.WriteStartElement("WorkingTime");
            writer.WriteElementString("FromTime", "13:00:00");
            writer.WriteElementString("ToTime", "17:00:00");
            writer.WriteEndElement(); // WorkingTime

            writer.WriteEndElement(); // WorkingTimes
        }

        writer.WriteEndElement(); // WeekDay
    }

    private static void WriteExtendedAttributeDefinitions(XmlWriter writer)
    {
        writer.WriteStartElement("ExtendedAttributes");

        // Text1 = Resource/Owner
        writer.WriteStartElement("ExtendedAttribute");
        writer.WriteElementString("FieldID", "188743731");
        writer.WriteElementString("FieldName", "Text1");
        writer.WriteElementString("Alias", "Resource/Owner");
        writer.WriteEndElement();

        // Text2 = Migration Type
        writer.WriteStartElement("ExtendedAttribute");
        writer.WriteElementString("FieldID", "188743734");
        writer.WriteElementString("FieldName", "Text2");
        writer.WriteElementString("Alias", "Migration Type");
        writer.WriteEndElement();

        // Text3 = USP Number
        writer.WriteStartElement("ExtendedAttribute");
        writer.WriteElementString("FieldID", "188743737");
        writer.WriteElementString("FieldName", "Text3");
        writer.WriteElementString("Alias", "USP Number");
        writer.WriteEndElement();

        writer.WriteEndElement(); // ExtendedAttributes
    }

    private static void WriteTasks(XmlWriter writer, ProjectModel project)
    {
        writer.WriteStartElement("Tasks");

        // Task UID 0 is the project summary task (required by MSP)
        writer.WriteStartElement("Task");
        writer.WriteElementString("UID", "0");
        writer.WriteElementString("ID", "0");
        writer.WriteElementString("Name", project.ProjectName);
        writer.WriteElementString("Type", "1");
        writer.WriteElementString("IsNull", "0");
        writer.WriteElementString("OutlineLevel", "0");
        writer.WriteElementString("OutlineNumber", "0");
        writer.WriteElementString("WBS", "0");
        writer.WriteElementString("Summary", "1");
        writer.WriteEndElement();

        foreach (var task in project.Tasks)
        {
            writer.WriteStartElement("Task");

            writer.WriteElementString("UID", task.TaskId.ToString());
            writer.WriteElementString("ID", task.TaskId.ToString());
            writer.WriteElementString("Name", task.TaskName);
            writer.WriteElementString("Type", "0"); // Fixed Units
            writer.WriteElementString("IsNull", "0");
            writer.WriteElementString("WBS", task.WbsCode);
            writer.WriteElementString("OutlineLevel", task.OutlineLevel.ToString());
            writer.WriteElementString("OutlineNumber", task.WbsCode);
            writer.WriteElementString("Summary", task.IsSummaryTask ? "1" : "0");

            if (!task.IsSummaryTask)
            {
                // Duration: PT{hours}H0M0S where hours = days * 8
                var durationHours = task.DurationDays * 8;
                writer.WriteElementString("Duration", $"PT{durationHours}H0M0S");
                writer.WriteElementString("DurationFormat", "7"); // 7 = days

                // Start/Finish dates
                if (task.StartDate.HasValue)
                    writer.WriteElementString("Start", FormatDate(task.StartDate, "08:00:00"));
                if (task.FinishDate.HasValue)
                    writer.WriteElementString("Finish", FormatDate(task.FinishDate, "17:00:00"));
                else if (task.StartDate.HasValue)
                {
                    // Calculate finish
                    task.CalculateFinishDate();
                    if (task.FinishDate.HasValue)
                        writer.WriteElementString("Finish", FormatDate(task.FinishDate, "17:00:00"));
                }

                writer.WriteElementString("PercentComplete", task.PercentComplete.ToString());
                writer.WriteElementString("Priority", "500"); // Normal priority

                // Notes: combine Notes, USP, and Migration Type
                var notes = BuildNotesString(task);
                if (!string.IsNullOrEmpty(notes))
                    writer.WriteElementString("Notes", notes);

                // Constraint
                writer.WriteElementString("ConstraintType", ((int)task.ConstraintType).ToString());
                if (task.ConstraintType.RequiresDate() && task.ConstraintDate.HasValue)
                    writer.WriteElementString("ConstraintDate", FormatDate(task.ConstraintDate, "00:00:00"));

                // Deadline
                if (task.Deadline.HasValue)
                    writer.WriteElementString("Deadline", FormatDate(task.Deadline, "17:00:00"));

                // Predecessor links
                foreach (var pred in task.Predecessors)
                {
                    writer.WriteStartElement("PredecessorLink");
                    writer.WriteElementString("PredecessorUID", pred.PredecessorTaskId.ToString());
                    writer.WriteElementString("Type", ((int)pred.Type).ToString());
                    writer.WriteElementString("CrossProject", "0");
                    // MSP lag = days * 4800 (tenths of minutes: 1 day = 8h * 60m * 10 = 4800)
                    writer.WriteElementString("LinkLag", (pred.LagDays * 4800).ToString());
                    writer.WriteElementString("LagFormat", "7"); // 7 = days
                    writer.WriteEndElement(); // PredecessorLink
                }

                // Extended attributes
                if (!string.IsNullOrWhiteSpace(task.ResourceOwner))
                {
                    writer.WriteStartElement("ExtendedAttribute");
                    writer.WriteElementString("FieldID", "188743731");
                    writer.WriteElementString("Value", task.ResourceOwner);
                    writer.WriteEndElement();
                }

                writer.WriteStartElement("ExtendedAttribute");
                writer.WriteElementString("FieldID", "188743734");
                writer.WriteElementString("Value", task.MigrationTypeTag.GetShortName());
                writer.WriteEndElement();

                if (!string.IsNullOrWhiteSpace(task.UspNumber))
                {
                    writer.WriteStartElement("ExtendedAttribute");
                    writer.WriteElementString("FieldID", "188743737");
                    writer.WriteElementString("Value", task.UspNumber);
                    writer.WriteEndElement();
                }
            }

            writer.WriteElementString("CalendarUID", "-1");
            writer.WriteEndElement(); // Task
        }

        writer.WriteEndElement(); // Tasks
    }

    private static void WriteResources(XmlWriter writer, ProjectModel project)
    {
        writer.WriteStartElement("Resources");

        // Extract unique resource names from tasks
        var resources = project.Tasks
            .Where(t => !string.IsNullOrWhiteSpace(t.ResourceOwner))
            .Select(t => t.ResourceOwner!)
            .Distinct()
            .OrderBy(r => r)
            .ToList();

        int uid = 1;
        foreach (var resource in resources)
        {
            writer.WriteStartElement("Resource");
            writer.WriteElementString("UID", uid.ToString());
            writer.WriteElementString("ID", uid.ToString());
            writer.WriteElementString("Name", resource);
            writer.WriteElementString("Type", "1"); // 1 = Work resource
            writer.WriteElementString("IsNull", "0");
            writer.WriteElementString("MaxUnits", "1.00");
            writer.WriteElementString("CalendarUID", "1");
            writer.WriteEndElement(); // Resource
            uid++;
        }

        writer.WriteEndElement(); // Resources
    }

    private static void WriteAssignments(XmlWriter writer, ProjectModel project)
    {
        writer.WriteStartElement("Assignments");

        // Build resource UID lookup
        var resourceMap = new Dictionary<string, int>();
        int rUid = 1;
        foreach (var resource in project.Tasks
            .Where(t => !string.IsNullOrWhiteSpace(t.ResourceOwner))
            .Select(t => t.ResourceOwner!)
            .Distinct()
            .OrderBy(r => r))
        {
            resourceMap[resource] = rUid++;
        }

        int assignmentUid = 1;
        foreach (var task in project.Tasks.Where(t => !t.IsSummaryTask && !string.IsNullOrWhiteSpace(t.ResourceOwner)))
        {
            if (resourceMap.TryGetValue(task.ResourceOwner!, out var resourceId))
            {
                writer.WriteStartElement("Assignment");
                writer.WriteElementString("UID", assignmentUid.ToString());
                writer.WriteElementString("TaskUID", task.TaskId.ToString());
                writer.WriteElementString("ResourceUID", resourceId.ToString());
                writer.WriteElementString("Units", "1");
                writer.WriteEndElement(); // Assignment
                assignmentUid++;
            }
        }

        writer.WriteEndElement(); // Assignments
    }

    private static string BuildNotesString(TaskModel task)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(task.Notes))
            parts.Add(task.Notes);
        if (!string.IsNullOrWhiteSpace(task.UspNumber))
            parts.Add($"USP: {task.UspNumber}");
        parts.Add($"Type: {task.MigrationTypeTag.GetShortName()}");
        return string.Join(" | ", parts);
    }

    private static string FormatDate(DateTime? date, string timeComponent)
    {
        if (!date.HasValue) return string.Empty;
        return $"{date.Value:yyyy-MM-dd}T{timeComponent}";
    }
}
