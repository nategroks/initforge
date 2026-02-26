using System.Text.Json;
using System.Text.Json.Serialization;
using MigrationScheduler.Blazor.Models;

namespace MigrationScheduler.Blazor.Services;

/// <summary>
/// Handles saving and loading project drafts as JSON files.
/// </summary>
public class DraftService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    /// <summary>
    /// Serializes the project to a JSON string.
    /// </summary>
    public string SerializeProject(ProjectModel project)
    {
        return JsonSerializer.Serialize(project, JsonOptions);
    }

    /// <summary>
    /// Deserializes a project from a JSON string.
    /// </summary>
    public ProjectModel? DeserializeProject(string json)
    {
        return JsonSerializer.Deserialize<ProjectModel>(json, JsonOptions);
    }

    /// <summary>
    /// Gets the default filename for a draft save.
    /// </summary>
    public static string GetDefaultDraftFileName(ProjectModel project)
    {
        var usp = string.IsNullOrWhiteSpace(project.UspNumber) ? "DRAFT" : project.UspNumber;
        return $"{usp}_draft.json";
    }

    /// <summary>
    /// Saves project state to a file path.
    /// </summary>
    public async Task SaveDraftAsync(ProjectModel project, string filePath)
    {
        var json = SerializeProject(project);
        await File.WriteAllTextAsync(filePath, json);
    }

    /// <summary>
    /// Loads project state from a file path.
    /// </summary>
    public async Task<ProjectModel?> LoadDraftAsync(string filePath)
    {
        if (!File.Exists(filePath))
            return null;

        var json = await File.ReadAllTextAsync(filePath);
        return DeserializeProject(json);
    }
}
