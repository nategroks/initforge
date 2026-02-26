# ============================================================================
# Honeywell Experion Migration Scheduler - Project Setup Script
# ============================================================================
# Run this script from the git repo root: C:\_hw_sw\initforgee\initforge
# It will create the MigrationScheduler/ directory and all project files.
# ============================================================================

$ErrorActionPreference = "Stop"
$utf8NoBom = New-Object System.Text.UTF8Encoding $false

function Write-ProjectFile {
    param(
        [string]$RelativePath,
        [string]$Content
    )
    $fullPath = Join-Path $PSScriptRoot $RelativePath
    $dir = Split-Path $fullPath -Parent
    if (!(Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($fullPath, $Content, $utf8NoBom)
    Write-Host "  Created: $RelativePath"
}

Write-Host ""
Write-Host "=============================================="
Write-Host " Migration Scheduler - Project File Generator"
Write-Host "=============================================="
Write-Host ""

# --- Create directory structure ---
$dirs = @(
    "MigrationScheduler"
    "MigrationScheduler\MigrationScheduler.Blazor"
    "MigrationScheduler\MigrationScheduler.Blazor\Components"
    "MigrationScheduler\MigrationScheduler.Blazor\Models"
    "MigrationScheduler\MigrationScheduler.Blazor\Pages"
    "MigrationScheduler\MigrationScheduler.Blazor\Properties"
    "MigrationScheduler\MigrationScheduler.Blazor\Services"
    "MigrationScheduler\MigrationScheduler.Blazor\Shared"
    "MigrationScheduler\MigrationScheduler.Blazor\wwwroot"
    "MigrationScheduler\MigrationScheduler.Blazor\wwwroot\css"
    "MigrationScheduler\MigrationScheduler.Blazor\wwwroot\js"
    "MigrationScheduler\MigrationScheduler.Host"
)

foreach ($d in $dirs) {
    $fullDir = Join-Path $PSScriptRoot $d
    if (!(Test-Path $fullDir)) {
        New-Item -ItemType Directory -Path $fullDir -Force | Out-Null
    }
}
Write-Host "Directories created."
Write-Host ""

# ============================================================================
# MigrationScheduler\MigrationScheduler.sln
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.sln" @'

Microsoft Visual Studio Solution File, Format Version 12.00
# Visual Studio Version 17
VisualStudioVersion = 17.8.34525.116
MinimumVisualStudioVersion = 10.0.40219.1
Project("{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}") = "MigrationScheduler.Blazor", "MigrationScheduler.Blazor\MigrationScheduler.Blazor.csproj", "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"
EndProject
Project("{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}") = "MigrationScheduler.Host", "MigrationScheduler.Host\MigrationScheduler.Host.csproj", "{B2C3D4E5-F6A7-8901-BCDE-F12345678901}"
EndProject
Global
	GlobalSection(SolutionConfigurationPlatforms) = preSolution
		Debug|Any CPU = Debug|Any CPU
		Release|Any CPU = Release|Any CPU
	EndGlobalSection
	GlobalSection(ProjectConfigurationPlatforms) = postSolution
		{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}.Debug|Any CPU.ActiveCfg = Debug|Any CPU
		{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}.Debug|Any CPU.Build.0 = Debug|Any CPU
		{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}.Release|Any CPU.ActiveCfg = Release|Any CPU
		{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}.Release|Any CPU.Build.0 = Release|Any CPU
		{B2C3D4E5-F6A7-8901-BCDE-F12345678901}.Debug|Any CPU.ActiveCfg = Debug|Any CPU
		{B2C3D4E5-F6A7-8901-BCDE-F12345678901}.Debug|Any CPU.Build.0 = Debug|Any CPU
		{B2C3D4E5-F6A7-8901-BCDE-F12345678901}.Release|Any CPU.ActiveCfg = Release|Any CPU
		{B2C3D4E5-F6A7-8901-BCDE-F12345678901}.Release|Any CPU.Build.0 = Release|Any CPU
	EndGlobalSection
EndGlobal
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\MigrationScheduler.Blazor.csproj
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\MigrationScheduler.Blazor.csproj" @'
<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <LangVersion>12.0</LangVersion>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="MudBlazor" Version="6.19.1" />
  </ItemGroup>

</Project>
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Program.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Program.cs" @'
using MigrationScheduler.Blazor.Services;
using MudBlazor.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddMudServices();

// Register application services
builder.Services.AddSingleton<ProjectStateService>();
builder.Services.AddScoped<PrerequisiteValidator>();
builder.Services.AddScoped<TaskTemplateService>();
builder.Services.AddScoped<XmlExportService>();
builder.Services.AddScoped<DraftService>();

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseStaticFiles();
app.UseRouting();

app.MapBlazorHub();
app.MapFallbackToPage("/_Host");

app.Run();
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\App.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\App.razor" @'
<Router AppAssembly="@typeof(App).Assembly">
    <Found Context="routeData">
        <RouteView RouteData="@routeData" DefaultLayout="@typeof(Shared.MainLayout)" />
        <FocusOnNavigate RouteData="@routeData" Selector="h1" />
    </Found>
    <NotFound>
        <PageTitle>Not found</PageTitle>
        <LayoutView Layout="@typeof(Shared.MainLayout)">
            <MudText Typo="Typo.h5" Class="mt-8 ml-4">Page not found.</MudText>
        </LayoutView>
    </NotFound>
</Router>
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\_Imports.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\_Imports.razor" @'
@using System.Net.Http
@using Microsoft.AspNetCore.Components.Forms
@using Microsoft.AspNetCore.Components.Routing
@using Microsoft.AspNetCore.Components.Web
@using Microsoft.JSInterop
@using MudBlazor
@using MigrationScheduler.Blazor
@using MigrationScheduler.Blazor.Shared
@using MigrationScheduler.Blazor.Models
@using MigrationScheduler.Blazor.Services
@using MigrationScheduler.Blazor.Components
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\appsettings.json
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\appsettings.json" @'
{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*"
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\appsettings.Development.json
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\appsettings.Development.json" @'
{
  "DetailedErrors": true,
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  }
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Properties\launchSettings.json
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Properties\launchSettings.json" @'
{
  "profiles": {
    "MigrationScheduler.Blazor": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "applicationUrl": "http://localhost:5199",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      }
    }
  }
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Pages\_Host.cshtml
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Pages\_Host.cshtml" @'
@page "/"
@using Microsoft.AspNetCore.Components.Web
@namespace MigrationScheduler.Blazor.Pages
@addTagHelper *, Microsoft.AspNetCore.Mvc.TagHelpers

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <base href="~/" />
    <title>Honeywell Experion Migration Scheduler</title>

    <!-- MudBlazor CSS -->
    <link href="https://fonts.googleapis.com/css?family=Roboto:300,400,500,700&display=swap" rel="stylesheet" />
    <link href="_content/MudBlazor/MudBlazor.min.css" rel="stylesheet" />

    <!-- Application CSS -->
    <link href="css/app.css" rel="stylesheet" />
</head>
<body>
    <component type="typeof(App)" render-mode="ServerPrerendered" />

    <div id="blazor-error-ui">
        <environment include="Staging,Production">
            An error has occurred. This application may no longer respond until reloaded.
        </environment>
        <environment include="Development">
            An unhandled exception has occurred. See browser dev tools for details.
        </environment>
        <a href="" class="reload">Reload</a>
        <a class="dismiss">Dismiss</a>
    </div>

    <script src="_framework/blazor.server.js"></script>
    <script src="_content/MudBlazor/MudBlazor.min.js"></script>
    <script src="js/interop.js"></script>
</body>
</html>
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Pages\ProjectSetup.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Pages\ProjectSetup.razor" @'
@page "/"
@inject ProjectStateService StateService
@inject PrerequisiteValidator PrereqValidator
@inject TaskTemplateService TemplateService
@inject DraftService DraftService
@inject NavigationManager Navigation
@inject ISnackbar Snackbar
@inject IJSRuntime JS

<PageTitle>Project Setup — Migration Scheduler</PageTitle>

<MudContainer MaxWidth="MaxWidth.Large" Class="py-4">
    <MudText Typo="Typo.h4" Class="mb-4">Project Setup Wizard</MudText>

    <MudStepper @bind-ActiveIndex="_activeStep" Linear="true" Color="Color.Primary">
        @* Step 1: Project Header *@
        <MudStep Title="Project Details" Icon="@Icons.Material.Filled.Description">
            <MudPaper Class="pa-6 mb-4" Elevation="1">
                <MudText Typo="Typo.h6" Class="mb-4">Project Header Information</MudText>

                <MudGrid>
                    <MudItem xs="12" md="6">
                        <MudTextField @bind-Value="_project.ProjectName"
                                      Label="Project Name"
                                      Required="true"
                                      RequiredError="Project Name is required"
                                      Variant="Variant.Outlined"
                                      Immediate="true" />
                    </MudItem>
                    <MudItem xs="12" md="6">
                        <MudTextField @bind-Value="_project.UspNumber"
                                      Label="USP Number"
                                      Required="true"
                                      RequiredError="USP Number is required"
                                      Placeholder="USP-XXXXXX"
                                      Validation="@(new Func<string, string?>(ValidateUsp))"
                                      Variant="Variant.Outlined"
                                      Immediate="true" />
                    </MudItem>
                    <MudItem xs="12" md="6">
                        <MudTextField @bind-Value="_project.CustomerSiteName"
                                      Label="Customer / Site Name"
                                      Required="true"
                                      RequiredError="Customer / Site Name is required"
                                      Variant="Variant.Outlined"
                                      Immediate="true" />
                    </MudItem>
                    <MudItem xs="12" md="6">
                        <MudTextField @bind-Value="_project.ProjectManager"
                                      Label="Project Manager"
                                      Required="true"
                                      RequiredError="Project Manager is required"
                                      Variant="Variant.Outlined"
                                      Immediate="true" />
                    </MudItem>
                    <MudItem xs="12" md="6">
                        <MudDatePicker @bind-Date="_project.ProjectStartDate"
                                       Label="Project Start Date"
                                       Required="true"
                                       RequiredError="Project Start Date is required"
                                       Variant="Variant.Outlined"
                                       DateFormat="yyyy-MM-dd" />
                    </MudItem>
                    <MudItem xs="12" md="6">
                        <MudDatePicker @bind-Date="_project.TargetCompletionDate"
                                       Label="Target Completion Date"
                                       Required="true"
                                       RequiredError="Target Completion Date is required"
                                       Variant="Variant.Outlined"
                                       DateFormat="yyyy-MM-dd"
                                       MinDate="@_project.ProjectStartDate" />
                    </MudItem>
                    <MudItem xs="12">
                        <MudTextField @bind-Value="_project.ProjectDescription"
                                      Label="Project Description"
                                      Lines="3"
                                      Variant="Variant.Outlined" />
                    </MudItem>
                </MudGrid>

                <MudDivider Class="my-4" />

                <MudStack Row="true" Justify="Justify.SpaceBetween">
                    <MudStack Row="true" Spacing="2">
                        <MudButton Variant="Variant.Outlined" Color="Color.Default"
                                   StartIcon="@Icons.Material.Filled.UploadFile"
                                   OnClick="LoadDraft">
                            Load Draft
                        </MudButton>
                        <MudButton Variant="Variant.Outlined" Color="Color.Default"
                                   StartIcon="@Icons.Material.Filled.Save"
                                   OnClick="SaveDraft"
                                   Disabled="@(!StateService.IsHeaderValid())">
                            Save Draft
                        </MudButton>
                    </MudStack>
                </MudStack>
            </MudPaper>
        </MudStep>

        @* Step 2: Migration Type Selection *@
        <MudStep Title="Migration Types" Icon="@Icons.Material.Filled.Category">
            <MudPaper Class="pa-6 mb-4" Elevation="1">
                <MudText Typo="Typo.h6" Class="mb-2">Select Migration Types</MudText>
                <MudText Typo="Typo.body2" Color="Color.Secondary" Class="mb-4">
                    Select all migration categories applicable to this project. Multiple selections are allowed.
                </MudText>

                <MigrationTypeSelector SelectedTypes="@_project.SelectedTypes"
                                       PreExistingTypes="@_project.PreExistingTypes"
                                       OnSelectionChanged="OnMigrationTypeChanged" />

                @if (_prerequisiteErrors.Count > 0)
                {
                    <PrerequisiteWarning Errors="@_prerequisiteErrors" />
                }
            </MudPaper>
        </MudStep>

        @* Step 3: Confirm & Proceed *@
        <MudStep Title="Confirm" Icon="@Icons.Material.Filled.CheckCircle">
            <MudPaper Class="pa-6 mb-4" Elevation="1">
                <MudText Typo="Typo.h6" Class="mb-4">Review Project Setup</MudText>

                <MudSimpleTable Dense="true" Bordered="true" Class="mb-4">
                    <tbody>
                        <tr><td><strong>Project Name</strong></td><td>@_project.ProjectName</td></tr>
                        <tr><td><strong>USP Number</strong></td><td>@_project.UspNumber</td></tr>
                        <tr><td><strong>Customer / Site</strong></td><td>@_project.CustomerSiteName</td></tr>
                        <tr><td><strong>Project Manager</strong></td><td>@_project.ProjectManager</td></tr>
                        <tr><td><strong>Start Date</strong></td><td>@_project.ProjectStartDate?.ToString("yyyy-MM-dd")</td></tr>
                        <tr><td><strong>Target Completion</strong></td><td>@_project.TargetCompletionDate?.ToString("yyyy-MM-dd")</td></tr>
                        <tr>
                            <td><strong>Migration Types</strong></td>
                            <td>
                                @foreach (var t in _project.SelectedTypes)
                                {
                                    <MudChip T="string" Size="Size.Small" Color="Color.Primary" Class="mr-1">@t.GetShortName()</MudChip>
                                }
                            </td>
                        </tr>
                        @if (_project.PreExistingTypes.Count > 0)
                        {
                            <tr>
                                <td><strong>Pre-existing</strong></td>
                                <td>
                                    @foreach (var t in _project.PreExistingTypes)
                                    {
                                        <MudChip T="string" Size="Size.Small" Color="Color.Info" Class="mr-1">@t.GetShortName()</MudChip>
                                    }
                                </td>
                            </tr>
                        }
                    </tbody>
                </MudSimpleTable>

                <MudAlert Severity="Severity.Info" Class="mb-4">
                    Proceeding will generate @(_estimatedTaskCount) default tasks based on your selected migration types.
                    You can modify all tasks in the Activity Editor.
                </MudAlert>

                <MudButton Variant="Variant.Filled" Color="Color.Primary"
                           Size="Size.Large"
                           StartIcon="@Icons.Material.Filled.ArrowForward"
                           OnClick="CompleteSetup"
                           Disabled="@(!CanCompleteSetup)">
                    Generate Tasks & Continue to Activity Editor
                </MudButton>
            </MudPaper>
        </MudStep>
    </MudStepper>
</MudContainer>

@code {
    private ProjectModel _project = new();
    private int _activeStep;
    private List<string> _prerequisiteErrors = [];
    private int _estimatedTaskCount;

    protected override void OnInitialized()
    {
        _project = StateService.Project;
        ValidatePrerequisites();
        UpdateTaskEstimate();
    }

    private string? ValidateUsp(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "USP Number is required";
        if (!System.Text.RegularExpressions.Regex.IsMatch(value, @"^USP-\d{6}$"))
            return "Format must be USP-XXXXXX (6 digits)";
        return null;
    }

    private void OnMigrationTypeChanged()
    {
        ValidatePrerequisites();
        UpdateTaskEstimate();
        StateChanged();
    }

    private void ValidatePrerequisites()
    {
        _prerequisiteErrors = PrereqValidator.Validate(_project);
    }

    private void UpdateTaskEstimate()
    {
        var tempTasks = TemplateService.GenerateDefaultTasks(
            _project.SelectedTypes, _project.UspNumber ?? "", _project.ProjectStartDate);
        _estimatedTaskCount = tempTasks.Count(t => !t.IsSummaryTask);
    }

    private bool CanCompleteSetup =>
        StateService.IsHeaderValid()
        && StateService.HasSelectedTypes()
        && _prerequisiteErrors.Count == 0;

    private void CompleteSetup()
    {
        // Generate default tasks if none exist yet, or regenerate if types changed
        _project.Tasks = TemplateService.GenerateDefaultTasks(
            _project.SelectedTypes, _project.UspNumber, _project.ProjectStartDate);

        _project.SetupComplete = true;
        StateService.NotifyStateChanged();
        Navigation.NavigateTo("/tasks");
    }

    private void StateChanged()
    {
        StateService.NotifyStateChanged();
        StateHasChanged();
    }

    private async Task SaveDraft()
    {
        try
        {
            var json = DraftService.SerializeProject(_project);
            var fileName = DraftService.GetDefaultDraftFileName(_project);
            await JS.InvokeVoidAsync("interop.downloadFile", fileName, json, "application/json");
            Snackbar.Add("Draft saved successfully.", Severity.Success);
        }
        catch (Exception ex)
        {
            Snackbar.Add($"Failed to save draft: {ex.Message}", Severity.Error);
        }
    }

    private async Task LoadDraft()
    {
        try
        {
            var json = await JS.InvokeAsync<string>("interop.uploadFile", ".json");
            if (!string.IsNullOrEmpty(json))
            {
                var loaded = DraftService.DeserializeProject(json);
                if (loaded is not null)
                {
                    StateService.LoadProject(loaded);
                    _project = StateService.Project;
                    ValidatePrerequisites();
                    UpdateTaskEstimate();
                    Snackbar.Add("Draft loaded successfully.", Severity.Success);
                    StateHasChanged();
                }
                else
                {
                    Snackbar.Add("Failed to parse draft file.", Severity.Error);
                }
            }
        }
        catch (Exception ex)
        {
            Snackbar.Add($"Failed to load draft: {ex.Message}", Severity.Error);
        }
    }
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Pages\ActivityEditor.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Pages\ActivityEditor.razor" @'
@page "/tasks"
@inject ProjectStateService StateService
@inject NavigationManager Navigation
@inject ISnackbar Snackbar

<PageTitle>Activity Editor — Migration Scheduler</PageTitle>

<MudContainer MaxWidth="MaxWidth.ExtraLarge" Class="py-4">
    <MudStack Row="true" Justify="Justify.SpaceBetween" AlignItems="AlignItems.Center" Class="mb-4">
        <MudText Typo="Typo.h4">Activity Editor</MudText>
        <MudStack Row="true" Spacing="2">
            <MudButton Variant="Variant.Outlined" Color="Color.Default"
                       StartIcon="@Icons.Material.Filled.ArrowBack"
                       OnClick="@(() => Navigation.NavigateTo("/"))">
                Back to Setup
            </MudButton>
            <MudButton Variant="Variant.Filled" Color="Color.Primary"
                       StartIcon="@Icons.Material.Filled.ArrowForward"
                       OnClick="ProceedToConstraints"
                       Disabled="@(!HasValidTasks)">
                Continue to Constraints
            </MudButton>
        </MudStack>
    </MudStack>

    @if (!StateService.Project.SetupComplete)
    {
        <MudAlert Severity="Severity.Warning" Class="mb-4">
            Please complete the Project Setup before editing tasks.
            <MudButton Variant="Variant.Text" Color="Color.Warning" OnClick="@(() => Navigation.NavigateTo("/"))">
                Go to Setup
            </MudButton>
        </MudAlert>
    }
    else
    {
        <MudPaper Class="pa-4 mb-4" Elevation="1">
            <MudStack Row="true" Spacing="2" Class="mb-3">
                <MudButton Variant="Variant.Filled" Color="Color.Primary"
                           StartIcon="@Icons.Material.Filled.Add"
                           OnClick="AddNewTask" Size="Size.Small">
                    Add Task
                </MudButton>
                <MudSelect T="MigrationType?" @bind-Value="_filterType" Label="Filter by Type"
                           Variant="Variant.Outlined" Dense="true" Margin="Margin.Dense"
                           Clearable="true" Style="max-width: 250px;">
                    @foreach (var t in GetAvailableTypes())
                    {
                        <MudSelectItem T="MigrationType?" Value="@t">@t.GetShortName()</MudSelectItem>
                    }
                </MudSelect>
                <MudSpacer />
                <MudText Typo="Typo.body2" Class="align-self-center" Color="Color.Secondary">
                    @_tasks.Count(t => !t.IsSummaryTask) tasks | @_tasks.Count(t => t.IsSummaryTask) sections
                </MudText>
            </MudStack>

            <MudTable Items="@FilteredTasks" Dense="true" Hover="true" Bordered="true" Striped="true"
                      FixedHeader="true" Height="calc(100vh - 320px)"
                      RowClassFunc="@((item, _) => item.IsSummaryTask ? "summary-row" : "")">
                <HeaderContent>
                    <MudTh Style="width:50px">ID</MudTh>
                    <MudTh Style="width:80px">WBS</MudTh>
                    <MudTh>Task Name</MudTh>
                    <MudTh Style="width:130px">Type</MudTh>
                    <MudTh Style="width:70px">Days</MudTh>
                    <MudTh Style="width:130px">Start</MudTh>
                    <MudTh Style="width:130px">Finish</MudTh>
                    <MudTh Style="width:120px">Resource</MudTh>
                    <MudTh Style="width:60px">%</MudTh>
                    <MudTh Style="width:100px">Actions</MudTh>
                </HeaderContent>
                <RowTemplate>
                    <MudTd DataLabel="ID">
                        <MudText Typo="Typo.body2">@context.TaskId</MudText>
                    </MudTd>
                    <MudTd DataLabel="WBS">
                        @if (_editingTaskId == context.TaskId)
                        {
                            <MudTextField @bind-Value="context.WbsCode" Variant="Variant.Text"
                                          Margin="Margin.Dense" Dense="true" />
                        }
                        else
                        {
                            <MudText Typo="Typo.body2" Style="@GetIndentStyle(context)">@context.WbsCode</MudText>
                        }
                    </MudTd>
                    <MudTd DataLabel="Task Name">
                        @if (_editingTaskId == context.TaskId)
                        {
                            <MudTextField @bind-Value="context.TaskName" Variant="Variant.Text"
                                          Margin="Margin.Dense" Dense="true" />
                        }
                        else
                        {
                            <MudText Typo="@(context.IsSummaryTask ? Typo.subtitle2 : Typo.body2)"
                                     Style="@GetIndentStyle(context)">
                                @(context.IsSummaryTask ? $"** {context.TaskName}" : context.TaskName)
                            </MudText>
                        }
                    </MudTd>
                    <MudTd DataLabel="Type">
                        @if (_editingTaskId == context.TaskId)
                        {
                            <MudSelect T="MigrationType" @bind-Value="context.MigrationTypeTag"
                                       Variant="Variant.Text" Dense="true" Margin="Margin.Dense">
                                @foreach (var t in GetAvailableTypes())
                                {
                                    <MudSelectItem T="MigrationType" Value="@t">@t.GetShortName()</MudSelectItem>
                                }
                            </MudSelect>
                        }
                        else
                        {
                            <MudChip T="string" Size="Size.Small" Color="Color.Default">@context.MigrationTypeTag.GetCode()</MudChip>
                        }
                    </MudTd>
                    <MudTd DataLabel="Days">
                        @if (_editingTaskId == context.TaskId && !context.IsSummaryTask)
                        {
                            <MudNumericField T="int" @bind-Value="context.DurationDays"
                                             Variant="Variant.Text" Min="1" Dense="true" Margin="Margin.Dense" />
                        }
                        else if (!context.IsSummaryTask)
                        {
                            <MudText Typo="Typo.body2">@context.DurationDays</MudText>
                        }
                    </MudTd>
                    <MudTd DataLabel="Start">
                        @if (_editingTaskId == context.TaskId && !context.IsSummaryTask)
                        {
                            <MudDatePicker @bind-Date="context.StartDate" Variant="Variant.Text"
                                           Dense="true" Margin="Margin.Dense" DateFormat="yyyy-MM-dd" />
                        }
                        else if (context.StartDate.HasValue)
                        {
                            <MudText Typo="Typo.body2">@context.StartDate.Value.ToString("yyyy-MM-dd")</MudText>
                        }
                    </MudTd>
                    <MudTd DataLabel="Finish">
                        @if (_editingTaskId == context.TaskId && !context.IsSummaryTask)
                        {
                            <MudDatePicker @bind-Date="context.FinishDate" Variant="Variant.Text"
                                           Dense="true" Margin="Margin.Dense" DateFormat="yyyy-MM-dd" />
                        }
                        else if (context.FinishDate.HasValue)
                        {
                            <MudText Typo="Typo.body2">@context.FinishDate.Value.ToString("yyyy-MM-dd")</MudText>
                        }
                    </MudTd>
                    <MudTd DataLabel="Resource">
                        @if (_editingTaskId == context.TaskId && !context.IsSummaryTask)
                        {
                            <MudTextField @bind-Value="context.ResourceOwner" Variant="Variant.Text"
                                          Margin="Margin.Dense" Dense="true" />
                        }
                        else
                        {
                            <MudText Typo="Typo.body2">@context.ResourceOwner</MudText>
                        }
                    </MudTd>
                    <MudTd DataLabel="%">
                        @if (_editingTaskId == context.TaskId && !context.IsSummaryTask)
                        {
                            <MudNumericField T="int" @bind-Value="context.PercentComplete"
                                             Variant="Variant.Text" Min="0" Max="100"
                                             Dense="true" Margin="Margin.Dense" />
                        }
                        else if (!context.IsSummaryTask)
                        {
                            <MudText Typo="Typo.body2">@context.PercentComplete%</MudText>
                        }
                    </MudTd>
                    <MudTd DataLabel="Actions">
                        @if (_editingTaskId == context.TaskId)
                        {
                            <MudIconButton Icon="@Icons.Material.Filled.Check" Size="Size.Small"
                                           Color="Color.Success" OnClick="SaveEdit" Title="Save" />
                            <MudIconButton Icon="@Icons.Material.Filled.Close" Size="Size.Small"
                                           Color="Color.Default" OnClick="CancelEdit" Title="Cancel" />
                        }
                        else
                        {
                            <MudIconButton Icon="@Icons.Material.Filled.Edit" Size="Size.Small"
                                           Color="Color.Primary" OnClick="@(() => StartEdit(context))" Title="Edit" />
                            <MudIconButton Icon="@Icons.Material.Filled.Delete" Size="Size.Small"
                                           Color="Color.Error" OnClick="@(() => DeleteTask(context))" Title="Delete" />
                        }
                    </MudTd>
                </RowTemplate>
            </MudTable>
        </MudPaper>
    }
</MudContainer>

@code {
    private List<TaskModel> _tasks = [];
    private MigrationType? _filterType;
    private int? _editingTaskId;
    private TaskModel? _editBackup;

    private IEnumerable<TaskModel> FilteredTasks =>
        _filterType.HasValue
            ? _tasks.Where(t => t.MigrationTypeTag == _filterType.Value)
            : _tasks;

    private bool HasValidTasks => _tasks.Any(t => !t.IsSummaryTask && !string.IsNullOrWhiteSpace(t.TaskName));

    protected override void OnInitialized()
    {
        _tasks = StateService.Project.Tasks;
    }

    private List<MigrationType> GetAvailableTypes()
    {
        var types = new List<MigrationType> { MigrationType.General };
        types.AddRange(StateService.Project.SelectedTypes.OrderBy(t => t.ToString()));
        return types;
    }

    private string GetIndentStyle(TaskModel task)
    {
        var level = task.OutlineLevel - 1;
        return level > 0 ? $"padding-left: {level * 20}px" : "";
    }

    private void AddNewTask()
    {
        var newTask = new TaskModel
        {
            TaskId = StateService.Project.GetNextTaskId(),
            WbsCode = "",
            TaskName = "New Task",
            MigrationTypeTag = MigrationType.General,
            UspNumber = StateService.Project.UspNumber,
            DurationDays = 5
        };
        _tasks.Add(newTask);
        _editingTaskId = newTask.TaskId;
        StateService.NotifyStateChanged();
    }

    private void StartEdit(TaskModel task)
    {
        _editBackup = task.Clone();
        _editingTaskId = task.TaskId;
    }

    private void SaveEdit()
    {
        var task = _tasks.FirstOrDefault(t => t.TaskId == _editingTaskId);
        if (task is not null && task.StartDate.HasValue && !task.FinishDate.HasValue)
        {
            task.CalculateFinishDate();
        }
        _editingTaskId = null;
        _editBackup = null;
        StateService.NotifyStateChanged();
    }

    private void CancelEdit()
    {
        if (_editBackup is not null && _editingTaskId.HasValue)
        {
            var index = _tasks.FindIndex(t => t.TaskId == _editingTaskId);
            if (index >= 0)
                _tasks[index] = _editBackup;
        }
        _editingTaskId = null;
        _editBackup = null;
    }

    private void DeleteTask(TaskModel task)
    {
        _tasks.Remove(task);
        // Remove any predecessor references to this task
        foreach (var t in _tasks)
            t.Predecessors.RemoveAll(p => p.PredecessorTaskId == task.TaskId);
        StateService.NotifyStateChanged();
    }

    private void ProceedToConstraints()
    {
        StateService.Project.TasksComplete = true;
        StateService.NotifyStateChanged();
        Navigation.NavigateTo("/constraints");
    }
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Pages\ConstraintManager.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Pages\ConstraintManager.razor" @'
@page "/constraints"
@inject ProjectStateService StateService
@inject NavigationManager Navigation
@inject ISnackbar Snackbar

<PageTitle>Constraints & Dependencies — Migration Scheduler</PageTitle>

<MudContainer MaxWidth="MaxWidth.ExtraLarge" Class="py-4">
    <MudStack Row="true" Justify="Justify.SpaceBetween" AlignItems="AlignItems.Center" Class="mb-4">
        <MudText Typo="Typo.h4">Constraint & Dependency Manager</MudText>
        <MudStack Row="true" Spacing="2">
            <MudButton Variant="Variant.Outlined" Color="Color.Default"
                       StartIcon="@Icons.Material.Filled.ArrowBack"
                       OnClick="@(() => Navigation.NavigateTo("/tasks"))">
                Back to Tasks
            </MudButton>
            <MudButton Variant="Variant.Filled" Color="Color.Primary"
                       StartIcon="@Icons.Material.Filled.ArrowForward"
                       OnClick="ProceedToExport">
                Continue to Export
            </MudButton>
        </MudStack>
    </MudStack>

    @if (!StateService.Project.TasksComplete)
    {
        <MudAlert Severity="Severity.Warning" Class="mb-4">
            Please complete the Activity Editor before managing constraints.
        </MudAlert>
    }
    else
    {
        <MudText Typo="Typo.body2" Color="Color.Secondary" Class="mb-4">
            Click on a task to expand its dependency and constraint settings. Define predecessors, constraint types, and deadlines for each task.
        </MudText>

        <MudExpansionPanels MultiExpansion="true" Elevation="1">
            @foreach (var task in _tasks.Where(t => !t.IsSummaryTask))
            {
                <MudExpansionPanel>
                    <TitleContent>
                        <MudStack Row="true" AlignItems="AlignItems.Center" Spacing="2">
                            <MudChip T="string" Size="Size.Small" Color="Color.Primary">@task.TaskId</MudChip>
                            <MudText Typo="Typo.subtitle2">@task.TaskName</MudText>
                            <MudSpacer />
                            @if (task.Predecessors.Count > 0)
                            {
                                <MudChip T="string" Size="Size.Small" Color="Color.Info">
                                    @task.Predecessors.Count pred.
                                </MudChip>
                            }
                            @if (task.ConstraintType != ConstraintType.ASAP)
                            {
                                <MudChip T="string" Size="Size.Small" Color="Color.Warning">
                                    @task.ConstraintType
                                </MudChip>
                            }
                            @if (task.Deadline.HasValue)
                            {
                                <MudChip T="string" Size="Size.Small" Color="Color.Error">
                                    Deadline
                                </MudChip>
                            }
                        </MudStack>
                    </TitleContent>
                    <ChildContent>
                        <MudGrid Class="pa-2">
                            @* Predecessor Section *@
                            <MudItem xs="12">
                                <MudText Typo="Typo.subtitle1" Class="mb-2">Predecessors</MudText>
                            </MudItem>

                            @if (task.Predecessors.Count > 0)
                            {
                                <MudItem xs="12">
                                    <MudSimpleTable Dense="true" Bordered="true" Class="mb-2">
                                        <thead>
                                            <tr>
                                                <th>Predecessor Task</th>
                                                <th style="width:180px">Dependency Type</th>
                                                <th style="width:100px">Lag (days)</th>
                                                <th style="width:60px">Remove</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            @foreach (var pred in task.Predecessors.ToList())
                                            {
                                                var predTask = _tasks.FirstOrDefault(t => t.TaskId == pred.PredecessorTaskId);
                                                <tr>
                                                    <td>@pred.PredecessorTaskId — @(predTask?.TaskName ?? "Unknown")</td>
                                                    <td>
                                                        <MudSelect T="DependencyType" @bind-Value="pred.Type"
                                                                   Variant="Variant.Text" Dense="true" Margin="Margin.Dense">
                                                            @foreach (DependencyType dt in Enum.GetValues<DependencyType>())
                                                            {
                                                                <MudSelectItem T="DependencyType" Value="@dt">@dt.GetDisplayName()</MudSelectItem>
                                                            }
                                                        </MudSelect>
                                                    </td>
                                                    <td>
                                                        <MudNumericField T="int" @bind-Value="pred.LagDays"
                                                                         Variant="Variant.Text" Dense="true" Margin="Margin.Dense" />
                                                    </td>
                                                    <td>
                                                        <MudIconButton Icon="@Icons.Material.Filled.RemoveCircle"
                                                                       Size="Size.Small" Color="Color.Error"
                                                                       OnClick="@(() => RemovePredecessor(task, pred))" />
                                                    </td>
                                                </tr>
                                            }
                                        </tbody>
                                    </MudSimpleTable>
                                </MudItem>
                            }

                            <MudItem xs="12" md="6">
                                <MudStack Row="true" Spacing="2" AlignItems="AlignItems.End">
                                    <MudSelect T="int?" @bind-Value="_selectedPredecessorId" Label="Add Predecessor"
                                               Variant="Variant.Outlined" Dense="true" Margin="Margin.Dense"
                                               Clearable="true" Style="min-width:300px">
                                        @foreach (var candidate in GetPredecessorCandidates(task))
                                        {
                                            <MudSelectItem T="int?" Value="@candidate.TaskId">
                                                @candidate.TaskId — @candidate.TaskName
                                            </MudSelectItem>
                                        }
                                    </MudSelect>
                                    <MudButton Variant="Variant.Outlined" Color="Color.Primary" Size="Size.Small"
                                               OnClick="@(() => AddPredecessor(task))"
                                               Disabled="@(!_selectedPredecessorId.HasValue)">
                                        Add
                                    </MudButton>
                                </MudStack>
                            </MudItem>

                            <MudItem xs="12"><MudDivider Class="my-2" /></MudItem>

                            @* Constraint Section *@
                            <MudItem xs="12">
                                <MudText Typo="Typo.subtitle1" Class="mb-2">Task Constraint</MudText>
                            </MudItem>
                            <MudItem xs="12" md="4">
                                <MudSelect T="ConstraintType" @bind-Value="task.ConstraintType"
                                           Label="Constraint Type" Variant="Variant.Outlined"
                                           Dense="true" Margin="Margin.Dense">
                                    @foreach (ConstraintType ct in Enum.GetValues<ConstraintType>())
                                    {
                                        <MudSelectItem T="ConstraintType" Value="@ct">@ct.GetDisplayName()</MudSelectItem>
                                    }
                                </MudSelect>
                            </MudItem>
                            <MudItem xs="12" md="4">
                                @if (task.ConstraintType.RequiresDate())
                                {
                                    <MudDatePicker @bind-Date="task.ConstraintDate" Label="Constraint Date"
                                                   Variant="Variant.Outlined" Dense="true" Margin="Margin.Dense"
                                                   DateFormat="yyyy-MM-dd"
                                                   Required="true" RequiredError="Constraint date is required for this type" />
                                }
                            </MudItem>
                            <MudItem xs="12" md="4">
                                <MudDatePicker @bind-Date="task.Deadline" Label="Deadline (optional)"
                                               Variant="Variant.Outlined" Dense="true" Margin="Margin.Dense"
                                               DateFormat="yyyy-MM-dd" Clearable="true" />
                            </MudItem>
                        </MudGrid>
                    </ChildContent>
                </MudExpansionPanel>
            }
        </MudExpansionPanels>
    }
</MudContainer>

@code {
    private List<TaskModel> _tasks = [];
    private int? _selectedPredecessorId;

    protected override void OnInitialized()
    {
        _tasks = StateService.Project.Tasks;
    }

    private IEnumerable<TaskModel> GetPredecessorCandidates(TaskModel currentTask)
    {
        return _tasks.Where(t =>
            !t.IsSummaryTask
            && t.TaskId != currentTask.TaskId
            && !currentTask.Predecessors.Any(p => p.PredecessorTaskId == t.TaskId));
    }

    private void AddPredecessor(TaskModel task)
    {
        if (_selectedPredecessorId.HasValue)
        {
            // Check for circular dependency before adding
            if (WouldCreateCycle(task.TaskId, _selectedPredecessorId.Value))
            {
                Snackbar.Add("Cannot add this predecessor: it would create a circular dependency.", Severity.Warning);
                return;
            }

            task.Predecessors.Add(new PredecessorLinkModel(_selectedPredecessorId.Value));
            _selectedPredecessorId = null;
            StateService.NotifyStateChanged();
        }
    }

    private void RemovePredecessor(TaskModel task, PredecessorLinkModel pred)
    {
        task.Predecessors.Remove(pred);
        StateService.NotifyStateChanged();
    }

    private bool WouldCreateCycle(int taskId, int predecessorId)
    {
        // BFS to check if taskId is reachable from predecessorId
        var visited = new HashSet<int>();
        var queue = new Queue<int>();
        queue.Enqueue(taskId);

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            if (current == predecessorId)
                return true;
            if (!visited.Add(current))
                continue;

            // Find all tasks that have 'current' as a predecessor (i.e., current is a successor of them)
            foreach (var t in _tasks)
            {
                if (t.Predecessors.Any(p => p.PredecessorTaskId == current) && !visited.Contains(t.TaskId))
                    queue.Enqueue(t.TaskId);
            }
        }

        return false;
    }

    private void ProceedToExport()
    {
        StateService.Project.ConstraintsComplete = true;
        StateService.NotifyStateChanged();
        Navigation.NavigateTo("/export");
    }
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Pages\ExportPage.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Pages\ExportPage.razor" @'
@page "/export"
@inject ProjectStateService StateService
@inject XmlExportService XmlService
@inject DraftService DraftService
@inject NavigationManager Navigation
@inject ISnackbar Snackbar
@inject IJSRuntime JS

<PageTitle>Export — Migration Scheduler</PageTitle>

<MudContainer MaxWidth="MaxWidth.ExtraLarge" Class="py-4">
    <MudStack Row="true" Justify="Justify.SpaceBetween" AlignItems="AlignItems.Center" Class="mb-4">
        <MudText Typo="Typo.h4">Export to Microsoft Project</MudText>
        <MudButton Variant="Variant.Outlined" Color="Color.Default"
                   StartIcon="@Icons.Material.Filled.ArrowBack"
                   OnClick="@(() => Navigation.NavigateTo("/constraints"))">
            Back to Constraints
        </MudButton>
    </MudStack>

    @if (!StateService.Project.ConstraintsComplete)
    {
        <MudAlert Severity="Severity.Warning" Class="mb-4">
            Please complete the Constraints & Dependencies step before exporting.
        </MudAlert>
    }
    else
    {
        @* Validation Results *@
        @if (_validationErrors.Count > 0)
        {
            <MudAlert Severity="Severity.Error" Class="mb-4">
                <MudText Typo="Typo.subtitle2">Please fix the following errors before exporting:</MudText>
                <MudList T="string" Dense="true">
                    @foreach (var error in _validationErrors)
                    {
                        <MudListItem T="string" Icon="@Icons.Material.Filled.Error" IconColor="Color.Error">
                            @error
                        </MudListItem>
                    }
                </MudList>
            </MudAlert>
        }

        @* Project Summary *@
        <MudGrid>
            <MudItem xs="12" md="4">
                <MudPaper Class="pa-4" Elevation="1">
                    <MudText Typo="Typo.h6" Class="mb-3">Project Summary</MudText>
                    <MudSimpleTable Dense="true">
                        <tbody>
                            <tr><td><strong>Project</strong></td><td>@_project.ProjectName</td></tr>
                            <tr><td><strong>USP</strong></td><td>@_project.UspNumber</td></tr>
                            <tr><td><strong>Customer</strong></td><td>@_project.CustomerSiteName</td></tr>
                            <tr><td><strong>PM</strong></td><td>@_project.ProjectManager</td></tr>
                            <tr><td><strong>Start</strong></td><td>@_project.ProjectStartDate?.ToString("yyyy-MM-dd")</td></tr>
                            <tr><td><strong>End</strong></td><td>@_project.TargetCompletionDate?.ToString("yyyy-MM-dd")</td></tr>
                            <tr><td><strong>Tasks</strong></td><td>@_project.Tasks.Count(t => !t.IsSummaryTask)</td></tr>
                            <tr><td><strong>Dependencies</strong></td><td>@_project.Tasks.Sum(t => t.Predecessors.Count)</td></tr>
                        </tbody>
                    </MudSimpleTable>
                </MudPaper>
            </MudItem>

            <MudItem xs="12" md="8">
                <MudPaper Class="pa-4" Elevation="1">
                    <MudText Typo="Typo.h6" Class="mb-3">Gantt Preview</MudText>
                    <GanttPreview Tasks="@_project.Tasks"
                                  ProjectStart="@_project.ProjectStartDate"
                                  ProjectEnd="@_project.TargetCompletionDate" />
                </MudPaper>
            </MudItem>
        </MudGrid>

        @* Export Actions *@
        <MudPaper Class="pa-4 mt-4" Elevation="1">
            <MudStack Row="true" Spacing="3" AlignItems="AlignItems.Center">
                <MudButton Variant="Variant.Filled" Color="Color.Primary" Size="Size.Large"
                           StartIcon="@Icons.Material.Filled.FileDownload"
                           OnClick="ExportXml"
                           Disabled="@(_validationErrors.Count > 0)">
                    Generate MSP XML
                </MudButton>
                <MudButton Variant="Variant.Outlined" Color="Color.Default"
                           StartIcon="@Icons.Material.Filled.Save"
                           OnClick="SaveDraft">
                    Save Draft
                </MudButton>
                <MudButton Variant="Variant.Outlined" Color="Color.Secondary"
                           StartIcon="@Icons.Material.Filled.Preview"
                           OnClick="PreviewXml"
                           Disabled="@(_validationErrors.Count > 0)">
                    Preview XML
                </MudButton>
            </MudStack>
        </MudPaper>

        @* XML Preview Dialog *@
        @if (_showXmlPreview)
        {
            <MudPaper Class="pa-4 mt-4" Elevation="1">
                <MudStack Row="true" Justify="Justify.SpaceBetween" AlignItems="AlignItems.Center" Class="mb-2">
                    <MudText Typo="Typo.h6">XML Preview</MudText>
                    <MudIconButton Icon="@Icons.Material.Filled.Close" OnClick="@(() => _showXmlPreview = false)" />
                </MudStack>
                <MudTextField Value="@_xmlPreview" Lines="20" ReadOnly="true"
                              Variant="Variant.Outlined"
                              Style="font-family: 'Consolas', 'Courier New', monospace; font-size: 12px;" />
            </MudPaper>
        }
    }
</MudContainer>

@code {
    private ProjectModel _project = new();
    private List<string> _validationErrors = [];
    private bool _showXmlPreview;
    private string _xmlPreview = string.Empty;

    protected override void OnInitialized()
    {
        _project = StateService.Project;
        Validate();
    }

    private void Validate()
    {
        _validationErrors = XmlService.ValidateForExport(_project);
    }

    private async Task ExportXml()
    {
        Validate();
        if (_validationErrors.Count > 0) return;

        try
        {
            var xml = XmlService.GenerateXml(_project);
            var fileName = XmlExportService.GetDefaultExportFileName(_project);

            await JS.InvokeVoidAsync("interop.downloadFile", fileName, xml, "application/xml");

            Snackbar.Add($"XML exported successfully as {fileName}", Severity.Success);
        }
        catch (Exception ex)
        {
            Snackbar.Add($"Export failed: {ex.Message}", Severity.Error);
        }
    }

    private void PreviewXml()
    {
        Validate();
        if (_validationErrors.Count > 0) return;

        _xmlPreview = XmlService.GenerateXml(_project);
        _showXmlPreview = true;
    }

    private async Task SaveDraft()
    {
        try
        {
            var json = DraftService.SerializeProject(_project);
            var fileName = DraftService.GetDefaultDraftFileName(_project);
            await JS.InvokeVoidAsync("interop.downloadFile", fileName, json, "application/json");
            Snackbar.Add("Draft saved successfully.", Severity.Success);
        }
        catch (Exception ex)
        {
            Snackbar.Add($"Failed to save draft: {ex.Message}", Severity.Error);
        }
    }
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Shared\MainLayout.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Shared\MainLayout.razor" @'
@inherits LayoutComponentBase
@inject ProjectStateService StateService
@inject NavigationManager Navigation

<MudThemeProvider @bind-IsDarkMode="@_isDarkMode" Theme="_theme" />
<MudPopoverProvider />
<MudDialogProvider />
<MudSnackbarProvider />

<MudLayout>
    <MudAppBar Elevation="1" Dense="true">
        <MudIconButton Icon="@Icons.Material.Filled.Menu" Color="Color.Inherit" Edge="Edge.Start"
                       OnClick="@ToggleDrawer" />
        <MudText Typo="Typo.h6" Class="ml-3">Honeywell Experion Migration Scheduler</MudText>
        <MudSpacer />
        <MudText Typo="Typo.body2" Class="mr-4" Style="opacity:0.8">
            @(string.IsNullOrEmpty(StateService.Project.UspNumber) ? "" : StateService.Project.UspNumber)
        </MudText>
        <MudToggleIconButton @bind-Toggled="@_isDarkMode"
                             Icon="@Icons.Material.Filled.DarkMode"
                             ToggledIcon="@Icons.Material.Filled.LightMode"
                             Color="Color.Inherit"
                             ToggledColor="Color.Warning"
                             Title="Toggle Dark/Light Mode" />
    </MudAppBar>

    <MudDrawer @bind-Open="_drawerOpen" ClipMode="DrawerClipMode.Always" Elevation="2" Variant="DrawerVariant.Mini"
               OpenMiniOnHover="true">
        <MudNavMenu>
            <MudNavLink Href="/" Match="NavLinkMatch.All"
                        Icon="@Icons.Material.Filled.Settings"
                        IconColor="@(StateService.Project.SetupComplete ? Color.Success : Color.Default)">
                Project Setup
            </MudNavLink>
            <MudNavLink Href="/tasks"
                        Icon="@Icons.Material.Filled.ListAlt"
                        IconColor="@(StateService.Project.TasksComplete ? Color.Success : Color.Default)"
                        Disabled="@(!StateService.Project.SetupComplete)">
                Activity Editor
            </MudNavLink>
            <MudNavLink Href="/constraints"
                        Icon="@Icons.Material.Filled.AccountTree"
                        IconColor="@(StateService.Project.ConstraintsComplete ? Color.Success : Color.Default)"
                        Disabled="@(!StateService.Project.TasksComplete)">
                Constraints
            </MudNavLink>
            <MudNavLink Href="/export"
                        Icon="@Icons.Material.Filled.FileDownload"
                        Disabled="@(!StateService.Project.ConstraintsComplete)">
                Export
            </MudNavLink>
        </MudNavMenu>
    </MudDrawer>

    <MudMainContent Class="pt-16 px-4">
        @Body
    </MudMainContent>
</MudLayout>

@code {
    private bool _drawerOpen = true;
    private bool _isDarkMode;

    private readonly MudTheme _theme = new()
    {
        PaletteLight = new PaletteLight
        {
            Primary = "#E51937",       // Honeywell Red
            Secondary = "#1A1A1A",
            AppbarBackground = "#E51937",
            AppbarText = "#FFFFFF",
            Background = "#F5F5F5",
            DrawerBackground = "#FFFFFF",
            DrawerText = "#333333"
        },
        PaletteDark = new PaletteDark
        {
            Primary = "#E51937",
            Secondary = "#FF6B6B",
            AppbarBackground = "#1A1A1A",
            AppbarText = "#FFFFFF",
            Background = "#121212",
            Surface = "#1E1E1E",
            DrawerBackground = "#1A1A1A",
            DrawerText = "#E0E0E0"
        },
        Typography = new Typography
        {
            Default = new DefaultTypography
            {
                FontFamily = ["Roboto", "Helvetica", "Arial", "sans-serif"]
            }
        }
    };

    private void ToggleDrawer() => _drawerOpen = !_drawerOpen;
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Components\MigrationTypeSelector.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Components\MigrationTypeSelector.razor" @'
@* Migration Type selection component with prerequisite "pre-existing" toggles *@

<MudGrid>
    @foreach (var type in _selectableTypes)
    {
        <MudItem xs="12" md="6">
            <MudPaper Class="pa-3" Elevation="0" Outlined="true"
                      Style="@(IsSelected(type) ? "border-color: var(--mud-palette-primary); border-width: 2px;" : "")">
                <MudStack Row="true" AlignItems="AlignItems.Center" Spacing="2">
                    <MudCheckBox T="bool" Value="@IsSelected(type)"
                                 ValueChanged="@((bool v) => ToggleType(type, v))"
                                 Color="Color.Primary" Dense="true" />
                    <MudStack Spacing="0" Style="flex:1">
                        <MudText Typo="Typo.subtitle2">@type.GetDisplayName()</MudText>
                        <MudText Typo="Typo.caption" Color="Color.Secondary">@type.GetDescription()</MudText>
                    </MudStack>
                </MudStack>

                @if (HasPreExistingOption(type) && !IsSelected(type))
                {
                    <MudStack Row="true" AlignItems="AlignItems.Center" Class="mt-2 ml-8">
                        <MudSwitch T="bool" Value="@IsPreExisting(type)"
                                   ValueChanged="@((bool v) => TogglePreExisting(type, v))"
                                   Color="Color.Info" Size="Size.Small" />
                        <MudText Typo="Typo.caption" Color="Color.Info">
                            Already Completed (Pre-existing)
                        </MudText>
                    </MudStack>
                }
            </MudPaper>
        </MudItem>
    }
</MudGrid>

@code {
    [Parameter] public HashSet<MigrationType> SelectedTypes { get; set; } = [];
    [Parameter] public HashSet<MigrationType> PreExistingTypes { get; set; } = [];
    [Parameter] public EventCallback OnSelectionChanged { get; set; }

    // Types available for selection (General is always auto-included, not shown here)
    private static readonly MigrationType[] _selectableTypes =
    [
        MigrationType.ELCN,
        MigrationType.EUCN,
        MigrationType.AMT,
        MigrationType.Virtualization,
        MigrationType.SafetyManager,
        MigrationType.UOC,
        MigrationType.HardwareRefresh
    ];

    // Types that can serve as prerequisites and thus have a "pre-existing" toggle
    private static readonly HashSet<MigrationType> _prerequisiteTypes =
    [
        MigrationType.ELCN,
        MigrationType.EUCN,
        MigrationType.HardwareRefresh
    ];

    private bool IsSelected(MigrationType type) => SelectedTypes.Contains(type);
    private bool IsPreExisting(MigrationType type) => PreExistingTypes.Contains(type);
    private bool HasPreExistingOption(MigrationType type) => _prerequisiteTypes.Contains(type);

    private async Task ToggleType(MigrationType type, bool selected)
    {
        if (selected)
        {
            SelectedTypes.Add(type);
            PreExistingTypes.Remove(type); // Can't be both
        }
        else
        {
            SelectedTypes.Remove(type);
        }
        await OnSelectionChanged.InvokeAsync();
    }

    private async Task TogglePreExisting(MigrationType type, bool preExisting)
    {
        if (preExisting)
        {
            PreExistingTypes.Add(type);
            SelectedTypes.Remove(type); // Can't be both
        }
        else
        {
            PreExistingTypes.Remove(type);
        }
        await OnSelectionChanged.InvokeAsync();
    }
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Components\PrerequisiteWarning.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Components\PrerequisiteWarning.razor" @'
@* Displays prerequisite validation errors as an amber warning banner *@

@if (Errors.Count > 0)
{
    <MudAlert Severity="Severity.Warning" Class="mt-4" Dense="true" Variant="Variant.Filled">
        <MudText Typo="Typo.subtitle2" Class="mb-1">Prerequisite Requirements Not Met</MudText>
        <MudList T="string" Dense="true" DisableGutters="true">
            @foreach (var error in Errors)
            {
                <MudListItem T="string" Dense="true" Icon="@Icons.Material.Filled.Warning" IconColor="Color.Inherit">
                    <MudText Typo="Typo.body2">@error</MudText>
                </MudListItem>
            }
        </MudList>
        <MudText Typo="Typo.caption" Class="mt-2">
            Resolve these by selecting the required types or marking them as "Already Completed (Pre-existing)".
        </MudText>
    </MudAlert>
}

@code {
    [Parameter] public List<string> Errors { get; set; } = [];
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Components\TaskGrid.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Components\TaskGrid.razor" @'
@* Reusable read-only task grid component for summary display *@

<MudSimpleTable Dense="true" Bordered="true" Hover="true" Striped="true"
                Style="max-height: 400px; overflow-y: auto;">
    <thead>
        <tr>
            <th>ID</th>
            <th>WBS</th>
            <th>Task Name</th>
            <th>Type</th>
            <th>Duration</th>
            <th>Start</th>
            <th>Finish</th>
            <th>Resource</th>
        </tr>
    </thead>
    <tbody>
        @foreach (var task in Tasks)
        {
            <tr style="@(task.IsSummaryTask ? "font-weight:bold; background-color: rgba(0,0,0,0.04);" : "")">
                <td>@task.TaskId</td>
                <td>@task.WbsCode</td>
                <td style="padding-left: @((task.OutlineLevel - 1) * 20 + 8)px">
                    @task.TaskName
                </td>
                <td>@task.MigrationTypeTag.GetCode()</td>
                <td>@(task.IsSummaryTask ? "" : $"{task.DurationDays}d")</td>
                <td>@task.StartDate?.ToString("yyyy-MM-dd")</td>
                <td>@task.FinishDate?.ToString("yyyy-MM-dd")</td>
                <td>@task.ResourceOwner</td>
            </tr>
        }
    </tbody>
</MudSimpleTable>

@code {
    [Parameter] public List<TaskModel> Tasks { get; set; } = [];
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Components\GanttPreview.razor
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Components\GanttPreview.razor" @'
@* Read-only mini Gantt chart rendered as HTML/SVG for export preview *@

<div style="overflow-x: auto; overflow-y: auto; max-height: 500px;">
    @if (_visibleTasks.Count == 0)
    {
        <MudText Typo="Typo.body2" Color="Color.Secondary" Class="pa-4">
            No tasks with dates to display. Assign start dates to see the Gantt chart.
        </MudText>
    }
    else
    {
        <div style="position: relative; min-width: @(_chartWidth)px; min-height: @(_chartHeight)px;">
            @* Header: month labels *@
            <div style="display:flex; height:30px; border-bottom:1px solid #ccc; position:sticky; top:0; background:var(--mud-palette-surface); z-index:1;">
                @foreach (var month in _months)
                {
                    <div style="position:absolute; left:@(month.OffsetPx)px; width:@(month.WidthPx)px;
                                text-align:center; font-size:11px; font-weight:600; padding-top:6px;
                                border-right:1px solid #ddd; color:var(--mud-palette-text-primary);">
                        @month.Label
                    </div>
                }
            </div>

            @* Task bars *@
            @{ int rowIndex = 0; }
            @foreach (var task in _visibleTasks)
            {
                var top = 30 + rowIndex * _rowHeight;
                <div style="position:absolute; left:0; top:@(top)px; width:100%; height:@(_rowHeight)px;
                            display:flex; align-items:center; border-bottom:1px solid #eee;">
                    @* Task name label *@
                    <div style="width:220px; min-width:220px; padding:0 8px; font-size:11px;
                                overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
                                color:var(--mud-palette-text-primary);"
                         title="@task.TaskName">
                        @task.TaskName
                    </div>
                    @* Bar *@
                    <div style="position:relative; flex:1;">
                        <div style="position:absolute; left:@(task.BarLeftPx)px; width:@(Math.Max(task.BarWidthPx, 4))px;
                                    height:16px; top:4px; border-radius:3px;
                                    background:@(task.Color); opacity:0.85;"
                             title="@task.TaskName: @task.Start?.ToString("MMM dd") — @task.End?.ToString("MMM dd") (@task.Duration days)">
                            @if (task.BarWidthPx > 40)
                            {
                                <span style="font-size:9px; color:#fff; padding-left:4px; line-height:16px;">
                                    @task.Duration d
                                </span>
                            }
                        </div>
                    </div>
                </div>
                rowIndex++;
            }
        </div>
    }
</div>

@code {
    [Parameter] public List<TaskModel> Tasks { get; set; } = [];
    [Parameter] public DateTime? ProjectStart { get; set; }
    [Parameter] public DateTime? ProjectEnd { get; set; }

    private List<GanttTask> _visibleTasks = [];
    private List<MonthHeader> _months = [];
    private int _chartWidth;
    private int _chartHeight;
    private const int _rowHeight = 24;
    private const double _pixelsPerDay = 4.0;

    protected override void OnParametersSet()
    {
        BuildChart();
    }

    private void BuildChart()
    {
        _visibleTasks.Clear();
        _months.Clear();

        if (Tasks.Count == 0) return;

        // Determine date range
        var tasksWithDates = Tasks.Where(t => !t.IsSummaryTask && t.StartDate.HasValue).ToList();
        if (tasksWithDates.Count == 0 && !ProjectStart.HasValue) return;

        var chartStart = ProjectStart ?? tasksWithDates.Min(t => t.StartDate!.Value);
        var chartEnd = ProjectEnd ?? DateTime.Today.AddMonths(6);

        // Extend chartEnd if any tasks go beyond
        foreach (var t in tasksWithDates)
        {
            var taskEnd = t.FinishDate ?? t.StartDate!.Value.AddDays(t.DurationDays);
            if (taskEnd > chartEnd) chartEnd = taskEnd;
        }

        chartStart = new DateTime(chartStart.Year, chartStart.Month, 1);
        chartEnd = new DateTime(chartEnd.Year, chartEnd.Month, 1).AddMonths(1);

        var totalDays = (chartEnd - chartStart).TotalDays;
        _chartWidth = (int)(totalDays * _pixelsPerDay) + 220; // 220 for label column

        // Build month headers
        var current = chartStart;
        while (current < chartEnd)
        {
            var nextMonth = current.AddMonths(1);
            var monthDays = (nextMonth - current).TotalDays;
            _months.Add(new MonthHeader
            {
                Label = current.ToString("MMM yy"),
                OffsetPx = (int)((current - chartStart).TotalDays * _pixelsPerDay) + 220,
                WidthPx = (int)(monthDays * _pixelsPerDay)
            });
            current = nextMonth;
        }

        // Build visible tasks
        var colorMap = new Dictionary<MigrationType, string>
        {
            [MigrationType.General] = "#607D8B",
            [MigrationType.ELCN] = "#2196F3",
            [MigrationType.EUCN] = "#4CAF50",
            [MigrationType.AMT] = "#FF9800",
            [MigrationType.Virtualization] = "#9C27B0",
            [MigrationType.SafetyManager] = "#F44336",
            [MigrationType.UOC] = "#00BCD4",
            [MigrationType.HardwareRefresh] = "#795548"
        };

        foreach (var task in Tasks.Where(t => !t.IsSummaryTask))
        {
            var start = task.StartDate;
            var end = task.FinishDate ?? (start?.AddDays(Math.Max(task.DurationDays - 1, 0)));

            var barLeft = start.HasValue
                ? (int)((start.Value - chartStart).TotalDays * _pixelsPerDay)
                : 0;
            var barWidth = start.HasValue && end.HasValue
                ? (int)((end.Value - start.Value).TotalDays * _pixelsPerDay)
                : (int)(task.DurationDays * _pixelsPerDay);

            _visibleTasks.Add(new GanttTask
            {
                TaskName = task.TaskName,
                Start = start,
                End = end,
                Duration = task.DurationDays,
                BarLeftPx = Math.Max(barLeft, 0),
                BarWidthPx = Math.Max(barWidth, 4),
                Color = colorMap.GetValueOrDefault(task.MigrationTypeTag, "#607D8B")
            });
        }

        _chartHeight = 30 + _visibleTasks.Count * _rowHeight;
    }

    private class GanttTask
    {
        public string TaskName { get; set; } = "";
        public DateTime? Start { get; set; }
        public DateTime? End { get; set; }
        public int Duration { get; set; }
        public int BarLeftPx { get; set; }
        public int BarWidthPx { get; set; }
        public string Color { get; set; } = "#607D8B";
    }

    private class MonthHeader
    {
        public string Label { get; set; } = "";
        public int OffsetPx { get; set; }
        public int WidthPx { get; set; }
    }
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Models\MigrationType.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Models\MigrationType.cs" @'
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
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Models\TaskConstraint.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Models\TaskConstraint.cs" @'
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
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Models\PredecessorLink.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Models\PredecessorLink.cs" @'
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
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Models\TaskModel.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Models\TaskModel.cs" @'
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
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Models\ProjectModel.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Models\ProjectModel.cs" @'
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
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Services\ProjectStateService.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Services\ProjectStateService.cs" @'
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
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Services\PrerequisiteValidator.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Services\PrerequisiteValidator.cs" @'
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
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Services\TaskTemplateService.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Services\TaskTemplateService.cs" @'
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
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Services\DraftService.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Services\DraftService.cs" @'
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
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\Services\XmlExportService.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\Services\XmlExportService.cs" @'
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
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\wwwroot\css\app.css
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\wwwroot\css\app.css" @'
/* Honeywell Experion Migration Scheduler — Application Styles */

html, body {
    font-family: 'Roboto', 'Helvetica', 'Arial', sans-serif;
    margin: 0;
    padding: 0;
    height: 100%;
}

/* Summary row styling in the Activity Editor */
.summary-row {
    font-weight: 600 !important;
    background-color: rgba(0, 0, 0, 0.04) !important;
}

.mud-dark .summary-row {
    background-color: rgba(255, 255, 255, 0.06) !important;
}

/* Stepper styling overrides */
.mud-stepper .mud-step-label-content {
    font-size: 0.875rem;
}

/* Gantt chart scroll container */
.gantt-container {
    overflow-x: auto;
    overflow-y: auto;
    max-height: 500px;
    position: relative;
}

/* Table fixed header alignment */
.mud-table-container {
    scrollbar-width: thin;
}

/* Drawer compact mode */
.mud-nav-link {
    padding: 8px 16px !important;
}

/* Blazor error UI */
#blazor-error-ui {
    background: lightyellow;
    bottom: 0;
    box-shadow: 0 -1px 2px rgba(0, 0, 0, 0.2);
    display: none;
    left: 0;
    padding: 0.6rem 1.25rem 0.7rem 1.25rem;
    position: fixed;
    width: 100%;
    z-index: 1000;
}

#blazor-error-ui .dismiss {
    cursor: pointer;
    position: absolute;
    right: 0.75rem;
    top: 0.5rem;
}

/* Print-friendly styles */
@media print {
    .mud-appbar,
    .mud-drawer,
    .no-print {
        display: none !important;
    }

    .mud-main-content {
        margin: 0 !important;
        padding: 0 !important;
    }
}

/* Responsive adjustments for 1366x768 */
@media (max-width: 1400px) {
    .mud-table th,
    .mud-table td {
        padding: 4px 8px !important;
        font-size: 0.8125rem;
    }

    .mud-expansion-panel-text {
        font-size: 0.875rem;
    }
}

/* Custom scrollbar for WebView2 */
::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

::-webkit-scrollbar-track {
    background: transparent;
}

::-webkit-scrollbar-thumb {
    background: rgba(0, 0, 0, 0.2);
    border-radius: 4px;
}

::-webkit-scrollbar-thumb:hover {
    background: rgba(0, 0, 0, 0.35);
}

.mud-dark ::-webkit-scrollbar-thumb {
    background: rgba(255, 255, 255, 0.2);
}

.mud-dark ::-webkit-scrollbar-thumb:hover {
    background: rgba(255, 255, 255, 0.35);
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Blazor\wwwroot\js\interop.js
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Blazor\wwwroot\js\interop.js" @'
/**
 * JavaScript interop functions for the Migration Scheduler Blazor application.
 * Provides file download and upload capabilities via the browser.
 */
window.interop = {
    /**
     * Downloads a file by creating a Blob and triggering a browser download.
     * @param {string} fileName - The default file name for the download.
     * @param {string} content - The file content as a string.
     * @param {string} mimeType - The MIME type (e.g., "application/xml", "application/json").
     */
    downloadFile: function (fileName, content, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = fileName;
        document.body.appendChild(anchor);
        anchor.click();
        document.body.removeChild(anchor);
        URL.revokeObjectURL(url);
    },

    /**
     * Opens a file picker dialog and returns the file content as a string.
     * @param {string} accept - File type filter (e.g., ".json", ".xml").
     * @returns {Promise<string>} The file content or empty string if cancelled.
     */
    uploadFile: function (accept) {
        return new Promise((resolve) => {
            const input = document.createElement("input");
            input.type = "file";
            input.accept = accept;
            input.style.display = "none";

            input.addEventListener("change", function () {
                if (input.files && input.files.length > 0) {
                    const reader = new FileReader();
                    reader.onload = function () {
                        resolve(reader.result);
                    };
                    reader.onerror = function () {
                        resolve("");
                    };
                    reader.readAsText(input.files[0]);
                } else {
                    resolve("");
                }
                document.body.removeChild(input);
            });

            input.addEventListener("cancel", function () {
                resolve("");
                document.body.removeChild(input);
            });

            document.body.appendChild(input);
            input.click();
        });
    }
};
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Host\MigrationScheduler.Host.csproj
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Host\MigrationScheduler.Host.csproj" @'
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net8.0-windows</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <UseWindowsForms>true</UseWindowsForms>
    <LangVersion>12.0</LangVersion>
  </PropertyGroup>

  <ItemGroup>
    <FrameworkReference Include="Microsoft.AspNetCore.App" />
    <PackageReference Include="Microsoft.Web.WebView2" Version="1.0.2365.46" />
    <ProjectReference Include="..\MigrationScheduler.Blazor\MigrationScheduler.Blazor.csproj" />
  </ItemGroup>

</Project>
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Host\Program.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Host\Program.cs" @'
using System;
using System.Windows.Forms;

namespace MigrationScheduler.Host;

static class Program
{
    [STAThread]
    static void Main()
    {
        Application.SetHighDpiMode(HighDpiMode.SystemAware);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}
'@

# ============================================================================
# MigrationScheduler\MigrationScheduler.Host\MainForm.cs
# ============================================================================
Write-ProjectFile "MigrationScheduler\MigrationScheduler.Host\MainForm.cs" @'
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace MigrationScheduler.Host;

/// <summary>
/// WinForms shell that hosts the Blazor Server application via WebView2.
/// </summary>
public class MainForm : Form
{
    private readonly WebView2 _webView;
    private readonly CancellationTokenSource _cts = new();
    private Thread? _blazorThread;

    public MainForm()
    {
        Text = "Honeywell Experion Migration Scheduler";
        Width = 1920;
        Height = 1080;
        MinimumSize = new System.Drawing.Size(1366, 768);
        StartPosition = FormStartPosition.CenterScreen;
        WindowState = FormWindowState.Maximized;

        _webView = new WebView2
        {
            Dock = DockStyle.Fill
        };
        Controls.Add(_webView);

        Load += MainForm_Load;
        FormClosing += MainForm_FormClosing;
    }

    private async void MainForm_Load(object? sender, EventArgs e)
    {
        // Start Blazor Server on a background thread
        var blazorReady = new TaskCompletionSource();

        _blazorThread = new Thread(() =>
        {
            try
            {
                var builder = Microsoft.AspNetCore.Builder.WebApplication.CreateBuilder();
                builder.Services.AddRazorPages();
                builder.Services.AddServerSideBlazor();
                builder.Services.AddMudServices();

                // Register application services
                builder.Services.AddSingleton<MigrationScheduler.Blazor.Services.ProjectStateService>();
                builder.Services.AddScoped<MigrationScheduler.Blazor.Services.PrerequisiteValidator>();
                builder.Services.AddScoped<MigrationScheduler.Blazor.Services.TaskTemplateService>();
                builder.Services.AddScoped<MigrationScheduler.Blazor.Services.XmlExportService>();
                builder.Services.AddScoped<MigrationScheduler.Blazor.Services.DraftService>();

                builder.WebHost.UseUrls("http://localhost:5199");

                var app = builder.Build();
                app.UseStaticFiles();
                app.UseRouting();
                app.MapBlazorHub();
                app.MapFallbackToPage("/_Host");

                blazorReady.SetResult();
                app.Run();
            }
            catch (Exception ex)
            {
                blazorReady.TrySetException(ex);
            }
        })
        {
            IsBackground = true,
            Name = "BlazorServerThread"
        };

        _blazorThread.Start();

        try
        {
            await blazorReady.Task;
            // Give the server a moment to fully start
            await Task.Delay(1000);

            // Initialize WebView2 and navigate to the Blazor app
            var env = await CoreWebView2Environment.CreateAsync();
            await _webView.EnsureCoreWebView2Async(env);
            _webView.Source = new Uri("http://localhost:5199");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to start the application:\n\n{ex.Message}\n\nPlease ensure WebView2 Runtime is installed.",
                "Startup Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        _cts.Cancel();
        _webView.Dispose();
    }
}

// MudBlazor service registration extension (needed since the Host project references it)
internal static class ServiceExtensions
{
    public static IServiceCollection AddMudServices(this IServiceCollection services)
    {
        return MudBlazor.Services.ServiceCollectionExtensions.AddMudServices(services);
    }
}
'@

# ============================================================================
# MigrationScheduler\README.md
# ============================================================================
Write-ProjectFile "MigrationScheduler\README.md" @'
# Honeywell Experion Migration Project Scheduler

A Blazor Server desktop application for configuring Honeywell Experion migration schedules and exporting Microsoft Project-compatible XML files.

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- [WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) (required for the desktop host; pre-installed on Windows 10/11 with Edge)
- Windows 10 or later (for the WinForms host)

## Project Structure

```
MigrationScheduler/
├── MigrationScheduler.sln                   # Solution file
├── MigrationScheduler.Host/                 # WinForms + WebView2 desktop shell
│   ├── Program.cs
│   ├── MainForm.cs
│   └── MigrationScheduler.Host.csproj
├── MigrationScheduler.Blazor/               # Blazor Server application
│   ├── Program.cs
│   ├── App.razor
│   ├── Pages/
│   │   ├── ProjectSetup.razor               # Step 1: Project setup wizard
│   │   ├── ActivityEditor.razor             # Step 2: Task editor grid
│   │   ├── ConstraintManager.razor          # Step 3: Dependencies & constraints
│   │   └── ExportPage.razor                 # Step 4: XML export & Gantt preview
│   ├── Components/
│   │   ├── TaskGrid.razor                   # Reusable task table
│   │   ├── PrerequisiteWarning.razor        # Prerequisite error banner
│   │   ├── GanttPreview.razor               # Read-only HTML Gantt chart
│   │   └── MigrationTypeSelector.razor      # Migration type checkboxes
│   ├── Models/
│   │   ├── ProjectModel.cs                  # Root project model
│   │   ├── TaskModel.cs                     # Task/activity model
│   │   ├── PredecessorLink.cs               # Dependency link model
│   │   ├── TaskConstraint.cs                # Constraint type enum
│   │   └── MigrationType.cs                 # Migration category enum
│   ├── Services/
│   │   ├── ProjectStateService.cs           # Singleton state management
│   │   ├── XmlExportService.cs              # MSP XML generation
│   │   ├── DraftService.cs                  # JSON save/load
│   │   ├── PrerequisiteValidator.cs         # Prerequisite rules engine
│   │   └── TaskTemplateService.cs           # Default task templates
│   ├── wwwroot/
│   │   ├── css/app.css
│   │   └── js/interop.js                   # JS interop for file save/load
│   └── MigrationScheduler.Blazor.csproj
└── README.md
```

## Build & Run

### Option 1: Blazor Server (Browser)

Run the Blazor application directly in a browser — no WebView2 required:

```bash
cd MigrationScheduler
dotnet restore
dotnet build
dotnet run --project MigrationScheduler.Blazor
```

Then open `http://localhost:5199` in your browser.

### Option 2: Desktop Application (WebView2)

Run the WinForms host which embeds the Blazor app in a native window:

```bash
cd MigrationScheduler
dotnet restore
dotnet build
dotnet run --project MigrationScheduler.Host
```

> **Note:** The Host project requires Windows and the WebView2 Runtime.

### Publish as a Self-Contained Executable

```bash
dotnet publish MigrationScheduler.Host -c Release -r win-x64 --self-contained
```

The output will be in `MigrationScheduler.Host/bin/Release/net8.0-windows/win-x64/publish/`.

## Usage Workflow

1. **Project Setup** — Enter project details (name, USP number, dates, PM), select migration types, and validate prerequisites.
2. **Activity Editor** — Review and edit the auto-generated task list. Add, delete, or reorder tasks.
3. **Constraints** — Define predecessor relationships (FS, SS, FF, SF with lag) and task constraints (ASAP, MSO, SNET, etc.).
4. **Export** — Preview the Gantt chart, validate the schedule, and export as MSP-compatible XML.

## Importing into Microsoft Project

1. Open Microsoft Project (2007 or later).
2. Go to **File > Open**.
3. Change the file type filter to **XML Format (*.xml)**.
4. Select the exported `{USP}_{ProjectName}_Schedule.xml` file.
5. Microsoft Project will import all tasks, dependencies, constraints, resources, and assignments.

## Migration Types Supported

| Code | Full Name | Description |
|------|-----------|-------------|
| ELCN | Enhanced Local Control Network | LCN to ELCN / FTE network migration |
| EUCN | Enhanced Universal Control Network | xPM to EHPM migration via NIM to ENIM/ENB |
| AMT | Advanced Migration Technology | ControlEdge PLC / RTU migrations |
| Virtualization | Experion Virtualization | Physical to VMware vSphere migration |
| Safety Manager | Safety Manager / TCMI | Triconex Communication Module Interface upgrade |
| UOC | Universal Operations Controller | Migration to ControlEdge UOC |
| Hardware Refresh | Hardware Refresh | Server/workstation/switch replacement |

## Prerequisite Rules

- **EUCN** requires ELCN
- **AMT** requires EUCN
- **Safety Manager** requires both EUCN and ELCN
- **UOC** requires ELCN
- **Virtualization** requires Hardware Refresh

Each prerequisite can be bypassed by marking it as "Already Completed (Pre-existing)" for brownfield projects.

## Save & Resume

Use the **Save Draft** button to export the current project state as a JSON file. Use **Load Draft** on the Setup page to resume a previous session.

## MSP XML Compatibility

The generated XML conforms to the Microsoft Project XML Schema (`xmlns="http://schemas.microsoft.com/project"`) and is compatible with:

- Microsoft Project 2007, 2010, 2013, 2016, 2019, 2021, and Project for Microsoft 365
- Duration format: ISO 8601 (`PT{hours}H0M0S`)
- Predecessor lag: stored as `days * 4800` (tenths of minutes)
- Standard calendar: 5-day work week, Mon-Fri, 08:00-17:00
'@

# ============================================================================
# MigrationScheduler\.gitignore
# ============================================================================
Write-ProjectFile "MigrationScheduler\.gitignore" @'
## .NET
bin/
obj/
*.user
*.suo
*.cache
*.vs/
.vscode/

## Build results
[Dd]ebug/
[Rr]elease/
x64/
x86/
[Bb]uild/
bld/

## NuGet
*.nupkg
**/packages/*
project.lock.json
project.fragment.lock.json
artifacts/

## IDE
.idea/
*.sln.docstates
*.userprefs

## OS
.DS_Store
Thumbs.db

## Publish
publish/
'@

Write-Host ""
Write-Host "=============================================="
Write-Host " All files created successfully!"
Write-Host " Next steps:"
Write-Host "   cd MigrationScheduler"
Write-Host "   dotnet restore"
Write-Host "   dotnet build"
Write-Host "   dotnet run --project MigrationScheduler.Blazor"
Write-Host "=============================================="
Write-Host ""
