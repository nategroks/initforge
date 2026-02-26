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
