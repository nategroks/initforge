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
