using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using SharePointToFoundrySync;

var builder = FunctionsApplication.CreateBuilder(args);

builder.Services
    .AddApplicationInsightsTelemetryWorkerService()
    .ConfigureFunctionsApplicationInsights();

// App settings use "__" as the section separator, e.g. SharePoint__ClientId.
builder.Services.Configure<SharePointOptions>(builder.Configuration.GetSection(SharePointOptions.Section));
builder.Services.Configure<SearchOptions>(builder.Configuration.GetSection(SearchOptions.Section));
builder.Services.Configure<SyncOptions>(builder.Configuration.GetSection(SyncOptions.Section));

// Knowledge-source uploads are processed synchronously and can take up to 180 s,
// so the HTTP timeout has to be comfortably above that.
builder.Services.AddHttpClient<SharePointClient>(c => c.Timeout = TimeSpan.FromMinutes(5));
builder.Services.AddHttpClient<KnowledgeSourceClient>(c => c.Timeout = TimeSpan.FromMinutes(5));
builder.Services.AddTransient<SyncService>();

builder.Build().Run();
