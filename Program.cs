using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using SharePointToFoundrySync;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices((context, services) =>
    {
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        // App settings use "__" as the section separator, e.g. SharePoint__ClientId.
        var config = context.Configuration;
        services.Configure<SharePointOptions>(config.GetSection(SharePointOptions.Section));
        services.Configure<SearchOptions>(config.GetSection(SearchOptions.Section));
        services.Configure<SyncOptions>(config.GetSection(SyncOptions.Section));

        // Knowledge-source uploads are processed synchronously and can take up to 180 s,
        // so the HTTP timeout has to be comfortably above that.
        services.AddHttpClient<SharePointClient>(c => c.Timeout = TimeSpan.FromMinutes(5));
        services.AddHttpClient<KnowledgeSourceClient>(c => c.Timeout = TimeSpan.FromMinutes(5));
        services.AddTransient<SyncService>();
    })
    .Build();

host.Run();
