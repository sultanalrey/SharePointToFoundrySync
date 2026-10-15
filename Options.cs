namespace SharePointToFoundrySync;

/// <summary>Entra ID app registration + the SharePoint library to read from.</summary>
public sealed class SharePointOptions
{
    public const string Section = "SharePoint";

    public string TenantId { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string ClientSecret { get; set; } = "";

    /// <summary>e.g. contoso.sharepoint.com</summary>
    public string Hostname { get; set; } = "";

    /// <summary>Server-relative site path, e.g. sites/AIKnowledge. Leave empty for the root site.</summary>
    public string SitePath { get; set; } = "";

    /// <summary>Library name as Graph reports it. The default "Shared Documents" library is called "Documents".</summary>
    public string LibraryName { get; set; } = "Documents";

    /// <summary>Optional folder inside the library to sync (e.g. "Policies/Current"). Empty = whole library.</summary>
    public string FolderPath { get; set; } = "";

    public void Validate()
    {
        Require(TenantId, nameof(TenantId));
        Require(ClientId, nameof(ClientId));
        Require(ClientSecret, nameof(ClientSecret));
        Require(Hostname, nameof(Hostname));
        Require(LibraryName, nameof(LibraryName));
    }

    private static void Require(string value, string name)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new InvalidOperationException($"App setting {Section}__{name} is required.");
    }
}

/// <summary>
/// Foundry IQ knowledge bases are stored on an Azure AI Search service. The key here is that
/// search service's ADMIN key (query keys are read-only and can't upload files).
/// </summary>
public sealed class SearchOptions
{
    public const string Section = "Search";

    /// <summary>e.g. https://my-search.search.windows.net</summary>
    public string Endpoint { get; set; } = "";
    public string ApiKey { get; set; } = "";

    /// <summary>Name of the file knowledge source (kind = "file") that your knowledge base references.</summary>
    public string KnowledgeSourceName { get; set; } = "";

    /// <summary>2026-08-01-preview is the first version with in-place file update and file metadata.</summary>
    public string ApiVersion { get; set; } = "2026-08-01-preview";

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(Endpoint)) throw new InvalidOperationException($"App setting {Section}__Endpoint is required.");
        if (string.IsNullOrWhiteSpace(ApiKey)) throw new InvalidOperationException($"App setting {Section}__ApiKey is required.");
        if (string.IsNullOrWhiteSpace(KnowledgeSourceName)) throw new InvalidOperationException($"App setting {Section}__KnowledgeSourceName is required.");
    }
}

public sealed class SyncOptions
{
    public const string Section = "Sync";

    /// <summary>
    /// Delete knowledge-source files whose SharePoint original was deleted or renamed. Only files
    /// this function uploaded (tagged source=sharepoint-sync) are ever deleted.
    /// </summary>
    public bool DeleteOrphans { get; set; } = true;

    /// <summary>Parallel uploads. The service returns 429 when its processing queue is full, so keep this modest.</summary>
    public int MaxConcurrency { get; set; } = 4;

    /// <summary>50 MB on Free/Basic search tiers, 100 MB on other tiers (2026-08-01-preview).</summary>
    public int MaxFileSizeMB { get; set; } = 100;

    /// <summary>Service limit per file knowledge source in 2026-08-01-preview.</summary>
    public int MaxFilesPerKnowledgeSource { get; set; } = 200;
}
