using System.Net.Http.Headers;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace SharePointToFoundrySync;

/// <param name="ItemId">Graph driveItem id (stable across renames and moves).</param>
/// <param name="RelativePath">Path relative to the synced folder, e.g. "HR/handbook.pdf". Used as the knowledge-source file name.</param>
/// <param name="CTag">Content tag; it changes only when the file's bytes change (not on metadata-only edits).</param>
public sealed record SharePointFile(string ItemId, string RelativePath, string CTag, long Size, DateTimeOffset LastModified);

/// <summary>Reads files from a SharePoint document library through Microsoft Graph with app-only (client secret) auth.</summary>
public sealed class SharePointClient
{
    private const string Graph = "https://graph.microsoft.com/v1.0";
    private static readonly TokenRequestContext GraphScope = new(["https://graph.microsoft.com/.default"]);
    private const string ItemSelect = "$select=id,name,size,cTag,lastModifiedDateTime,file,folder&$top=999";

    private readonly HttpClient _http;
    private readonly SharePointOptions _options;
    private readonly ILogger<SharePointClient> _log;
    private readonly TokenCredential _credential;
    private string? _driveId;

    public SharePointClient(HttpClient http, IOptions<SharePointOptions> options, ILogger<SharePointClient> log)
    {
        _http = http;
        _options = options.Value;
        _options.Validate();
        _log = log;
        _credential = new ClientSecretCredential(_options.TenantId, _options.ClientId, _options.ClientSecret);
    }

    /// <summary>Lists every file (recursively) under the configured library/folder.</summary>
    public async Task<List<SharePointFile>> ListFilesAsync(CancellationToken ct)
    {
        var driveId = await GetDriveIdAsync(ct);
        var folder = _options.FolderPath.Trim('/');

        var startUrl = folder.Length == 0
            ? $"{Graph}/drives/{driveId}/root/children?{ItemSelect}"
            : $"{Graph}/drives/{driveId}/root:/{EscapePath(folder)}:/children?{ItemSelect}";

        var files = new List<SharePointFile>();
        var pending = new Queue<(string Url, string Prefix)>();
        pending.Enqueue((startUrl, ""));

        while (pending.Count > 0)
        {
            var (url, prefix) = pending.Dequeue();
            string? next = url;
            while (next is not null)
            {
                using var doc = await GetJsonAsync(next, "List SharePoint folder", ct);
                foreach (var item in doc.RootElement.GetProperty("value").EnumerateArray())
                {
                    var name = item.GetProperty("name").GetString()!;
                    var id = item.GetProperty("id").GetString()!;

                    if (item.TryGetProperty("folder", out _))
                    {
                        pending.Enqueue(($"{Graph}/drives/{driveId}/items/{id}/children?{ItemSelect}", prefix + name + "/"));
                    }
                    else if (item.TryGetProperty("file", out _))
                    {
                        files.Add(new SharePointFile(
                            ItemId: id,
                            RelativePath: prefix + name,
                            CTag: item.TryGetProperty("cTag", out var c) ? c.GetString() ?? "" : "",
                            Size: item.TryGetProperty("size", out var s) ? s.GetInt64() : 0,
                            LastModified: item.TryGetProperty("lastModifiedDateTime", out var m) ? m.GetDateTimeOffset() : DateTimeOffset.MinValue));
                    }
                }

                next = doc.RootElement.TryGetProperty("@odata.nextLink", out var nl) ? nl.GetString() : null;
            }
        }

        return files;
    }

    public async Task<byte[]> DownloadAsync(string itemId, CancellationToken ct)
    {
        var driveId = await GetDriveIdAsync(ct);
        // Graph answers with a 302 to a pre-authenticated URL. HttpClient follows it and drops the
        // Authorization header on the redirect, which is what that URL expects.
        using var response = await _http.SendWithRetryAsync(
            () => AuthorizedRequestAsync(HttpMethod.Get, $"{Graph}/drives/{driveId}/items/{itemId}/content", ct),
            _log, ct);
        await response.EnsureSuccessAsync($"Download SharePoint item {itemId}", ct);
        return await response.Content.ReadAsByteArrayAsync(ct);
    }

    private async Task<string> GetDriveIdAsync(CancellationToken ct)
    {
        if (_driveId is not null) return _driveId;

        var sitePath = _options.SitePath.Trim('/');
        var siteUrl = sitePath.Length == 0
            ? $"{Graph}/sites/{_options.Hostname}?$select=id"
            : $"{Graph}/sites/{_options.Hostname}:/{EscapePath(sitePath)}?$select=id";

        string siteId;
        using (var site = await GetJsonAsync(siteUrl, "Resolve SharePoint site", ct))
            siteId = site.RootElement.GetProperty("id").GetString()!;

        using var drives = await GetJsonAsync($"{Graph}/sites/{siteId}/drives?$select=id,name,webUrl", "List document libraries", ct);
        var available = new List<string>();
        foreach (var d in drives.RootElement.GetProperty("value").EnumerateArray())
        {
            var name = d.GetProperty("name").GetString() ?? "";
            var webUrl = d.TryGetProperty("webUrl", out var w) ? w.GetString() ?? "" : "";
            // Match either the display name ("Documents") or the URL segment ("Shared Documents").
            var urlSegment = Uri.UnescapeDataString(webUrl.TrimEnd('/').Split('/').LastOrDefault() ?? "");
            available.Add(name);

            if (string.Equals(name, _options.LibraryName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(urlSegment, _options.LibraryName, StringComparison.OrdinalIgnoreCase))
            {
                _driveId = d.GetProperty("id").GetString()!;
                _log.LogInformation("Using document library '{Library}' (drive {DriveId}).", name, _driveId);
                return _driveId;
            }
        }

        throw new InvalidOperationException(
            $"Document library '{_options.LibraryName}' not found on site. Available: {string.Join(", ", available)}");
    }

    private async Task<JsonDocument> GetJsonAsync(string url, string operation, CancellationToken ct)
    {
        using var response = await _http.SendWithRetryAsync(
            () => AuthorizedRequestAsync(HttpMethod.Get, url, ct), _log, ct);
        await response.EnsureSuccessAsync(operation, ct);
        await using var stream = await response.Content.ReadAsStreamAsync(ct);
        return await JsonDocument.ParseAsync(stream, cancellationToken: ct);
    }

    private async Task<HttpRequestMessage> AuthorizedRequestAsync(HttpMethod method, string url, CancellationToken ct)
    {
        // ClientSecretCredential caches the token in memory and only refreshes it when it's close to expiring.
        var token = await _credential.GetTokenAsync(GraphScope, ct);
        var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        return request;
    }

    private static string EscapePath(string path) =>
        string.Join('/', path.Split('/', StringSplitOptions.RemoveEmptyEntries).Select(Uri.EscapeDataString));
}
