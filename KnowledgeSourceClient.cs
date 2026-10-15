using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace SharePointToFoundrySync;

public sealed record KnowledgeSourceFile(
    string FileId,
    string FileName,
    IReadOnlyDictionary<string, string> Metadata,
    DateTimeOffset? LastUpdatedAt,
    string? ErrorMessage);

/// <summary>
/// File operations on a Foundry IQ "file" knowledge source (Azure AI Search REST API, 2026-08-01-preview).
/// Uploads are synchronous: when a call returns, the file has been extracted, chunked, embedded and indexed.
/// </summary>
public sealed class KnowledgeSourceClient
{
    private readonly HttpClient _http;
    private readonly SearchOptions _options;
    private readonly ILogger<KnowledgeSourceClient> _log;
    private readonly string _endpoint;
    private readonly string _ks;

    public KnowledgeSourceClient(HttpClient http, IOptions<SearchOptions> options, ILogger<KnowledgeSourceClient> log)
    {
        _http = http;
        _options = options.Value;
        _options.Validate();
        _log = log;
        _endpoint = _options.Endpoint.TrimEnd('/');
        _ks = Uri.EscapeDataString(_options.KnowledgeSourceName);
    }

    private string Api => $"api-version={_options.ApiVersion}";

    public async Task<List<KnowledgeSourceFile>> ListFilesAsync(CancellationToken ct)
    {
        var files = new List<KnowledgeSourceFile>();
        string? next = $"{_endpoint}/knowledgesources/{_ks}/files?{Api}";

        while (next is not null)
        {
            var url = next;
            using var response = await _http.SendWithRetryAsync(() => Task.FromResult(Request(HttpMethod.Get, url)), _log, ct);
            await response.EnsureSuccessAsync("List knowledge source files", ct);

            await using var stream = await response.Content.ReadAsStreamAsync(ct);
            using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct);

            foreach (var f in doc.RootElement.GetProperty("value").EnumerateArray())
            {
                var metadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (f.TryGetProperty("metadata", out var md) && md.ValueKind == JsonValueKind.Object)
                {
                    foreach (var p in md.EnumerateObject())
                        metadata[p.Name] = p.Value.ValueKind == JsonValueKind.String ? p.Value.GetString()! : p.Value.GetRawText();
                }

                files.Add(new KnowledgeSourceFile(
                    FileId: f.GetProperty("fileId").GetString()!,
                    FileName: f.TryGetProperty("fileName", out var n) ? n.GetString() ?? "" : "",
                    Metadata: metadata,
                    LastUpdatedAt: f.TryGetProperty("lastUpdatedAt", out var lu) && lu.ValueKind == JsonValueKind.String ? lu.GetDateTimeOffset() : null,
                    ErrorMessage: f.TryGetProperty("errorMessage", out var e) && e.ValueKind == JsonValueKind.String ? e.GetString() : null));
            }

            // Follow nextLink exactly as returned (the service rejects edited continuation URLs).
            next = doc.RootElement.TryGetProperty("@odata.nextLink", out var nl) ? nl.GetString() : null;
        }

        return files;
    }

    /// <summary>Uploads a new file. Note: uploading never replaces; the same name would create a duplicate.</summary>
    public async Task<string> UploadAsync(string fileName, byte[] content, IDictionary<string, string> metadata, CancellationToken ct)
    {
        var url = $"{_endpoint}/knowledgesources('{_ks}')/files?{Api}";
        using var response = await _http.SendWithRetryAsync(
            () => Task.FromResult(Request(HttpMethod.Post, url, BuildMultipart(fileName, content, metadata))), _log, ct);
        await response.EnsureSuccessAsync($"Upload '{fileName}'", ct);
        return await ReadFileIdAsync(response, ct);
    }

    /// <summary>Replaces the content (and metadata) of an existing file in place, keeping its fileId.</summary>
    public async Task UpdateAsync(string fileId, string fileName, byte[] content, IDictionary<string, string> metadata, CancellationToken ct)
    {
        var url = $"{_endpoint}/knowledgesources('{_ks}')/files('{Uri.EscapeDataString(fileId)}')?{Api}";
        using var response = await _http.SendWithRetryAsync(
            () => Task.FromResult(Request(HttpMethod.Put, url, BuildMultipart(fileName, content, metadata))), _log, ct);
        await response.EnsureSuccessAsync($"Update '{fileName}' ({fileId})", ct);
    }

    public async Task DeleteAsync(string fileId, CancellationToken ct)
    {
        var url = $"{_endpoint}/knowledgesources/{_ks}/files/{Uri.EscapeDataString(fileId)}?{Api}";
        using var response = await _http.SendWithRetryAsync(() => Task.FromResult(Request(HttpMethod.Delete, url)), _log, ct);
        if (response.StatusCode == System.Net.HttpStatusCode.NotFound) return; // already gone
        await response.EnsureSuccessAsync($"Delete file {fileId}", ct);
    }

    private HttpRequestMessage Request(HttpMethod method, string url, HttpContent? content = null)
    {
        var request = new HttpRequestMessage(method, url) { Content = content };
        request.Headers.Add("api-key", _options.ApiKey);
        return request;
    }

    /// <summary>
    /// multipart/form-data with a JSON "metadata" part and a binary "content" part.
    /// metadata.fileName takes precedence over the part's filename, so the full relative path
    /// (which may contain non-ASCII characters) goes in the JSON, and the header gets a safe ASCII name.
    /// </summary>
    private static MultipartFormDataContent BuildMultipart(string fileName, byte[] bytes, IDictionary<string, string> metadata)
    {
        var boundary = "ks-" + Guid.NewGuid().ToString("N");
        var multipart = new MultipartFormDataContent(boundary);
        // Send the boundary unquoted; some services are strict about this.
        multipart.Headers.ContentType = MediaTypeHeaderValue.Parse($"multipart/form-data; boundary={boundary}");

        var json = JsonSerializer.Serialize(new { fileName, metadata });
        var metadataPart = new StringContent(json, Encoding.UTF8, "application/json");
        metadataPart.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data") { Name = "\"metadata\"" };
        multipart.Add(metadataPart);

        var filePart = new ByteArrayContent(bytes);
        filePart.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
        filePart.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data")
        {
            Name = "\"content\"",
            FileName = "\"" + AsciiLeafName(fileName) + "\""
        };
        multipart.Add(filePart);

        return multipart;
    }

    private static string AsciiLeafName(string path)
    {
        var leaf = path[(path.LastIndexOf('/') + 1)..];
        var sb = new StringBuilder(leaf.Length);
        foreach (var ch in leaf)
            sb.Append(ch is >= ' ' and <= '~' && ch != '"' && ch != '\\' ? ch : '_');
        return sb.ToString();
    }

    private static async Task<string> ReadFileIdAsync(HttpResponseMessage response, CancellationToken ct)
    {
        var body = await response.Content.ReadAsStringAsync(ct);
        if (string.IsNullOrWhiteSpace(body)) return "";
        using var doc = JsonDocument.Parse(body);
        return doc.RootElement.TryGetProperty("fileId", out var id) ? id.GetString() ?? "" : "";
    }
}
