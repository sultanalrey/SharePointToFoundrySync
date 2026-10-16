using System.Diagnostics;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace SharePointToFoundrySync;

public sealed record SyncResult(int Uploaded, int Updated, int Unchanged, int Deleted, int Skipped, int Failed, TimeSpan Elapsed);

/// <summary>
/// Mirrors a SharePoint library into a file knowledge source:
///   new in SharePoint     -> upload
///   changed in SharePoint -> update in place (same fileId, content replaced)
///   unchanged             -> skip (compared by SharePoint cTag stored as file metadata)
///   gone from SharePoint  -> delete (only files this function created)
/// </summary>
public sealed class SyncService
{
    public const string SourceTag = "sharepoint-sync";

    // Types a file knowledge source accepts with "minimal" extraction. Images need "standard"
    // extraction, so they're skipped instead of failing with 415.
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pdf", ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx",
        ".json", ".txt", ".md", ".markdown", ".html", ".htm", ".csv", ".sh"
    };

    private readonly SharePointClient sharePoint;
    private readonly KnowledgeSourceClient knowledgeSource;
    private readonly SyncOptions _options;
    private readonly ILogger<SyncService> log;

    public SyncService(
        SharePointClient sharePoint,
        KnowledgeSourceClient knowledgeSource,
        IOptions<SyncOptions> options,
        ILogger<SyncService> log)
    {
        this.sharePoint = sharePoint;
        this.knowledgeSource = knowledgeSource;
        _options = options.Value;
        this.log = log;
    }

    public async Task<SyncResult> RunAsync(CancellationToken ct)
    {
        var sw = Stopwatch.StartNew();
        int uploaded = 0, updated = 0, unchanged = 0, deleted = 0, skipped = 0, failed = 0;
        var maxBytes = (long)_options.MaxFileSizeMB * 1024 * 1024;

        // 1. What should be there.
        var all = await sharePoint.ListFilesAsync(ct);
        var desired = new List<SharePointFile>();
        foreach (var f in all)
        {
            if (!SupportedExtensions.Contains(Path.GetExtension(f.RelativePath)))
            {
                log.LogDebug("Skipping unsupported type: {File}", f.RelativePath);
                skipped++;
            }
            else if (f.Size == 0)
            {
                log.LogInformation("Skipping empty file: {File}", f.RelativePath);
                skipped++;
            }
            else if (f.Size > maxBytes)
            {
                log.LogWarning("Skipping {File}: {SizeMB:F1} MB exceeds the {Max} MB limit.", f.RelativePath, f.Size / 1048576.0, _options.MaxFileSizeMB);
                skipped++;
            }
            else desired.Add(f);
        }

        if (desired.Count > _options.MaxFilesPerKnowledgeSource)
            log.LogWarning("SharePoint has {Count} eligible files but a file knowledge source holds at most {Max}; uploads past the limit will fail with 409. Consider a blob knowledge source.",
                desired.Count, _options.MaxFilesPerKnowledgeSource);

        // 2. What is there.
        var existing = await knowledgeSource.ListFilesAsync(ct);
        var existingByName = existing
            .GroupBy(f => f.FileName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.OrderByDescending(f => f.LastUpdatedAt ?? DateTimeOffset.MinValue).ToList(),
                StringComparer.OrdinalIgnoreCase);

        log.LogInformation("SharePoint: {Desired} eligible files ({Skipped} skipped). Knowledge source: {Existing} files.",
            desired.Count, skipped, existing.Count);

        // 3. Remove orphans first, which frees capacity before uploading.
        if (_options.DeleteOrphans)
        {
            var desiredNames = desired.Select(f => f.RelativePath).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var orphans = existing
                .Where(f => !desiredNames.Contains(f.FileName) && IsManagedByUs(f))
                .ToList();

            if (desired.Count == 0 && orphans.Count > 0)
            {
                // Safety net: an empty listing is far more likely a permissions/config problem than a
                // genuinely emptied library. Don't wipe the knowledge base on that signal alone.
                log.LogWarning("SharePoint returned no eligible files; NOT deleting {Count} existing knowledge-source files.", orphans.Count);
            }
            else
            {
                foreach (var orphan in orphans)
                {
                    try
                    {
                        await knowledgeSource.DeleteAsync(orphan.FileId, ct);
                        deleted++;
                        log.LogInformation("Deleted {File} (no longer in SharePoint).", orphan.FileName);
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException)
                    {
                        failed++;
                        log.LogError(ex, "Failed to delete orphan {File} ({FileId}).", orphan.FileName, orphan.FileId);
                    }
                }
            }
        }

        // 4. Upload new / replace changed.
        await Parallel.ForEachAsync(
            desired,
            new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, _options.MaxConcurrency), CancellationToken = ct },
            async (file, token) =>
            {
                try
                {
                    var metadata = new Dictionary<string, string>
                    {
                        ["source"] = SourceTag,
                        ["spItemId"] = file.ItemId,
                        ["spCTag"] = file.CTag,
                        ["spLastModified"] = file.LastModified.ToString("o")
                    };

                    if (!existingByName.TryGetValue(file.RelativePath, out var matches))
                    {
                        var bytes = await sharePoint.DownloadAsync(file.ItemId, token);
                        await knowledgeSource.UploadAsync(file.RelativePath, bytes, metadata, token);
                        Interlocked.Increment(ref uploaded);
                        log.LogInformation("Uploaded {File}.", file.RelativePath);
                        return;
                    }

                    // Uploading the same name twice creates duplicates. Keep the newest and remove the rest.
                    var primary = matches[0];
                    foreach (var dup in matches.Skip(1))
                    {
                        await knowledgeSource.DeleteAsync(dup.FileId, token);
                        Interlocked.Increment(ref deleted);
                        log.LogInformation("Removed duplicate {File} ({FileId}).", dup.FileName, dup.FileId);
                    }

                    var isCurrent = primary.ErrorMessage is null
                                    && file.CTag.Length > 0
                                    && primary.Metadata.TryGetValue("spCTag", out var storedTag)
                                    && storedTag == file.CTag;
                    if (isCurrent)
                    {
                        Interlocked.Increment(ref unchanged);
                        return;
                    }

                    var content = await sharePoint.DownloadAsync(file.ItemId, token);
                    await knowledgeSource.UpdateAsync(primary.FileId, file.RelativePath, content, metadata, token);
                    Interlocked.Increment(ref updated);
                    log.LogInformation("Replaced {File} ({FileId}).", file.RelativePath, primary.FileId);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    // One bad file shouldn't stop the run. It will be retried on the next tick
                    // because its stored cTag still won't match.
                    Interlocked.Increment(ref failed);
                    log.LogError(ex, "Failed to sync {File}.", file.RelativePath);
                }
            });

        sw.Stop();
        return new SyncResult(uploaded, updated, unchanged, deleted, skipped, failed, sw.Elapsed);
    }

    private static bool IsManagedByUs(KnowledgeSourceFile f) =>
        f.Metadata.TryGetValue("source", out var s) && s == SourceTag;
}
