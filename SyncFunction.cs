using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace SharePointToFoundrySync;

public sealed class SyncFunction(SyncService sync, ILogger<SyncFunction> log)
{
    // Every 15 minutes, on the quarter hour (NCRONTAB: sec min hour day month day-of-week).
    // Timer triggers take a singleton lock, so a run that overruns won't overlap the next one.
    [Function("SharePointToKnowledgeSourceSync")]
    public async Task Run([TimerTrigger("0 */15 * * * *")] TimerInfo timer, CancellationToken ct)
    {
        if (timer.IsPastDue)
            log.LogWarning("Timer is past due; running now.");

        var r = await sync.RunAsync(ct);

        log.LogInformation(
            "Sync finished in {Elapsed:mm\\:ss}: {Uploaded} uploaded, {Updated} replaced, {Unchanged} unchanged, {Deleted} deleted, {Skipped} skipped, {Failed} failed.",
            r.Elapsed, r.Uploaded, r.Updated, r.Unchanged, r.Deleted, r.Skipped, r.Failed);

        // Surface partial failures as a failed invocation so they show up in monitoring and alerts.
        if (r.Failed > 0)
            throw new InvalidOperationException($"{r.Failed} file operation(s) failed; see the logged errors above. They will be retried on the next run.");
    }
}
