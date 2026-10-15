using System.Net;
using Microsoft.Extensions.Logging;

namespace SharePointToFoundrySync;

internal static class HttpHelpers
{
    private const int MaxAttempts = 5;

    /// <summary>
    /// Sends a request, retrying throttling (429) and transient 5xx responses with exponential backoff.
    /// The factory is called once per attempt because an HttpRequestMessage can't be sent twice.
    /// 504 is deliberately NOT retried: from the knowledge source it means the file took more than
    /// 180 s to process, and sending it again won't change that.
    /// </summary>
    public static async Task<HttpResponseMessage> SendWithRetryAsync(
        this HttpClient http,
        Func<Task<HttpRequestMessage>> requestFactory,
        ILogger log,
        CancellationToken ct)
    {
        for (var attempt = 1; ; attempt++)
        {
            using var request = await requestFactory();
            HttpResponseMessage response;
            try
            {
                response = await http.SendAsync(request, HttpCompletionOption.ResponseContentRead, ct);
            }
            catch (HttpRequestException ex) when (attempt < MaxAttempts)
            {
                var wait = Backoff(attempt);
                log.LogWarning(ex, "{Method} {Url} failed (attempt {Attempt}); retrying in {Wait}.",
                    request.Method, Redact(request.RequestUri), attempt, wait);
                await Task.Delay(wait, ct);
                continue;
            }

            if (!IsTransient(response.StatusCode) || attempt >= MaxAttempts)
                return response;

            var delay = response.Headers.RetryAfter?.Delta
                        ?? (response.Headers.RetryAfter?.Date is { } date ? date - DateTimeOffset.UtcNow : (TimeSpan?)null)
                        ?? Backoff(attempt);
            if (delay < TimeSpan.Zero) delay = TimeSpan.FromSeconds(1);

            log.LogWarning("{Method} {Url} returned {Status} (attempt {Attempt}); retrying in {Wait}.",
                request.Method, Redact(request.RequestUri), (int)response.StatusCode, attempt, delay);
            response.Dispose();
            await Task.Delay(delay, ct);
        }
    }

    public static async Task EnsureSuccessAsync(this HttpResponseMessage response, string operation, CancellationToken ct)
    {
        if (response.IsSuccessStatusCode) return;

        var body = await response.Content.ReadAsStringAsync(ct);
        if (body.Length > 2000) body = body[..2000] + "…";
        throw new HttpRequestException(
            $"{operation} failed with {(int)response.StatusCode} {response.ReasonPhrase}: {body}",
            inner: null,
            statusCode: response.StatusCode);
    }

    private static bool IsTransient(HttpStatusCode status) =>
        status is HttpStatusCode.TooManyRequests
            or HttpStatusCode.InternalServerError
            or HttpStatusCode.BadGateway
            or HttpStatusCode.ServiceUnavailable;

    private static TimeSpan Backoff(int attempt) =>
        TimeSpan.FromSeconds(Math.Pow(2, attempt)) + TimeSpan.FromMilliseconds(Random.Shared.Next(0, 1000));

    // Never log query strings (pre-authenticated download URLs carry tokens in them).
    private static string Redact(Uri? uri) => uri is null ? "" : uri.GetLeftPart(UriPartial.Path);
}
