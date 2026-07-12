using System;
using System.Net;
using System.Net.Http;
using System.Text.Json;

namespace SalesforceRestAddin.Core.Rest;

/// <summary>
/// Classifies Salesforce REST and OAuth token-endpoint auth failures for session recovery.
/// </summary>
public static class SalesforceApiErrors
{
    public static bool IsSessionAuthFailure(HttpResponseMessage response)
    {
        if (response is null)
        {
            throw new ArgumentNullException(nameof(response));
        }

        if (response.StatusCode == HttpStatusCode.Unauthorized)
        {
            return true;
        }

        return TryReadBody(response, out var body) && ContainsSessionAuthErrorCode(body);
    }

    public static bool IsSessionAuthFailure(string? responseBody) =>
        responseBody is not null
        && !string.IsNullOrWhiteSpace(responseBody)
        && ContainsSessionAuthErrorCode(responseBody);

    public static bool IsInvalidGrant(HttpResponseMessage response, string? responseBody = null)
    {
        if (response is null)
        {
            throw new ArgumentNullException(nameof(response));
        }

        if (response.StatusCode is not (HttpStatusCode.BadRequest or HttpStatusCode.Unauthorized))
        {
            return false;
        }

        responseBody ??= TryReadBodySync(response);
        return responseBody is not null && ContainsInvalidGrant(responseBody);
    }

    public static bool IsInvalidGrant(string? messageOrBody) =>
        messageOrBody is not null
        && !string.IsNullOrWhiteSpace(messageOrBody)
        && ContainsInvalidGrant(messageOrBody);

    public static bool IsInvalidGrant(Exception exception)
    {
        for (var current = exception; current is not null; current = current.InnerException)
        {
            if (!string.IsNullOrWhiteSpace(current.Message) && ContainsInvalidGrant(current.Message))
            {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsSessionAuthErrorCode(string body)
    {
        if (body.Contains("INVALID_SESSION_ID", StringComparison.Ordinal)
            || body.Contains("INVALID_AUTH_HEADER", StringComparison.Ordinal)
            || body.Contains("Missing_OAuth_Token", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        try
        {
            using var document = JsonDocument.Parse(body);
            var root = document.RootElement;

            if (root.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in root.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.Object)
                    {
                        continue;
                    }

                    if (item.TryGetProperty("errorCode", out var errorCode)
                        && IsSessionAuthErrorCode(errorCode.GetString()))
                    {
                        return true;
                    }
                }
            }
            else if (root.ValueKind == JsonValueKind.Object
                     && root.TryGetProperty("errorCode", out var objectErrorCode)
                     && IsSessionAuthErrorCode(objectErrorCode.GetString()))
            {
                return true;
            }
        }
        catch (JsonException)
        {
            // Non-JSON bodies are matched by substring above.
        }

        return false;
    }

    private static bool IsSessionAuthErrorCode(string? errorCode) =>
        string.Equals(errorCode, "INVALID_SESSION_ID", StringComparison.OrdinalIgnoreCase)
        || string.Equals(errorCode, "INVALID_AUTH_HEADER", StringComparison.OrdinalIgnoreCase)
        || string.Equals(errorCode, "Missing_OAuth_Token", StringComparison.OrdinalIgnoreCase);

    private static bool ContainsInvalidGrant(string body) =>
        body.Contains("\"error\":\"invalid_grant\"", StringComparison.OrdinalIgnoreCase)
        || body.Contains("invalid_grant", StringComparison.OrdinalIgnoreCase);

    private static bool TryReadBody(HttpResponseMessage response, out string body)
    {
        body = TryReadBodySync(response);
        return !string.IsNullOrWhiteSpace(body);
    }

    private static string TryReadBodySync(HttpResponseMessage response)
    {
        try
        {
            return response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        }
        catch
        {
            return string.Empty;
        }
    }
}
