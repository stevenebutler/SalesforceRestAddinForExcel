using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json;

namespace SalesforceRestAddin.Core.OAuth;

public sealed class OAuthTokenResponse
{
    private static readonly HashSet<string> SecretPropertyNames = new(StringComparer.Ordinal)
    {
        "access_token",
        "refresh_token",
    };

    private static readonly HashSet<string> KnownPropertyNames = new(StringComparer.Ordinal)
    {
        "access_token",
        "refresh_token",
        "instance_url",
        "id",
        "token_type",
        "issued_at",
        "scope",
        "signature",
        "sfdc_community_url",
        "sfdc_community_id",
    };

    public required string AccessToken { get; init; }

    public string? RefreshToken { get; init; }

    public required string InstanceUrl { get; init; }

    public required string Id { get; init; }

    public string? TokenType { get; init; }

    public long? IssuedAtUnixMs { get; init; }

    public string? Scope { get; init; }

    public string? Signature { get; init; }

    public string? SfdcCommunityUrl { get; init; }

    public string? SfdcCommunityId { get; init; }

    /// <summary>
    /// Any other token-response properties except secret token values.
    /// </summary>
    public IReadOnlyDictionary<string, string> AdditionalProperties { get; init; } =
        new Dictionary<string, string>();

    public bool HasAccessToken => !string.IsNullOrWhiteSpace(AccessToken);

    public bool HasRefreshToken => !string.IsNullOrWhiteSpace(RefreshToken);

    public void ApplyTo(SessionContext session)
    {
        session.AccessToken = AccessToken;
        if (!string.IsNullOrWhiteSpace(RefreshToken))
        {
            session.RefreshToken = RefreshToken;
        }

        session.InstanceUrl = InstanceUrl;
        session.Id = Id;
        session.IssuedAtUnixMs = IssuedAtUnixMs;
    }

    /// <summary>
    /// Human-readable OAuth token metadata for diagnostics (never includes token secrets).
    /// </summary>
    public string FormatMetadataSummary(DateTimeOffset? utcNow = null)
    {
        utcNow ??= DateTimeOffset.UtcNow;
        var lines = new List<string>
        {
            "OAuth token response (secrets omitted):",
            $"access_token: {(HasAccessToken ? "present" : "absent")}",
            $"refresh_token: {(HasRefreshToken ? "present" : "absent")}",
            $"instance_url: {InstanceUrl}",
            $"id: {Id}",
            $"token_type: {TokenType ?? "(not returned)"}",
            $"issued_at: {FormatIssuedAt()}",
            $"scope: {Scope ?? "(not returned)"}",
            $"signature: {Signature ?? "(not returned)"}",
            $"sfdc_community_url: {SfdcCommunityUrl ?? "(not returned)"}",
            $"sfdc_community_id: {SfdcCommunityId ?? "(not returned)"}",
        };

        foreach (var pair in AdditionalProperties.OrderBy(static entry => entry.Key, StringComparer.Ordinal))
        {
            lines.Add($"{pair.Key}: {pair.Value}");
        }

        lines.Add(string.Empty);
        lines.Add("Lifetime notes:");
        lines.Add("- Salesforce token responses do not include expires_in for the access token.");
        lines.Add("- issued_at is a Unix timestamp in milliseconds when Salesforce issued the access token.");
        if (IssuedAtUnixMs is long issuedAtMs)
        {
            var age = utcNow.Value - DateTimeOffset.FromUnixTimeMilliseconds(issuedAtMs);
            lines.Add($"- Access token age (now - issued_at): {FormatDuration(age)}.");
            lines.Add("- The add-in currently treats access tokens older than 2 hours as expired (heuristic).");
        }
        else
        {
            lines.Add("- issued_at was not returned; connector expiry heuristic is not applied.");
        }

        if (HasRefreshToken)
        {
            lines.Add("- Refresh token lifetime is controlled by Salesforce org/session policy, not returned here.");
        }

        return string.Join(Environment.NewLine, lines);
    }

    public static OAuthTokenResponse ParseJson(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;

        var accessToken = GetRequiredString(root, "access_token");
        var instanceUrl = GetRequiredString(root, "instance_url");
        var id = GetRequiredString(root, "id");

        return new OAuthTokenResponse
        {
            AccessToken = accessToken,
            RefreshToken = GetOptionalString(root, "refresh_token"),
            InstanceUrl = instanceUrl,
            Id = id,
            TokenType = GetOptionalString(root, "token_type"),
            IssuedAtUnixMs = ParseIssuedAtUnixMs(root),
            Scope = GetOptionalString(root, "scope"),
            Signature = GetOptionalString(root, "signature"),
            SfdcCommunityUrl = GetOptionalString(root, "sfdc_community_url"),
            SfdcCommunityId = GetOptionalString(root, "sfdc_community_id"),
            AdditionalProperties = ParseAdditionalProperties(root),
        };
    }

    private string FormatIssuedAt()
    {
        if (IssuedAtUnixMs is not long issuedAtMs)
        {
            return "(not returned)";
        }

        var utc = DateTimeOffset.FromUnixTimeMilliseconds(issuedAtMs).ToUniversalTime();
        return $"{issuedAtMs} ({utc:yyyy-MM-dd HH:mm:ss} UTC)";
    }

    private static long? ParseIssuedAtUnixMs(JsonElement root)
    {
        if (!root.TryGetProperty("issued_at", out var issuedAtElement))
        {
            return null;
        }

        if (issuedAtElement.ValueKind == JsonValueKind.Number && issuedAtElement.TryGetInt64(out var numeric))
        {
            return numeric;
        }

        var issuedAtText = issuedAtElement.GetString();
        return long.TryParse(issuedAtText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static IReadOnlyDictionary<string, string> ParseAdditionalProperties(JsonElement root)
    {
        var additional = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var property in root.EnumerateObject())
        {
            if (SecretPropertyNames.Contains(property.Name) || KnownPropertyNames.Contains(property.Name))
            {
                continue;
            }

            additional[property.Name] = property.Value.ValueKind switch
            {
                JsonValueKind.String => property.Value.GetString() ?? string.Empty,
                JsonValueKind.Number => property.Value.GetRawText(),
                JsonValueKind.True => "true",
                JsonValueKind.False => "false",
                JsonValueKind.Null => "(null)",
                _ => property.Value.GetRawText(),
            };
        }

        return additional;
    }

    private static string FormatDuration(TimeSpan duration)
    {
        if (duration < TimeSpan.Zero)
        {
            duration = -duration;
        }

        if (duration.TotalHours >= 1)
        {
            return $"{duration.TotalHours:0.#} hours";
        }

        if (duration.TotalMinutes >= 1)
        {
            return $"{duration.TotalMinutes:0.#} minutes";
        }

        return $"{duration.TotalSeconds:0.#} seconds";
    }

    private static string GetRequiredString(JsonElement root, string propertyName)
    {
        if (!root.TryGetProperty(propertyName, out var value))
        {
            throw new InvalidOperationException($"OAuth token response is missing '{propertyName}'.");
        }

        var text = value.GetString();
        if (text is null || string.IsNullOrWhiteSpace(text))
        {
            throw new InvalidOperationException($"OAuth token response property '{propertyName}' is empty.");
        }

        return text;
    }

    private static string? GetOptionalString(JsonElement root, string propertyName) =>
        root.TryGetProperty(propertyName, out var value) ? value.GetString() : null;
}
