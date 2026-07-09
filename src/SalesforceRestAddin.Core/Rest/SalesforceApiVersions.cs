using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;

namespace SalesforceRestAddin.Core.Rest;

/// <summary>
/// Salesforce REST API version strings as returned by <c>GET /services/data/</c>.
/// </summary>
public static class SalesforceApiVersions
{
    public static IReadOnlyList<string> ParseSupportedVersionsFromJson(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("Salesforce API versions response must be a JSON array.");
        }

        var versions = new List<string>();
        foreach (var entry in root.EnumerateArray())
        {
            if (!entry.TryGetProperty("version", out var versionElement))
            {
                continue;
            }

            var versionText = versionElement.GetString();
            if (versionText is null || string.IsNullOrWhiteSpace(versionText))
            {
                continue;
            }

            versions.Add(Normalize(versionText));
        }

        if (versions.Count == 0)
        {
            throw new InvalidOperationException("Salesforce API versions response did not contain a version.");
        }

        versions.Sort(CompareDescending);
        return versions;
    }

    public static string ParseLatestFromVersionsJson(string json) =>
        ParseSupportedVersionsFromJson(json)[0];

    public static string Normalize(string version)
    {
        if (string.IsNullOrWhiteSpace(version))
        {
            throw new ArgumentException("API version is required.", nameof(version));
        }

        if (!double.TryParse(version.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed))
        {
            throw new ArgumentException($"API version '{version}' is not valid.", nameof(version));
        }

        return parsed.ToString("0.0", CultureInfo.InvariantCulture);
    }

    public static bool AreEqual(string left, string right) =>
        string.Equals(Normalize(left), Normalize(right), StringComparison.Ordinal);

    public static int CompareDescending(string left, string right) =>
        ParseSortKey(right).CompareTo(ParseSortKey(left));

    private static decimal ParseSortKey(string version) =>
        decimal.Parse(Normalize(version), CultureInfo.InvariantCulture);
}

/// <summary>
/// Chooses a REST version from org-supported versions and an optional user preference.
/// </summary>
public static class SalesforceApiVersionSelector
{
    public static ApiVersionResolution Resolve(string? preferredVersion, IReadOnlyList<string> supportedVersions)
    {
        if (supportedVersions is null)
        {
            throw new ArgumentNullException(nameof(supportedVersions));
        }

        if (supportedVersions.Count == 0)
        {
            if (preferredVersion is null || string.IsNullOrWhiteSpace(preferredVersion))
            {
                throw new InvalidOperationException(
                    "Cannot resolve API version without org version list or a preferred version.");
            }

            var offline = SalesforceApiVersions.Normalize(preferredVersion);
            return new ApiVersionResolution(
                SelectedVersion: offline,
                SupportedVersions: supportedVersions,
                PreferredVersionWasInvalid: false,
                UsedLatestFromOrg: false);
        }

        var latest = supportedVersions[0];
        if (preferredVersion is null || string.IsNullOrWhiteSpace(preferredVersion))
        {
            return new ApiVersionResolution(
                SelectedVersion: latest,
                SupportedVersions: supportedVersions,
                PreferredVersionWasInvalid: false,
                UsedLatestFromOrg: true);
        }

        var normalizedPreferred = SalesforceApiVersions.Normalize(preferredVersion);
        if (supportedVersions.Any(version => SalesforceApiVersions.AreEqual(version, normalizedPreferred)))
        {
            return new ApiVersionResolution(
                SelectedVersion: normalizedPreferred,
                SupportedVersions: supportedVersions,
                PreferredVersionWasInvalid: false,
                UsedLatestFromOrg: false);
        }

        return new ApiVersionResolution(
            SelectedVersion: latest,
            SupportedVersions: supportedVersions,
            PreferredVersionWasInvalid: true,
            UsedLatestFromOrg: true);
    }
}

public sealed record ApiVersionResolution(
    string SelectedVersion,
    IReadOnlyList<string> SupportedVersions,
    bool PreferredVersionWasInvalid,
    bool UsedLatestFromOrg);
