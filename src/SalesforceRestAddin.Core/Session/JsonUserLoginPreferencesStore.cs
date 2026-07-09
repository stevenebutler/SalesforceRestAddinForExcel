using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Persists non-secret login dialog defaults to a JSON file.
/// </summary>
public sealed class JsonUserLoginPreferencesStore : IUserLoginPreferencesStore
{
    private static readonly System.Text.Json.JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase,
        WriteIndented = true,
    };

    public JsonUserLoginPreferencesStore(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path is required.", nameof(filePath));
        }

        FilePath = filePath;
    }

    public string FilePath { get; }

    public UserLoginPreferences? TryLoadValid()
    {
        if (!File.Exists(FilePath))
        {
            return null;
        }

        try
        {
            var json = File.ReadAllText(FilePath);
            var document = System.Text.Json.JsonDocument.Parse(json);
            var root = document.RootElement;

            if (!root.TryGetProperty("environment", out var environmentElement))
            {
                return null;
            }

            var environmentText = environmentElement.GetString();
            if (string.IsNullOrWhiteSpace(environmentText))
            {
                return null;
            }

            SalesforceEnvironment environment;
            if (string.Equals(environmentText, "Production", StringComparison.OrdinalIgnoreCase))
            {
                environment = SalesforceEnvironment.Production;
            }
            else if (string.Equals(environmentText, "Sandbox", StringComparison.OrdinalIgnoreCase))
            {
                environment = SalesforceEnvironment.Sandbox;
            }
            else
            {
                return null;
            }

            return new UserLoginPreferences
            {
                Tenant = GetOptionalString(root, "tenant"),
                Environment = environment,
                SandboxId = GetOptionalString(root, "sandboxId"),
                ApiVersion = GetOptionalApiVersion(root, "apiVersion"),
                CachedSupportedApiVersions = GetOptionalStringArray(root, "cachedSupportedApiVersions"),
                ShowLoginOptionsOnNextUse = GetOptionalBool(root, "showLoginOptionsOnNextUse"),
                JwtAccessToken = NormalizeOptionalToken(GetOptionalString(root, "jwtAccessToken")),
            };
        }
        catch (System.Text.Json.JsonException)
        {
            return null;
        }
    }

    public UserLoginPreferences Load() => TryLoadValid() ?? new UserLoginPreferences();

    public void Save(UserLoginPreferences preferences)
    {
        if (preferences is null)
        {
            throw new ArgumentNullException(nameof(preferences));
        }

        var directory = Path.GetDirectoryName(FilePath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // Omit jwtAccessToken when unset so normal saves never introduce the debug key.
        object payload = string.IsNullOrWhiteSpace(preferences.JwtAccessToken)
            ? new
            {
                tenant = preferences.Tenant,
                environment = preferences.Environment.ToString(),
                sandboxId = preferences.SandboxId,
                apiVersion = preferences.ApiVersion,
                cachedSupportedApiVersions = preferences.CachedSupportedApiVersions.Count == 0
                    ? null
                    : preferences.CachedSupportedApiVersions,
                showLoginOptionsOnNextUse = preferences.ShowLoginOptionsOnNextUse,
            }
            : new
            {
                tenant = preferences.Tenant,
                environment = preferences.Environment.ToString(),
                sandboxId = preferences.SandboxId,
                apiVersion = preferences.ApiVersion,
                cachedSupportedApiVersions = preferences.CachedSupportedApiVersions.Count == 0
                    ? null
                    : preferences.CachedSupportedApiVersions,
                showLoginOptionsOnNextUse = preferences.ShowLoginOptionsOnNextUse,
                jwtAccessToken = preferences.JwtAccessToken,
            };

        File.WriteAllText(FilePath, System.Text.Json.JsonSerializer.Serialize(payload, SerializerOptions));
    }

    private static string? NormalizeOptionalToken(string? value) =>
        value is null || string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static string? GetOptionalString(System.Text.Json.JsonElement root, string propertyName) =>
        root.TryGetProperty(propertyName, out var value) ? value.GetString() : null;

    private static bool GetOptionalBool(System.Text.Json.JsonElement root, string propertyName) =>
        root.TryGetProperty(propertyName, out var value)
        && value.ValueKind == System.Text.Json.JsonValueKind.True;

    private static string? GetOptionalApiVersion(System.Text.Json.JsonElement root, string propertyName)
    {
        var text = GetOptionalString(root, propertyName);
        return text is null || string.IsNullOrWhiteSpace(text)
            ? null
            : SalesforceApiVersions.Normalize(text);
    }

    private static IReadOnlyList<string> GetOptionalStringArray(
        System.Text.Json.JsonElement root,
        string propertyName)
    {
        if (!root.TryGetProperty(propertyName, out var value) || value.ValueKind != System.Text.Json.JsonValueKind.Array)
        {
            return Array.Empty<string>();
        }

        return value.EnumerateArray()
            .Select(element => element.ValueKind == System.Text.Json.JsonValueKind.Number
                ? element.GetDouble().ToString("0.0", System.Globalization.CultureInfo.InvariantCulture)
                : element.GetString())
            .Where(version => !string.IsNullOrWhiteSpace(version))
            .Select(version => SalesforceApiVersions.Normalize(version!))
            .ToArray();
    }
}
