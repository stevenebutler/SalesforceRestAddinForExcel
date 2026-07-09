using System;
using System.IO;
using System.Text.Json;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Persists connector behavior options. Missing file, invalid JSON, or invalid per-field values
/// fall back to <see cref="ConnectorOptions.Default"/> for each boolean (legacy RegDB absent-key behavior).
/// </summary>
public sealed class JsonConnectorOptionsStore : IConnectorOptionsStore
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
    };

    public JsonConnectorOptionsStore(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path is required.", nameof(filePath));
        }

        FilePath = filePath;
    }

    public string FilePath { get; }

    public ConnectorOptions Load()
    {
        if (!File.Exists(FilePath))
        {
            return ConnectorOptions.Default;
        }

        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(FilePath));
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                return ConnectorOptions.Default;
            }

            return new ConnectorOptions
            {
                UseReference = GetBoolOrDefault(root, "useReference"),
                NoWarning = GetBoolOrDefault(root, "noWarning"),
                ConfirmLargeQuery = GetBoolOrDefault(root, "confirmLargeQuery"),
                NoQueryLimit = GetBoolOrDefault(root, "noQueryLimit"),
                AutoAssignRule = GetBoolOrDefault(root, "autoAssignRule"),
                IncludeHiddenCells = ReadIncludeHiddenCells(root),
                CompositeBatchSize = GetIntOrDefault(root, "compositeBatchSize", ConnectorOptions.Default.CompositeBatchSize),
            };
        }
        catch (JsonException)
        {
            return ConnectorOptions.Default;
        }
    }

    public void Save(ConnectorOptions options)
    {
        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        var directory = Path.GetDirectoryName(FilePath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var payload = new
        {
            useReference = options.UseReference,
            noWarning = options.NoWarning,
            confirmLargeQuery = options.ConfirmLargeQuery,
            noQueryLimit = options.NoQueryLimit,
            autoAssignRule = options.AutoAssignRule,
            includeHiddenCells = options.IncludeHiddenCells,
            compositeBatchSize = options.CompositeBatchSize,
        };

        File.WriteAllText(FilePath, JsonSerializer.Serialize(payload, SerializerOptions));
    }

    /// <summary>
    /// Prefer <c>includeHiddenCells</c>; migrate legacy <c>skipHiddenCells</c> by inversion.
    /// Missing both → false (skip hidden by default).
    /// </summary>
    private static bool ReadIncludeHiddenCells(JsonElement root)
    {
        if (TryGetBool(root, "includeHiddenCells", out var include))
        {
            return include;
        }

        if (TryGetBool(root, "skipHiddenCells", out var skip))
        {
            return !skip;
        }

        return false;
    }

    private static bool GetBoolOrDefault(JsonElement root, string propertyName) =>
        TryGetBool(root, propertyName, out var value) && value;

    private static bool TryGetBool(JsonElement root, string propertyName, out bool value)
    {
        value = false;
        if (!root.TryGetProperty(propertyName, out var element))
        {
            return false;
        }

        switch (element.ValueKind)
        {
            case JsonValueKind.True:
                value = true;
                return true;
            case JsonValueKind.False:
                value = false;
                return true;
            default:
                return false;
        }
    }

    private static int GetIntOrDefault(JsonElement root, string propertyName, int defaultValue)
    {
        if (!root.TryGetProperty(propertyName, out var value))
        {
            return defaultValue;
        }

        return value.ValueKind switch
        {
            JsonValueKind.Number when value.TryGetInt32(out var n) && n > 0 => n,
            _ => defaultValue,
        };
    }
}
