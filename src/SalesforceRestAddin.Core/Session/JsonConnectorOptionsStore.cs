using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Persists connector behavior options. Malformed files return defaults; invalid recognized
/// values fall back to that setting's default and are recorded in the session trace.
/// </summary>
public sealed class JsonConnectorOptionsStore : IConnectorOptionsStore
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
    };

    private readonly Action<string> _warningLogger;

    public JsonConnectorOptionsStore(string filePath)
        : this(filePath, SessionFlowTrace.Log)
    {
    }

    internal JsonConnectorOptionsStore(string filePath, Action<string> warningLogger)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path is required.", nameof(filePath));
        }

        _warningLogger = warningLogger ?? throw new ArgumentNullException(nameof(warningLogger));

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
                LogWarning("Connector options: settings root is not an object; using defaults.");
                return ConnectorOptions.Default;
            }

            var invalidSettings = new List<string>();
            var options = new ConnectorOptions
            {
                UseReference = ReadBool(root, "useReference", ConnectorOptions.Default.UseReference, invalidSettings),
                NoWarning = ReadBool(root, "noWarning", ConnectorOptions.Default.NoWarning, invalidSettings),
                NoConfirmQueryDownload = ReadBool(root, "noConfirmQueryDownload", ConnectorOptions.Default.NoConfirmQueryDownload, invalidSettings),
                NoQueryLimit = ReadBool(root, "noQueryLimit", ConnectorOptions.Default.NoQueryLimit, invalidSettings),
                AutoAssignRule = ReadBool(root, "autoAssignRule", ConnectorOptions.Default.AutoAssignRule, invalidSettings),
                IncludeHiddenCells = ReadIncludeHiddenCells(root, invalidSettings),
                ColumnSizingMode = ReadColumnSizingMode(root, invalidSettings),
                RowSizingMode = ReadRowSizingMode(root, invalidSettings),
                CompositeBatchSize = ReadPositiveInt(root, "compositeBatchSize", ConnectorOptions.Default.CompositeBatchSize, invalidSettings),
            };
            LogInvalidSettings(invalidSettings);
            return options;
        }
        catch (JsonException)
        {
            LogWarning("Connector options: settings JSON is invalid; using defaults.");
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
            noConfirmQueryDownload = options.NoConfirmQueryDownload,
            noQueryLimit = options.NoQueryLimit,
            autoAssignRule = options.AutoAssignRule,
            includeHiddenCells = options.IncludeHiddenCells,
            columnSizingMode = ToJsonValue(options.ColumnSizingMode),
            rowSizingMode = ToJsonValue(options.RowSizingMode),
            compositeBatchSize = options.CompositeBatchSize,
        };

        File.WriteAllText(FilePath, JsonSerializer.Serialize(payload, SerializerOptions));
    }

    /// <summary>
    /// Prefer <c>includeHiddenCells</c>; migrate legacy <c>skipHiddenCells</c> by inversion.
    /// Missing both → false (skip hidden by default).
    /// </summary>
    private static bool ReadIncludeHiddenCells(JsonElement root, ICollection<string> invalidSettings)
    {
        if (root.TryGetProperty("includeHiddenCells", out _))
        {
            return ReadBool(root, "includeHiddenCells", ConnectorOptions.Default.IncludeHiddenCells, invalidSettings);
        }

        if (root.TryGetProperty("skipHiddenCells", out _))
        {
            return !ReadBool(root, "skipHiddenCells", defaultValue: true, invalidSettings);
        }

        return ConnectorOptions.Default.IncludeHiddenCells;
    }

    private static bool ReadBool(
        JsonElement root,
        string propertyName,
        bool defaultValue,
        ICollection<string> invalidSettings)
    {
        if (!root.TryGetProperty(propertyName, out var element))
        {
            return defaultValue;
        }

        switch (element.ValueKind)
        {
            case JsonValueKind.True:
                return true;
            case JsonValueKind.False:
                return false;
            default:
                invalidSettings.Add(propertyName);
                return defaultValue;
        }
    }

    private static int ReadPositiveInt(
        JsonElement root,
        string propertyName,
        int defaultValue,
        ICollection<string> invalidSettings)
    {
        if (!root.TryGetProperty(propertyName, out var value))
        {
            return defaultValue;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number) && number > 0)
        {
            return number;
        }

        invalidSettings.Add(propertyName);
        return defaultValue;
    }

    private static ColumnSizingMode ReadColumnSizingMode(JsonElement root, ICollection<string> invalidSettings) =>
        ReadEnum(
            root,
            "columnSizingMode",
            ConnectorOptions.Default.ColumnSizingMode,
            invalidSettings,
            new Dictionary<string, ColumnSizingMode>(StringComparer.Ordinal)
            {
                ["allDownloadedData"] = ColumnSizingMode.AllDownloadedData,
                ["firstDownloadedPage"] = ColumnSizingMode.FirstDownloadedPage,
                ["headersOnly"] = ColumnSizingMode.HeadersOnly,
            });

    private static RowSizingMode ReadRowSizingMode(JsonElement root, ICollection<string> invalidSettings) =>
        ReadEnum(
            root,
            "rowSizingMode",
            ConnectorOptions.Default.RowSizingMode,
            invalidSettings,
            new Dictionary<string, RowSizingMode>(StringComparer.Ordinal)
            {
                ["fitEachPage"] = RowSizingMode.FitEachPage,
                ["forceSingleLine"] = RowSizingMode.ForceSingleLine,
                ["none"] = RowSizingMode.None,
            });

    private static T ReadEnum<T>(
        JsonElement root,
        string propertyName,
        T defaultValue,
        ICollection<string> invalidSettings,
        IReadOnlyDictionary<string, T> values)
    {
        if (!root.TryGetProperty(propertyName, out var element))
        {
            return defaultValue;
        }

        if (element.ValueKind == JsonValueKind.String
            && element.GetString() is { } text
            && values.TryGetValue(text, out var value))
        {
            return value;
        }

        invalidSettings.Add(propertyName);
        return defaultValue;
    }

    private static string ToJsonValue(ColumnSizingMode mode) =>
        mode switch
        {
            ColumnSizingMode.AllDownloadedData => "allDownloadedData",
            ColumnSizingMode.FirstDownloadedPage => "firstDownloadedPage",
            ColumnSizingMode.HeadersOnly => "headersOnly",
            _ => throw new ArgumentOutOfRangeException(nameof(mode)),
        };

    private static string ToJsonValue(RowSizingMode mode) =>
        mode switch
        {
            RowSizingMode.FitEachPage => "fitEachPage",
            RowSizingMode.ForceSingleLine => "forceSingleLine",
            RowSizingMode.None => "none",
            _ => throw new ArgumentOutOfRangeException(nameof(mode)),
        };

    private void LogInvalidSettings(IReadOnlyCollection<string> invalidSettings)
    {
        if (invalidSettings.Count > 0)
        {
            LogWarning(
                $"Connector options: invalid settings ({string.Join(", ", invalidSettings)}); using defaults for those settings.");
        }
    }

    private void LogWarning(string message) => _warningLogger(message);
}
