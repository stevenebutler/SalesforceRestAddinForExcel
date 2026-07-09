using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SalesforceRestAddin.Core.Rest;

/// <summary>
/// Disk-backed metadata cache under <c>%LOCALAPPDATA%\SalesforceRestAddin\metadata-cache\</c>.
/// Keys: Salesforce instance URL (per org/sandbox) + API version (+ object API name for describe).
/// </summary>
/// <remarks>
/// <see cref="ClearInstance"/> only deletes under a root that this type has previously written
/// (sentinel marker present), and only the folder for that Salesforce instance.
/// Construction rejects well-known dangerous roots (profile, drive root, etc.).
/// </remarks>
public sealed class FileMetadataCache : IMetadataCache
{
    internal const string MarkerFileName = ".forceconnector-metadata-cache";

    private readonly string _rootDirectory;
    private readonly object _gate = new();

    public FileMetadataCache(string rootDirectory)
    {
        if (string.IsNullOrWhiteSpace(rootDirectory))
        {
            throw new ArgumentException("Root directory is required.", nameof(rootDirectory));
        }

        var fullPath = Path.GetFullPath(rootDirectory.Trim());
        if (!IsSafeCacheRoot(fullPath))
        {
            throw new ArgumentException(
                "Metadata cache root must not be a user profile, app-data root, or filesystem drive root.",
                nameof(rootDirectory));
        }

        _rootDirectory = fullPath;
    }

    public bool TryGetObjectList(string instanceUrl, string apiVersion, out string? json)
    {
        if (!TryNormalizeInstanceKey(instanceUrl, out var instanceKey)
            || string.IsNullOrWhiteSpace(apiVersion))
        {
            json = null;
            return false;
        }

        return TryRead(ObjectListPath(instanceKey, apiVersion), out json);
    }

    public void SetObjectList(string instanceUrl, string apiVersion, string json)
    {
        if (!TryNormalizeInstanceKey(instanceUrl, out var instanceKey)
            || string.IsNullOrWhiteSpace(apiVersion))
        {
            return;
        }

        Write(ObjectListPath(instanceKey, apiVersion), json);
    }

    public bool TryGetDescribe(string instanceUrl, string apiVersion, string objectApiName, out string? json)
    {
        if (!TryNormalizeInstanceKey(instanceUrl, out var instanceKey)
            || string.IsNullOrWhiteSpace(apiVersion)
            || string.IsNullOrWhiteSpace(objectApiName))
        {
            json = null;
            return false;
        }

        return TryRead(DescribePath(instanceKey, apiVersion, objectApiName), out json);
    }

    public void SetDescribe(string instanceUrl, string apiVersion, string objectApiName, string json)
    {
        if (!TryNormalizeInstanceKey(instanceUrl, out var instanceKey)
            || string.IsNullOrWhiteSpace(apiVersion)
            || string.IsNullOrWhiteSpace(objectApiName))
        {
            return;
        }

        Write(DescribePath(instanceKey, apiVersion, objectApiName), json);
    }

    public void ClearInstance(string instanceUrl)
    {
        if (!TryNormalizeInstanceKey(instanceUrl, out var instanceKey))
        {
            return;
        }

        lock (_gate)
        {
            if (!Directory.Exists(_rootDirectory) || !File.Exists(MarkerPath))
            {
                // No sentinel → refuse to delete (protects against a mis-pointed root).
                return;
            }

            var instanceDirectory = Path.Combine(_rootDirectory, instanceKey);
            if (!Directory.Exists(instanceDirectory))
            {
                return;
            }

            // Ensure we only delete a direct child of the cache root.
            var fullInstance = Path.GetFullPath(instanceDirectory);
            var fullRoot = Path.GetFullPath(_rootDirectory);
            if (!fullInstance.StartsWith(
                    fullRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    + Path.DirectorySeparatorChar,
                    StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            try
            {
                Directory.Delete(fullInstance, recursive: true);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }
    }

    private string MarkerPath => Path.Combine(_rootDirectory, MarkerFileName);

    private string ObjectListPath(string instanceKey, string apiVersion) =>
        Path.Combine(VersionDirectory(instanceKey, apiVersion), "sobjects.json");

    private string DescribePath(string instanceKey, string apiVersion, string objectApiName) =>
        Path.Combine(VersionDirectory(instanceKey, apiVersion), "describe", SanitizeSegment(objectApiName) + ".json");

    private string VersionDirectory(string instanceKey, string apiVersion) =>
        Path.Combine(_rootDirectory, instanceKey, SanitizeSegment(apiVersion));

    /// <summary>
    /// Builds a stable filesystem key from a Salesforce instance URL.
    /// Different hosts (production vs sandbox) must never share a key.
    /// Returns false for blank or non-absolute URLs so callers skip the cache entirely.
    /// </summary>
    internal static bool TryNormalizeInstanceKey(string? instanceUrl, out string instanceKey)
    {
        instanceKey = string.Empty;
        if (instanceUrl is null || string.IsNullOrWhiteSpace(instanceUrl))
        {
            return false;
        }

        if (!Uri.TryCreate(instanceUrl.Trim(), UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeHttp)
            || string.IsNullOrWhiteSpace(uri.Host))
        {
            return false;
        }

        // Host is the org/sandbox discriminator; ignore path/query noise on instance_url.
        var normalized = uri.GetLeftPart(UriPartial.Authority).TrimEnd('/');
        instanceKey = SanitizeSegment(normalized);
        return !string.IsNullOrEmpty(instanceKey) && instanceKey != "_";
    }

    private bool TryRead(string path, out string? json)
    {
        lock (_gate)
        {
            if (!File.Exists(path))
            {
                json = null;
                return false;
            }

            try
            {
                json = File.ReadAllText(path, Encoding.UTF8);
                return !string.IsNullOrWhiteSpace(json);
            }
            catch (IOException)
            {
                json = null;
                return false;
            }
        }
    }

    private void Write(string path, string json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return;
        }

        lock (_gate)
        {
            EnsureMarkerUnlocked();

            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(path, json, Encoding.UTF8);
        }
    }

    private void EnsureMarkerUnlocked()
    {
        Directory.CreateDirectory(_rootDirectory);
        if (!File.Exists(MarkerPath))
        {
            File.WriteAllText(
                MarkerPath,
                "SalesforceRestAddin metadata cache root. Cleared only when this marker is present.",
                Encoding.UTF8);
        }
    }

    /// <summary>
    /// Returns false for empty paths, filesystem/drive roots, and well-known user/app-data roots.
    /// Pure path policy — no I/O. Safe to call with real profile paths in tests.
    /// </summary>
    internal static bool IsSafeCacheRoot(string rootDirectory) =>
        IsSafeCacheRoot(rootDirectory, GetDefaultForbiddenRoots());

    /// <summary>
    /// Path-policy check with an injectable forbidden-root list (for unit tests).
    /// Never deletes; compares normalized full paths only.
    /// </summary>
    internal static bool IsSafeCacheRoot(string rootDirectory, IEnumerable<string> forbiddenRoots)
    {
        if (string.IsNullOrWhiteSpace(rootDirectory))
        {
            return false;
        }

        string fullPath;
        try
        {
            fullPath = NormalizePath(Path.GetFullPath(rootDirectory.Trim()));
        }
        catch (Exception)
        {
            return false;
        }

        if (IsFilesystemOrDriveRoot(fullPath))
        {
            return false;
        }

        foreach (var forbidden in forbiddenRoots)
        {
            if (string.IsNullOrWhiteSpace(forbidden))
            {
                continue;
            }

            string forbiddenFull;
            try
            {
                forbiddenFull = NormalizePath(Path.GetFullPath(forbidden.Trim()));
            }
            catch (Exception)
            {
                continue;
            }

            if (string.Equals(fullPath, forbiddenFull, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        return true;
    }

    internal static IReadOnlyList<string> GetDefaultForbiddenRoots()
    {
        var folders = new[]
        {
            Environment.SpecialFolder.UserProfile,
            Environment.SpecialFolder.LocalApplicationData,
            Environment.SpecialFolder.ApplicationData,
            Environment.SpecialFolder.CommonApplicationData,
            Environment.SpecialFolder.DesktopDirectory,
            Environment.SpecialFolder.MyDocuments,
            Environment.SpecialFolder.MyMusic,
            Environment.SpecialFolder.MyPictures,
            Environment.SpecialFolder.MyVideos,
        };

        var list = new List<string>();
        foreach (var folder in folders)
        {
            var path = Environment.GetFolderPath(folder);
            if (!string.IsNullOrWhiteSpace(path))
            {
                list.Add(path);
            }
        }

        var temp = Path.GetTempPath();
        if (!string.IsNullOrWhiteSpace(temp))
        {
            // Temp root itself is too broad; subdirs under temp are allowed.
            list.Add(temp);
        }

        return list;
    }

    private static bool IsFilesystemOrDriveRoot(string fullPath)
    {
        var pathRoot = Path.GetPathRoot(fullPath);
        if (string.IsNullOrEmpty(pathRoot))
        {
            return true;
        }

        return string.Equals(
            NormalizePath(fullPath),
            NormalizePath(pathRoot),
            StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizePath(string path) =>
        path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

    /// <summary>Filesystem-safe key segment from an instance URL or API name.</summary>
    internal static string SanitizeSegment(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "_";
        }

        var trimmed = value.Trim().TrimEnd('/');
        var builder = new StringBuilder(trimmed.Length);
        foreach (var ch in trimmed)
        {
            if (char.IsLetterOrDigit(ch) || ch is '-' or '_' or '.')
            {
                builder.Append(ch);
            }
            else
            {
                builder.Append('_');
            }
        }

        var result = builder.ToString();
        return string.IsNullOrEmpty(result) ? "_" : result;
    }
}
