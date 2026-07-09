using System;
using System.Collections.Generic;
using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Core.Rest;

/// <summary>
/// In-process L1 cache of parsed Salesforce metadata (object lists, describes, field catalogs).
/// Durable JSON remains in <see cref="IMetadataCache"/> / <see cref="FileMetadataCache"/>.
/// </summary>
public sealed class ParsedMetadataCache
{
    private readonly object _gate = new();
    private readonly Dictionary<string, IReadOnlyList<SObjectSummary>> _objectLists =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DescribeCacheEntry> _describes =
        new(StringComparer.OrdinalIgnoreCase);

    public bool TryGetObjectList(
        string instanceUrl,
        string apiVersion,
        out IReadOnlyList<SObjectSummary>? objects)
    {
        if (!TryBuildObjectListKey(instanceUrl, apiVersion, out var key))
        {
            objects = null;
            return false;
        }

        lock (_gate)
        {
            if (_objectLists.TryGetValue(key, out var cached))
            {
                objects = cached;
                return true;
            }
        }

        objects = null;
        return false;
    }

    public void SetObjectList(string instanceUrl, string apiVersion, IReadOnlyList<SObjectSummary> objects)
    {
        if (objects is null
            || !TryBuildObjectListKey(instanceUrl, apiVersion, out var key))
        {
            return;
        }

        lock (_gate)
        {
            _objectLists[key] = objects;
        }
    }

    public bool TryGetDescribe(
        string instanceUrl,
        string apiVersion,
        string objectApiName,
        out SObjectDescribe? describe)
    {
        if (!TryGetEntry(instanceUrl, apiVersion, objectApiName, out var entry))
        {
            describe = null;
            return false;
        }

        describe = entry!.Describe;
        return true;
    }

    public bool TryGetFieldCatalog(
        string instanceUrl,
        string apiVersion,
        string objectApiName,
        out FieldCatalog? catalog)
    {
        if (!TryGetEntry(instanceUrl, apiVersion, objectApiName, out var entry))
        {
            catalog = null;
            return false;
        }

        catalog = entry!.Catalog;
        return true;
    }

    public void SetDescribe(string instanceUrl, string apiVersion, string objectApiName, SObjectDescribe describe)
    {
        if (describe is null
            || !TryBuildDescribeKey(instanceUrl, apiVersion, objectApiName, out var key))
        {
            return;
        }

        var entry = new DescribeCacheEntry(describe, new FieldCatalog(describe));
        lock (_gate)
        {
            _describes[key] = entry;
        }
    }

    /// <summary>Drops all L1 entries for the given Salesforce instance (all API versions / objects).</summary>
    public void ClearInstance(string instanceUrl)
    {
        if (!FileMetadataCache.TryNormalizeInstanceKey(instanceUrl, out var instanceKey))
        {
            return;
        }

        var prefix = instanceKey + "|";
        lock (_gate)
        {
            RemoveByPrefix(_objectLists, prefix);
            RemoveByPrefix(_describes, prefix);
        }
    }

    private bool TryGetEntry(
        string instanceUrl,
        string apiVersion,
        string objectApiName,
        out DescribeCacheEntry? entry)
    {
        if (!TryBuildDescribeKey(instanceUrl, apiVersion, objectApiName, out var key))
        {
            entry = null;
            return false;
        }

        lock (_gate)
        {
            if (_describes.TryGetValue(key, out var cached))
            {
                entry = cached;
                return true;
            }
        }

        entry = null;
        return false;
    }

    private static void RemoveByPrefix<T>(Dictionary<string, T> map, string prefix)
    {
        var toRemove = new List<string>();
        foreach (var key in map.Keys)
        {
            if (key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                toRemove.Add(key);
            }
        }

        foreach (var key in toRemove)
        {
            map.Remove(key);
        }
    }

    private static bool TryBuildObjectListKey(string instanceUrl, string apiVersion, out string key)
    {
        key = string.Empty;
        if (!FileMetadataCache.TryNormalizeInstanceKey(instanceUrl, out var instanceKey)
            || string.IsNullOrWhiteSpace(apiVersion))
        {
            return false;
        }

        key = instanceKey + "|" + FileMetadataCache.SanitizeSegment(apiVersion);
        return true;
    }

    private static bool TryBuildDescribeKey(
        string instanceUrl,
        string apiVersion,
        string objectApiName,
        out string key)
    {
        key = string.Empty;
        if (!FileMetadataCache.TryNormalizeInstanceKey(instanceUrl, out var instanceKey)
            || string.IsNullOrWhiteSpace(apiVersion)
            || string.IsNullOrWhiteSpace(objectApiName))
        {
            return false;
        }

        key = instanceKey
            + "|"
            + FileMetadataCache.SanitizeSegment(apiVersion)
            + "|"
            + FileMetadataCache.SanitizeSegment(objectApiName);
        return true;
    }

    private sealed class DescribeCacheEntry
    {
        public DescribeCacheEntry(SObjectDescribe describe, FieldCatalog catalog)
        {
            Describe = describe;
            Catalog = catalog;
        }

        public SObjectDescribe Describe { get; }

        public FieldCatalog Catalog { get; }
    }
}
