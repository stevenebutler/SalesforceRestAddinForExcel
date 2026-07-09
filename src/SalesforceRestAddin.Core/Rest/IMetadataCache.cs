namespace SalesforceRestAddin.Core.Rest;

/// <summary>
/// Persistent cache for Salesforce metadata (global object list and per-object describe).
/// Entries are partitioned by Salesforce <paramref name="instanceUrl"/> so sandboxes and
/// production never share metadata. See NFR-META-1 in docs/ribbon/common-performance-requirements.md.
/// </summary>
public interface IMetadataCache
{
    /// <summary>Returns false when <paramref name="instanceUrl"/> is missing/invalid or the entry is absent.</summary>
    bool TryGetObjectList(string instanceUrl, string apiVersion, out string? json);

    /// <summary>No-op when <paramref name="instanceUrl"/> is missing/invalid.</summary>
    void SetObjectList(string instanceUrl, string apiVersion, string json);

    /// <summary>Returns false when <paramref name="instanceUrl"/> is missing/invalid or the entry is absent.</summary>
    bool TryGetDescribe(string instanceUrl, string apiVersion, string objectApiName, out string? json);

    /// <summary>No-op when <paramref name="instanceUrl"/> is missing/invalid.</summary>
    void SetDescribe(string instanceUrl, string apiVersion, string objectApiName, string json);

    /// <summary>
    /// Removes durable cache entries for a single Salesforce instance (all API versions / objects).
    /// Does not affect other instances or sandboxes.
    /// </summary>
    void ClearInstance(string instanceUrl);
}
