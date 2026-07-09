namespace SalesforceRestAddin.Core.Rest;

/// <summary>Always misses — used by unit tests that drive HTTP via mock handlers.</summary>
public sealed class NoOpMetadataCache : IMetadataCache
{
    public static NoOpMetadataCache Instance { get; } = new();

    public bool TryGetObjectList(string instanceUrl, string apiVersion, out string? json)
    {
        json = null;
        return false;
    }

    public void SetObjectList(string instanceUrl, string apiVersion, string json)
    {
    }

    public bool TryGetDescribe(string instanceUrl, string apiVersion, string objectApiName, out string? json)
    {
        json = null;
        return false;
    }

    public void SetDescribe(string instanceUrl, string apiVersion, string objectApiName, string json)
    {
    }

    public void ClearInstance(string instanceUrl)
    {
    }
}
