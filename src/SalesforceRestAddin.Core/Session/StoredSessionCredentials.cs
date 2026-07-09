namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Refresh token and metadata persisted per Salesforce login host.
/// </summary>
public sealed class StoredSessionCredentials
{
    public required string HostKey { get; init; }

    public required string RefreshToken { get; init; }

    /// <summary>
    /// OAuth token endpoint host used when this refresh token was issued
    /// (e.g. login.salesforce.com after generic login, or the My Domain host).
    /// </summary>
    public string? OAuthTokenHost { get; init; }

    public string? InstanceUrl { get; init; }

    public string? Id { get; init; }
}
