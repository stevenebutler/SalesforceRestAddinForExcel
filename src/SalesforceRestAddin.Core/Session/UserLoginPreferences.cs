using System;
using System.Collections.Generic;
using SalesforceRestAddin.Core.OAuth;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Login target fields restored between sessions.
/// </summary>
public sealed class UserLoginPreferences
{
    public string? Tenant { get; init; }

    public SalesforceEnvironment Environment { get; init; } = SalesforceEnvironment.Production;

    public string? SandboxId { get; init; }

    /// <summary>
    /// Preferred REST API version (e.g. <c>"66.0"</c>). <c>null</c> means latest supported by the org.
    /// </summary>
    public string? ApiVersion { get; init; }

    /// <summary>
    /// Supported versions last returned by <c>GET /services/data/</c> for the signed-in org.
    /// </summary>
    public IReadOnlyList<string> CachedSupportedApiVersions { get; init; } = Array.Empty<string>();

    /// <summary>
    /// When true, the next Salesforce API call shows login options (set by toolbar logout).
    /// </summary>
    public bool ShowLoginOptionsOnNextUse { get; init; }

    /// <summary>
    /// Debug-only: hand-edited JWT bearer from offline signing. When set, refresh-token login is skipped.
    /// Not shown in the Options UI; elided from JSON when unset.
    /// </summary>
    public string? JwtAccessToken { get; init; }

    public bool HasJwtAccessToken => !string.IsNullOrWhiteSpace(JwtAccessToken);

    public SalesforceLoginOptions ToLoginOptions() => new()
    {
        Tenant = Tenant,
        Environment = Environment,
        SandboxId = SandboxId,
    };
}
