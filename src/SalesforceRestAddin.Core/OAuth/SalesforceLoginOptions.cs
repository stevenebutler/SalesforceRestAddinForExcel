namespace SalesforceRestAddin.Core.OAuth;

/// <summary>
/// User-selected login target before OAuth authorize.
/// </summary>
public sealed class SalesforceLoginOptions
{
    public string? Tenant { get; init; }

    public SalesforceEnvironment Environment { get; init; } = SalesforceEnvironment.Production;

    /// <summary>
    /// Sandbox name used with My Domain, e.g. "dev" in kjr--dev.sandbox.my.salesforce.com.
    /// </summary>
    public string? SandboxId { get; init; }
}
