namespace SalesforceRestAddin.Core.OAuth;

/// <summary>
/// Normalizes login options from plain tenant names or pasted Salesforce URLs.
/// </summary>
public static class SalesforceLoginOptionsNormalizer
{
    public static SalesforceLoginOptions Normalize(SalesforceLoginOptions options)
    {
        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        var tenantInput = options.Tenant?.Trim();
        if (tenantInput is null || string.IsNullOrWhiteSpace(tenantInput))
        {
            return new SalesforceLoginOptions
            {
                Tenant = null,
                Environment = options.Environment,
                SandboxId = NormalizeSandboxId(options.SandboxId),
            };
        }

        var parsed = SalesforceInstanceHost.TryParse(tenantInput);
        if (parsed is not null)
        {
            if (parsed.IsGenericLoginHost)
            {
                return new SalesforceLoginOptions
                {
                    Tenant = null,
                    SandboxId = null,
                    Environment = parsed.Environment,
                };
            }

            if (parsed.HasResolvableTenant)
            {
                return new SalesforceLoginOptions
                {
                    Tenant = parsed.Tenant,
                    SandboxId = parsed.SandboxId ?? NormalizeSandboxId(options.SandboxId),
                    Environment = parsed.Environment,
                };
            }
        }

        return new SalesforceLoginOptions
        {
            Tenant = tenantInput.ToLowerInvariant(),
            Environment = options.Environment,
            SandboxId = NormalizeSandboxId(options.SandboxId),
        };
    }

    private static string? NormalizeSandboxId(string? sandboxId) =>
        sandboxId is null || string.IsNullOrWhiteSpace(sandboxId)
            ? null
            : sandboxId.Trim().ToLowerInvariant();
}
