using System;
using System.Text.RegularExpressions;

namespace SalesforceRestAddin.Core.OAuth;

/// <summary>
/// Parsed Salesforce login or instance hostname (My Domain, generic login, or named sandbox).
/// </summary>
public sealed class SalesforceInstanceHost
{
    private static readonly Regex NamedSandboxHost = new(
        @"^(?<tenant>[a-z0-9-]+)--(?<sandbox>[a-z0-9-]+)\.sandbox\.my\.salesforce\.com$",
        RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture);

    private static readonly Regex SandboxMyDomainHost = new(
        @"^(?<tenant>[a-z0-9-]+)\.sandbox\.my\.salesforce\.com$",
        RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture);

    private static readonly Regex ProductionMyDomainHost = new(
        @"^(?<tenant>[a-z0-9-]+)\.my\.salesforce\.com$",
        RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture);

    private SalesforceInstanceHost(
        string hostname,
        string? tenant,
        string? sandboxId,
        SalesforceEnvironment environment,
        bool isGenericLoginHost)
    {
        Hostname = hostname;
        Tenant = tenant;
        SandboxId = sandboxId;
        Environment = environment;
        IsGenericLoginHost = isGenericLoginHost;
    }

    public string Hostname { get; }

    public string? Tenant { get; }

    public string? SandboxId { get; }

    public SalesforceEnvironment Environment { get; }

    public bool IsGenericLoginHost { get; }

    public bool HasResolvableTenant => Tenant is not null;

    public string HostKey => Hostname;

    public static SalesforceInstanceHost? TryParse(string? urlOrHost)
    {
        var hostname = ExtractHostname(urlOrHost);
        if (hostname is null)
        {
            return null;
        }

        if (string.Equals(hostname, "login.salesforce.com", StringComparison.Ordinal))
        {
            return new SalesforceInstanceHost(hostname, null, null, SalesforceEnvironment.Production, true);
        }

        if (string.Equals(hostname, "test.salesforce.com", StringComparison.Ordinal))
        {
            return new SalesforceInstanceHost(hostname, null, null, SalesforceEnvironment.Sandbox, true);
        }

        var namedSandbox = NamedSandboxHost.Match(hostname);
        if (namedSandbox.Success)
        {
            return new SalesforceInstanceHost(
                hostname,
                namedSandbox.Groups["tenant"].Value,
                namedSandbox.Groups["sandbox"].Value,
                SalesforceEnvironment.Sandbox,
                false);
        }

        var sandboxMyDomain = SandboxMyDomainHost.Match(hostname);
        if (sandboxMyDomain.Success)
        {
            return new SalesforceInstanceHost(
                hostname,
                sandboxMyDomain.Groups["tenant"].Value,
                null,
                SalesforceEnvironment.Sandbox,
                false);
        }

        var productionMyDomain = ProductionMyDomainHost.Match(hostname);
        if (productionMyDomain.Success)
        {
            return new SalesforceInstanceHost(
                hostname,
                productionMyDomain.Groups["tenant"].Value,
                null,
                SalesforceEnvironment.Production,
                false);
        }

        return null;
    }

    public static SalesforceInstanceHost? TryParseInstanceUrl(string? instanceUrl) => TryParse(instanceUrl);

    public SalesforceLoginOptions ToLoginOptions() => new()
    {
        Tenant = Tenant,
        SandboxId = SandboxId,
        Environment = Environment,
    };

    private static string? ExtractHostname(string? input)
    {
        if (input is null || string.IsNullOrWhiteSpace(input))
        {
            return null;
        }

        var trimmed = input.Trim();
        if (trimmed.IndexOf("://", StringComparison.Ordinal) >= 0)
        {
            if (!Uri.TryCreate(trimmed, UriKind.Absolute, out var absoluteUri))
            {
                return null;
            }

            return absoluteUri.Host.ToLowerInvariant();
        }

        if (trimmed.IndexOf('/') >= 0)
        {
            if (Uri.TryCreate("https://" + trimmed, UriKind.Absolute, out var uriWithPath))
            {
                return uriWithPath.Host.ToLowerInvariant();
            }

            return null;
        }

        return trimmed.TrimEnd('/').ToLowerInvariant();
    }
}
