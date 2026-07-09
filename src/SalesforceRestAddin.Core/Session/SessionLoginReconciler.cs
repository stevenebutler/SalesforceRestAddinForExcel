using SalesforceRestAddin.Core.OAuth;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// After OAuth, maps the org instance URL to canonical tenant/sandbox login options and credential keys.
/// </summary>
public static class SessionLoginReconciler
{
    public static SessionLoginReconciliation Reconcile(
        SalesforceLoginOptions loginOptionsUsed,
        OAuthTokenResponse token)
    {
        if (token is null)
        {
            throw new ArgumentNullException(nameof(token));
        }

        var normalizedUsed = SalesforceLoginOptionsNormalizer.Normalize(loginOptionsUsed);
        var oauthTokenHost = new Uri(SalesforceLoginUrlBuilder.BuildLoginBaseUrl(normalizedUsed)).Host;
        var fromInstance = SalesforceInstanceHost.TryParseInstanceUrl(token.InstanceUrl);

        if (fromInstance is not null && fromInstance.HasResolvableTenant)
        {
            return new SessionLoginReconciliation(
                ResolvedLoginOptions: fromInstance.ToLoginOptions(),
                CanonicalHostKey: fromInstance.HostKey,
                OAuthTokenHost: oauthTokenHost);
        }

        var fallbackHost = TryGetInstanceHostname(token.InstanceUrl) ?? oauthTokenHost;
        return new SessionLoginReconciliation(
            ResolvedLoginOptions: normalizedUsed,
            CanonicalHostKey: fallbackHost,
            OAuthTokenHost: oauthTokenHost);
    }

    private static string? TryGetInstanceHostname(string? instanceUrl)
    {
        if (string.IsNullOrWhiteSpace(instanceUrl))
        {
            return null;
        }

        return Uri.TryCreate(instanceUrl, UriKind.Absolute, out var uri)
            ? uri.Host.ToLowerInvariant()
            : null;
    }
}

public sealed record SessionLoginReconciliation(
    SalesforceLoginOptions ResolvedLoginOptions,
    string CanonicalHostKey,
    string OAuthTokenHost);
