namespace SalesforceRestAddin.Core.OAuth;

/// <summary>
/// Builds Salesforce login/instance base URLs from tenant and environment selection.
/// </summary>
public static class SalesforceLoginUrlBuilder
{
    /// <summary>
    /// OAuth authorize/token host base URL (no trailing slash).
    /// Production tenant "kjr" → https://kjr.my.salesforce.com
    /// Sandbox tenant "kjr" + id "dev" → https://kjr--dev.sandbox.my.salesforce.com
    /// </summary>
    public static string BuildLoginBaseUrl(SalesforceLoginOptions options)
    {
        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        options = SalesforceLoginOptionsNormalizer.Normalize(options);

        var tenant = NormalizeSegment(options.Tenant);
        var sandboxId = NormalizeSegment(options.SandboxId);

        if (options.Environment == SalesforceEnvironment.Production)
        {
            return tenant is not null
                ? $"https://{tenant}.my.salesforce.com"
                : "https://login.salesforce.com";
        }

        if (tenant is not null && sandboxId is not null)
        {
            return $"https://{tenant}--{sandboxId}.sandbox.my.salesforce.com";
        }

        if (tenant is not null)
        {
            return $"https://{tenant}.sandbox.my.salesforce.com";
        }

        return "https://test.salesforce.com";
    }

    public static string BuildAuthorizeUrl(
        SalesforceLoginOptions login,
        PkcePair pkce,
        SalesforceOAuthOptions? oauth = null)
    {
        if (login is null)
        {
            throw new ArgumentNullException(nameof(login));
        }

        if (pkce is null)
        {
            throw new ArgumentNullException(nameof(pkce));
        }

        oauth ??= SalesforceOAuthOptions.Default;
        var baseUrl = BuildLoginBaseUrl(login);

        var query = new List<string>
        {
            $"client_id={Uri.EscapeDataString(oauth.ClientId)}",
            $"redirect_uri={Uri.EscapeDataString(oauth.RedirectUri)}",
            "response_type=code",
            $"code_challenge={Uri.EscapeDataString(pkce.Challenge)}",
            "code_challenge_method=S256",
            "display=popup",
            "prompt=login",
        };

        if (oauth.Scopes.Count > 0)
        {
            query.Add($"scope={Uri.EscapeDataString(string.Join(" ", oauth.Scopes))}");
        }

        return $"{baseUrl}/services/oauth2/authorize?{string.Join("&", query)}";
    }

    public static string BuildTokenUrl(SalesforceLoginOptions login) =>
        $"{BuildLoginBaseUrl(login)}/services/oauth2/token";

    private static string? NormalizeSegment(string? value)
    {
        if (value is null || string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return value.Trim().ToLowerInvariant();
    }
}
