namespace SalesforceRestAddin.Core.OAuth;

/// <summary>
/// OAuth connected-app settings. Defaults match the Salesforce CLI global connected app
/// so users can sign in without creating their own Connected App.
/// </summary>
public sealed class SalesforceOAuthOptions
{
    /// <summary>
    /// Salesforce CLI / VS Code global connected app (PKCE, no secret required).
    /// </summary>
    public const string DefaultClientId = "PlatformCLI";

    /// <summary>
    /// Legacy Force.com Connector and SFDX loopback callback path.
    /// </summary>
    public const string DefaultRedirectUri = "http://localhost:1717/OauthRedirect";

    public string ClientId { get; init; } = DefaultClientId;

    public string RedirectUri { get; init; } = DefaultRedirectUri;

    /// <summary>
    /// Optional client secret for custom Connected Apps that are not PKCE-only.
    /// Do not hard-code secrets in source; load from user config when needed.
    /// </summary>
    public string? ClientSecret { get; init; }

    public IReadOnlyList<string> Scopes { get; init; } = ["api", "refresh_token", "offline_access"];

    public static SalesforceOAuthOptions Default { get; } = new();
}
