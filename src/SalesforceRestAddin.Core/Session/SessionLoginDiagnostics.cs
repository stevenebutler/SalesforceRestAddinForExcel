namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Trace of sign-in decisions for spike UI and tests.
/// </summary>
public sealed class SessionLoginDiagnostics
{
    public bool LoginOptionsShown { get; set; }

    public bool ForceSignInAgain { get; set; }

    public bool SilentRefreshAttempted { get; set; }

    public bool SilentRefreshSucceeded { get; set; }

    public bool InteractiveLoginUsed { get; set; }

    public string? SilentRefreshFailureReason { get; set; }

    public string? InteractiveFailureReason { get; set; }

    public string? HostKey { get; set; }

    public bool StoredCredentialFound { get; set; }

    public string? OAuthTokenHost { get; set; }

    public bool PreferencesLoaded { get; set; }

    public bool ShowLoginOptions { get; set; }
}
