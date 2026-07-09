using System;
using System.Collections.Generic;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>
/// Spike-friendly trace of sign-in and persistence decisions.
/// </summary>
public sealed class LoginFlowDiagnostics
{
    public string PreferencesPath { get; set; } = string.Empty;

    public string? LoadedTenant { get; set; }

    public string HostKey { get; set; } = string.Empty;

    /// <summary>
    /// Host key after resolving tenant/sandbox from the org instance URL.
    /// </summary>
    public string? ResolvedHostKey { get; set; }

    public string CredentialTarget { get; set; } = string.Empty;

    public string? CredentialTargetAfterSignIn { get; set; }

    public bool PreferencesFileExists { get; set; }

    public bool StoredCredentialFound { get; set; }

    public bool ForceSignInAgain { get; set; }

    public bool SilentRefreshAttempted { get; set; }

    public bool SilentRefreshSucceeded { get; set; }

    public bool InteractiveLoginUsed { get; set; }

    public bool RefreshTokenInOAuthResponse { get; set; }

    public bool RefreshTokenSaved { get; set; }

    public bool StoredCredentialAfterSignIn { get; set; }

    public string? CredentialSaveError { get; set; }

    public string? SilentRefreshFailureReason { get; set; }

    public string FormatSummary()
    {
        var lines = new List<string>
        {
            $"Preferences file: {PreferencesPath} ({(PreferencesFileExists ? "found" : "not found yet")})",
            $"Loaded tenant: {LoadedTenant ?? "(none)"}",
            $"OAuth host at sign-in: {HostKey}",
            $"Resolved host after login: {ResolvedHostKey ?? HostKey}",
            $"Credential Manager target at sign-in: {CredentialTarget}",
            $"Credential Manager target after sign-in: {CredentialTargetAfterSignIn ?? CredentialTarget}",
            $"Stored refresh token at sign-in start: {(StoredCredentialFound ? "yes" : "no")}",
            $"Sign in again selected: {(ForceSignInAgain ? "yes" : "no")}",
            $"Silent refresh attempted: {(SilentRefreshAttempted ? "yes" : "no")}",
            $"Silent refresh succeeded: {(SilentRefreshSucceeded ? "yes" : "no")}",
            $"Interactive Salesforce sign-in: {(InteractiveLoginUsed ? "yes" : "no")}",
            $"Refresh token in OAuth response: {(RefreshTokenInOAuthResponse ? "yes" : "no")}",
            $"Refresh token saved to Credential Manager: {(RefreshTokenSaved ? "yes" : "no")}",
            $"Stored refresh token after sign-in: {(StoredCredentialAfterSignIn ? "yes" : "no")}",
        };

        // Only surface when silent refresh failed and interactive sign-in did not recover.
        if (!string.IsNullOrWhiteSpace(SilentRefreshFailureReason)
            && !SilentRefreshSucceeded
            && !(InteractiveLoginUsed && StoredCredentialAfterSignIn))
        {
            lines.Add($"Silent refresh note: {SilentRefreshFailureReason}");
        }

        if (!string.IsNullOrWhiteSpace(CredentialSaveError))
        {
            lines.Add($"Credential save error: {CredentialSaveError}");
        }

        return string.Join(Environment.NewLine, lines);
    }
}
