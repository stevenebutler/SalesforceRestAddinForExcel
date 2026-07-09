using System;
using System.Text;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Windows.Ui;

public static class SessionLoginDiagnosticsText
{
    public static string Format(SessionLoginDiagnostics? diagnostics)
    {
        if (diagnostics is null)
        {
            return "(no gate diagnostics recorded)";
        }

        var builder = new StringBuilder();
        builder.AppendLine($"Preferences loaded: {diagnostics.PreferencesLoaded}");
        builder.AppendLine($"Show login options: {diagnostics.ShowLoginOptions}");
        builder.AppendLine($"Login options shown: {diagnostics.LoginOptionsShown}");
        builder.AppendLine($"Force sign-in again: {diagnostics.ForceSignInAgain}");
        builder.AppendLine($"Host key: {diagnostics.HostKey ?? "(none)"}");
        builder.AppendLine($"Stored credential found: {diagnostics.StoredCredentialFound}");
        builder.AppendLine($"OAuth token host: {diagnostics.OAuthTokenHost ?? "(none)"}");
        builder.AppendLine($"Silent refresh attempted: {diagnostics.SilentRefreshAttempted}");
        builder.AppendLine($"Silent refresh succeeded: {diagnostics.SilentRefreshSucceeded}");
        if (!string.IsNullOrWhiteSpace(diagnostics.SilentRefreshFailureReason))
        {
            builder.AppendLine($"Silent refresh note: {diagnostics.SilentRefreshFailureReason}");
        }

        builder.AppendLine($"Interactive login used: {diagnostics.InteractiveLoginUsed}");
        if (!string.IsNullOrWhiteSpace(diagnostics.InteractiveFailureReason))
        {
            builder.AppendLine($"Interactive failure: {diagnostics.InteractiveFailureReason}");
        }

        return builder.ToString().TrimEnd();
    }
}
