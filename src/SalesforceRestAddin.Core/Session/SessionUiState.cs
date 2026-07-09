using System;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Ribbon label and enablement derived from login preferences + live session.
/// </summary>
public sealed class SessionUiState
{
    public bool HasTarget { get; init; }

    public bool IsLoggedIn { get; init; }

    public bool LogoutEnabled => HasTarget;

    public bool OptionsEnabled => true;

    public required string GroupLabel { get; init; }

    public string? InstanceHost { get; init; }

    public string? DisplayName { get; init; }

    public static SessionUiState From(IUserLoginPreferencesStore preferencesStore, SessionContext session)
    {
        if (preferencesStore is null)
        {
            throw new ArgumentNullException(nameof(preferencesStore));
        }

        if (session is null)
        {
            throw new ArgumentNullException(nameof(session));
        }

        var prefs = preferencesStore.TryLoadValid();
        var committedTarget = prefs is not null && !prefs.ShowLoginOptionsOnNextUse;
        var liveHost = TryHostFromInstanceUrl(session.InstanceUrl);
        var hasTarget = committedTarget || liveHost is not null;

        string? host = null;
        if (liveHost is not null && (session.IsLoggedIn || !committedTarget))
        {
            // Prefer live instance host when logged in (or when that is the only target signal).
            host = liveHost;
        }
        else if (committedTarget)
        {
            host = SessionHostKey.FromLoginOptions(prefs!.ToLoginOptions());
        }

        var displayName = session.IsLoggedIn ? session.DisplayName : null;
        var label = FormatGroupLabel(host, session.IsLoggedIn, displayName);

        return new SessionUiState
        {
            HasTarget = hasTarget,
            IsLoggedIn = session.IsLoggedIn,
            GroupLabel = label,
            InstanceHost = host,
            DisplayName = displayName,
        };
    }

    public static string FormatGroupLabel(
        string? instanceHost,
        bool isLoggedIn,
        string? displayName)
    {
        var instance = string.IsNullOrWhiteSpace(instanceHost)
            ? "(none)"
            : instanceHost!.Trim();
        string user;
        if (!isLoggedIn)
        {
            user = "(not signed in)";
        }
        else if (string.IsNullOrWhiteSpace(displayName))
        {
            user = "(unknown)";
        }
        else
        {
            user = displayName!.Trim();
        }

        return $"Instance: {instance}  User: {user}";
    }

    private static string? TryHostFromInstanceUrl(string? instanceUrl)
    {
        if (string.IsNullOrWhiteSpace(instanceUrl))
        {
            return null;
        }

        if (!Uri.TryCreate(instanceUrl, UriKind.Absolute, out var uri))
        {
            return null;
        }

        return string.IsNullOrWhiteSpace(uri.Host) ? null : uri.Host;
    }
}
