using System;
using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Lazy on-demand session resolution before Salesforce REST calls.
/// </summary>
public sealed class SessionGate
{
    private readonly SessionContext _session;
    private readonly IUserLoginPreferencesStore _preferencesStore;
    private readonly ILoginOptionsPresenter _loginOptionsPresenter;
    private readonly SessionLoginOrchestrator _loginOrchestrator;
    private readonly SalesforceIdentityClient? _identityClient;
    private string? _lastHostKey;

    public SessionGate(
        SessionContext session,
        IUserLoginPreferencesStore preferencesStore,
        ILoginOptionsPresenter loginOptionsPresenter,
        SessionLoginOrchestrator loginOrchestrator,
        SalesforceIdentityClient? identityClient = null)
    {
        _session = session ?? throw new ArgumentNullException(nameof(session));
        _preferencesStore = preferencesStore ?? throw new ArgumentNullException(nameof(preferencesStore));
        _loginOptionsPresenter = loginOptionsPresenter
                                   ?? throw new ArgumentNullException(nameof(loginOptionsPresenter));
        _loginOrchestrator = loginOrchestrator ?? throw new ArgumentNullException(nameof(loginOrchestrator));
        _identityClient = identityClient;
    }

    public SessionLoginDiagnostics? LastDiagnostics { get; private set; }

    /// <param name="reportStatus">
    /// Optional UI hook for silent/refresh phases (e.g. Excel status bar).
    /// Interactive WebView2 OAuth has its own window; this covers the idle wait before that.
    /// </param>
    public async Task EnsureLoggedInAsync(
        CancellationToken cancellationToken = default,
        Action<string>? reportStatus = null)
    {
        if (_session.IsLoggedIn)
        {
            LastDiagnostics = null;
            if (string.IsNullOrWhiteSpace(_session.DisplayName))
            {
                reportStatus?.Invoke("Loading metadata...");
                await TryLoadDisplayNameAsync(cancellationToken).ConfigureAwait(true);
            }

            return;
        }

        SessionFlowTrace.BeginScope("EnsureLoggedInAsync");

        var prefs = _preferencesStore.TryLoadValid();
        var showOptions = prefs is null || prefs.ShowLoginOptionsOnNextUse;
        SessionFlowTrace.Log(
            $"Preferences loaded={(prefs is not null)} showLoginOptions={showOptions} preferencesPath={SalesforceRestAddinDataPaths.PreferencesFile}");

        // Debug JWT path: hand-edited jwtAccessToken skips OAuth / refresh entirely.
        if (prefs is not null && prefs.HasJwtAccessToken)
        {
            TryHydrateJwtAccessToken(prefs, diagnostics: new SessionLoginDiagnostics
            {
                PreferencesLoaded = true,
                ShowLoginOptions = false,
            });
            return;
        }

        SalesforceLoginOptions loginOptions;
        string? preferredApiVersion;
        var forceSignInAgain = false;
        var diagnostics = new SessionLoginDiagnostics
        {
            PreferencesLoaded = prefs is not null,
            ShowLoginOptions = showOptions,
        };

        if (showOptions)
        {
            diagnostics.LoginOptionsShown = true;
            SessionFlowTrace.Log("Login options dialog: showing.");
            var optionsResult = await _loginOptionsPresenter
                .ShowAsync(prefs, cancellationToken)
                .ConfigureAwait(true);

            if (optionsResult is null)
            {
                SessionFlowTrace.Log("Login options dialog: cancelled.");
                throw new SalesforceLoginCancelledException("Login options were cancelled.");
            }

            loginOptions = optionsResult.LoginOptions;
            preferredApiVersion = optionsResult.PreferredApiVersion;
            forceSignInAgain = optionsResult.ForceSignInAgain;
            diagnostics.ForceSignInAgain = forceSignInAgain;
            SessionFlowTrace.Log(
                $"Login options dialog: accepted tenant={loginOptions.Tenant} environment={loginOptions.Environment} forceSignInAgain={forceSignInAgain}");
        }
        else
        {
            loginOptions = prefs!.ToLoginOptions();
            preferredApiVersion = prefs.ApiVersion;
            SessionFlowTrace.Log(
                $"Login options dialog: skipped tenant={loginOptions.Tenant} environment={loginOptions.Environment} apiVersion={preferredApiVersion ?? "(none)"}");
        }

        var hostKey = SessionHostKey.FromLoginOptions(loginOptions);
        diagnostics.HostKey = hostKey;
        _lastHostKey = hostKey;

        var loginResult = await _loginOrchestrator
            .TryLoginAsync(loginOptions, preferredApiVersion, forceSignInAgain, cancellationToken, reportStatus)
            .ConfigureAwait(true);

        MergeDiagnostics(diagnostics, loginResult.Diagnostics);
        LastDiagnostics = diagnostics;

        if (loginResult.Cancelled)
        {
            SessionFlowTrace.Log("EnsureLoggedInAsync: cancelled.");
            throw new SalesforceLoginCancelledException("Salesforce sign-in was cancelled.");
        }

        if (!loginResult.Succeeded)
        {
            var reason = loginResult.FailureReason ?? "Salesforce sign-in failed.";
            SessionFlowTrace.Log($"EnsureLoggedInAsync: failed ({reason}).");
            throw new SalesforceLoginFailedException(reason) { Diagnostics = diagnostics };
        }

        var resolved = loginResult.ResolvedLoginOptions ?? loginOptions;
        SavePreferencesAfterLogin(resolved, preferredApiVersion, loginResult.ApiVersionResolution, prefs?.JwtAccessToken);
        await TryLoadDisplayNameAsync(cancellationToken).ConfigureAwait(true);
        SessionFlowTrace.Log(
            $"EnsureLoggedInAsync: succeeded silentRefresh={loginResult.UsedSilentRefresh} instance={_session.InstanceUrl} apiVersion={_session.ApiVersion}");
    }

    /// <summary>
    /// Clears the live session and target instance (forces login options next time).
    /// Persisted refresh tokens are kept so the same host can silent-refresh after re-selecting.
    /// Hand-edited <c>jwtAccessToken</c> in preferences is preserved.
    /// </summary>
    /// <returns>True when a live session and/or committed target was cleared.</returns>
    public bool Logout()
    {
        var prefs = _preferencesStore.TryLoadValid() ?? _preferencesStore.Load();
        var hadCommittedTarget = !string.IsNullOrWhiteSpace(prefs.Tenant)
            && !prefs.ShowLoginOptionsOnNextUse;
        var hadSomething = _session.IsLoggedIn
            || !string.IsNullOrWhiteSpace(_session.InstanceUrl)
            || hadCommittedTarget;

        _session.Invalidate();
        LastDiagnostics = null;

        _preferencesStore.Save(new UserLoginPreferences
        {
            Tenant = prefs.Tenant,
            Environment = prefs.Environment,
            SandboxId = prefs.SandboxId,
            ApiVersion = prefs.ApiVersion,
            CachedSupportedApiVersions = prefs.CachedSupportedApiVersions,
            ShowLoginOptionsOnNextUse = true,
            JwtAccessToken = prefs.JwtAccessToken,
        });

        _lastHostKey = null;
        return hadSomething;
    }

    private void TryHydrateJwtAccessToken(UserLoginPreferences prefs, SessionLoginDiagnostics diagnostics)
    {
        if (string.IsNullOrWhiteSpace(prefs.Tenant))
        {
            const string reason = "jwtAccessToken is set but tenant is missing in login preferences.";
            SessionFlowTrace.Log($"EnsureLoggedInAsync: JWT mode failed ({reason}).");
            LastDiagnostics = diagnostics;
            throw new SalesforceLoginFailedException(reason) { Diagnostics = diagnostics };
        }

        if (prefs.Environment == SalesforceEnvironment.Sandbox
            && string.IsNullOrWhiteSpace(prefs.SandboxId))
        {
            const string reason = "jwtAccessToken is set for Sandbox but sandboxId is missing in login preferences.";
            SessionFlowTrace.Log($"EnsureLoggedInAsync: JWT mode failed ({reason}).");
            LastDiagnostics = diagnostics;
            throw new SalesforceLoginFailedException(reason) { Diagnostics = diagnostics };
        }

        if (string.IsNullOrWhiteSpace(prefs.ApiVersion)
            && prefs.CachedSupportedApiVersions.Count == 0)
        {
            const string reason =
                "jwtAccessToken is set but apiVersion (or cachedSupportedApiVersions) is missing in login preferences.";
            SessionFlowTrace.Log($"EnsureLoggedInAsync: JWT mode failed ({reason}).");
            LastDiagnostics = diagnostics;
            throw new SalesforceLoginFailedException(reason) { Diagnostics = diagnostics };
        }

        var loginOptions = prefs.ToLoginOptions();
        var instanceUrl = SalesforceLoginUrlBuilder.BuildLoginBaseUrl(loginOptions);
        var hostKey = SessionHostKey.FromLoginOptions(loginOptions);
        diagnostics.HostKey = hostKey;
        _lastHostKey = hostKey;

        _session.AccessToken = prefs.JwtAccessToken!.Trim();
        _session.InstanceUrl = instanceUrl;
        _session.RefreshToken = null;
        _session.Id = null;
        _session.DisplayName = null;
        _session.IssuedAtUnixMs = null;
        _session.IsJwtAccessTokenMode = true;

        var resolution = SalesforceApiVersionSelector.Resolve(prefs.ApiVersion, prefs.CachedSupportedApiVersions);
        _session.SetApiVersion(resolution.SelectedVersion);

        // Persist committed target without clearing the hand-edited JWT.
        _preferencesStore.Save(new UserLoginPreferences
        {
            Tenant = prefs.Tenant,
            Environment = prefs.Environment,
            SandboxId = prefs.SandboxId,
            ApiVersion = prefs.ApiVersion,
            CachedSupportedApiVersions = prefs.CachedSupportedApiVersions,
            ShowLoginOptionsOnNextUse = false,
            JwtAccessToken = prefs.JwtAccessToken,
        });

        LastDiagnostics = diagnostics;
        SessionFlowTrace.Log(
            $"EnsureLoggedInAsync: JWT access token mode instance={_session.InstanceUrl} apiVersion={_session.ApiVersion}");
    }

    private async Task TryLoadDisplayNameAsync(CancellationToken cancellationToken)
    {
        if (_identityClient is null
            || string.IsNullOrWhiteSpace(_session.Id)
            || string.IsNullOrWhiteSpace(_session.AccessToken))
        {
            return;
        }

        try
        {
            var identity = await _identityClient
                .GetIdentityAsync(_session.Id!, _session.AccessToken!, cancellationToken)
                .ConfigureAwait(true);
            _session.DisplayName = identity.TryGetDisplayName();
            if (!string.IsNullOrWhiteSpace(_session.DisplayName))
            {
                SessionFlowTrace.Log($"Identity display_name loaded: {_session.DisplayName}");
            }
        }
        catch (Exception ex)
        {
            // Soft-fail: ribbon keeps instance host without user name.
            SessionFlowTrace.Log($"Identity display_name fetch failed: {ex.Message}");
        }
    }

    private void SavePreferencesAfterLogin(
        SalesforceLoginOptions loginOptions,
        string? preferredApiVersion,
        ApiVersionResolution? apiVersionResolution,
        string? preserveJwtAccessToken)
    {
        var apiVersion = apiVersionResolution?.PreferredVersionWasInvalid == true
            ? null
            : preferredApiVersion;

        _preferencesStore.Save(new UserLoginPreferences
        {
            Tenant = loginOptions.Tenant,
            Environment = loginOptions.Environment,
            SandboxId = loginOptions.SandboxId,
            ApiVersion = apiVersion,
            CachedSupportedApiVersions = apiVersionResolution?.SupportedVersions ?? Array.Empty<string>(),
            ShowLoginOptionsOnNextUse = false,
            JwtAccessToken = preserveJwtAccessToken,
        });
    }

    private static void MergeDiagnostics(SessionLoginDiagnostics target, SessionLoginDiagnostics source)
    {
        target.SilentRefreshAttempted = source.SilentRefreshAttempted;
        target.SilentRefreshSucceeded = source.SilentRefreshSucceeded;
        target.SilentRefreshFailureReason = source.SilentRefreshFailureReason;
        target.InteractiveLoginUsed = source.InteractiveLoginUsed;
        target.InteractiveFailureReason = source.InteractiveFailureReason;
        target.HostKey = source.HostKey ?? target.HostKey;
        target.StoredCredentialFound = source.StoredCredentialFound;
        target.OAuthTokenHost = source.OAuthTokenHost ?? target.OAuthTokenHost;
        if (source.ForceSignInAgain)
        {
            target.ForceSignInAgain = true;
        }
    }
}
