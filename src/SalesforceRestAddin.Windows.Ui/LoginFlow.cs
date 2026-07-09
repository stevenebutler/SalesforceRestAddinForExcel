using System;
using System.IO;
using System.Net.Http;
using System.Windows;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.Net;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class LoginFlowResult
{
    public required SalesforceLoginOptions LoginOptions { get; init; }

    public required string LoginBaseUrl { get; init; }

    public required LoginFlowDiagnostics Diagnostics { get; init; }

    public string? AuthorizeUrl { get; init; }

    public PkcePair? Pkce { get; init; }

    public string? AuthorizationCode { get; init; }

    public OAuthTokenResponse? TokenResponse { get; init; }

    public string? IdentitySummary { get; init; }

    public string? ErrorMessage { get; init; }

    public bool UsedSavedSession { get; init; }

    public bool Succeeded => TokenResponse is not null && SessionContext.Current.IsLoggedIn;
}

/// <summary>
/// Legacy "login first" composition root for spike diagnostics and one-shot sign-in flows.
/// XLL API entry points use <see cref="SessionGate"/> at call time instead.
/// </summary>
public static class LoginFlow
{
    public static LoginFlowResult? Run(
        Window? owner = null,
        HttpClient? http = null,
        ISessionCredentialStore? credentialStore = null,
        IUserLoginPreferencesStore? preferencesStore = null)
    {
        credentialStore ??= new WindowsCredentialSessionStore();
        preferencesStore ??= new JsonUserLoginPreferencesStore(SalesforceRestAddinDataPaths.PreferencesFile);
        var preferencesPath = preferencesStore is JsonUserLoginPreferencesStore jsonStore
            ? jsonStore.FilePath
            : SalesforceRestAddinDataPaths.PreferencesFile;

        var preferences = preferencesStore.Load();
        var optionsWindow = new LoginOptionsWindow(preferences, credentialStore);
        if (owner is not null)
        {
            optionsWindow.Owner = owner;
        }

        if (optionsWindow.ShowDialog() != true)
        {
            return null;
        }

        var dialogResult = optionsWindow.GetResult();
        if (dialogResult is null)
        {
            return null;
        }

        var loginOptions = SalesforceLoginOptionsNormalizer.Normalize(dialogResult.LoginOptions);
        var preferredApiVersion = dialogResult.PreferredApiVersion;
        preferencesStore.Save(ToPreferences(loginOptions, preferredApiVersion, preferences.CachedSupportedApiVersions));

        var hostKey = SessionHostKey.FromLoginOptions(loginOptions);
        var diagnostics = new LoginFlowDiagnostics
        {
            PreferencesPath = preferencesPath,
            PreferencesFileExists = File.Exists(preferencesPath),
            LoadedTenant = preferences.Tenant,
            HostKey = hostKey,
            CredentialTarget = WindowsCredentialSessionStore.TargetNameFor(hostKey),
            StoredCredentialFound = credentialStore.Load(hostKey) is not null,
            ForceSignInAgain = dialogResult.ForceSignInAgain,
        };

        var oauth = SalesforceOAuthOptions.Default;
        var loginBaseUrl = SalesforceLoginUrlBuilder.BuildLoginBaseUrl(loginOptions);
        using var httpClient = http ?? SalesforceHttpClientFactory.Create();
        var authenticator = new SessionAuthenticator(
            credentialStore,
            new SalesforceOAuthClient(httpClient),
            oauth);
        var orchestrator = new SessionLoginOrchestrator(
            SessionContext.Current,
            authenticator,
            new WindowsInteractiveLoginHandler(authenticator, oauth));

        try
        {
            var loginResult = orchestrator
                .TryLoginAsync(loginOptions, preferredApiVersion, dialogResult.ForceSignInAgain)
                .GetAwaiter()
                .GetResult();

            CopyDiagnostics(diagnostics, loginResult.Diagnostics);

            if (loginResult.Cancelled)
            {
                return null;
            }

            if (!loginResult.Succeeded || loginResult.Token is null)
            {
                return FailedResult(
                    loginOptions,
                    loginBaseUrl,
                    diagnostics,
                    loginResult.FailureReason ?? "Sign-in failed.");
            }

            loginOptions = loginResult.ResolvedLoginOptions ?? loginOptions;
            hostKey = SessionHostKey.FromLoginOptions(loginOptions);
            loginBaseUrl = SalesforceLoginUrlBuilder.BuildLoginBaseUrl(loginOptions);
            UpdateResolvedHostDiagnostics(diagnostics, hostKey);

            SavePreferencesAfterLogin(
                preferencesStore,
                loginOptions,
                loginResult.ApiVersionResolution,
                preferredApiVersion);

            diagnostics.RefreshTokenInOAuthResponse = !string.IsNullOrWhiteSpace(loginResult.Token.RefreshToken);
            diagnostics.RefreshTokenSaved = credentialStore.Load(hostKey) is not null;
            FinalizeStoredCredentialState(diagnostics, credentialStore, hostKey);
            var identitySummary = FetchIdentitySummary(httpClient, SessionContext.Current);

            return new LoginFlowResult
            {
                LoginOptions = loginOptions,
                LoginBaseUrl = loginBaseUrl,
                Diagnostics = diagnostics,
                TokenResponse = loginResult.Token,
                IdentitySummary = identitySummary,
                UsedSavedSession = loginResult.UsedSilentRefresh,
            };
        }
        catch (Exception ex)
        {
            return FailedResult(loginOptions, loginBaseUrl, diagnostics, ex.Message);
        }
    }

    private static void CopyDiagnostics(LoginFlowDiagnostics target, SessionLoginDiagnostics source)
    {
        target.SilentRefreshAttempted = source.SilentRefreshAttempted;
        target.SilentRefreshSucceeded = source.SilentRefreshSucceeded;
        target.SilentRefreshFailureReason = source.SilentRefreshFailureReason;
        target.InteractiveLoginUsed = source.InteractiveLoginUsed;
        target.ForceSignInAgain = source.ForceSignInAgain;
    }

    private static LoginFlowResult FailedResult(
        SalesforceLoginOptions loginOptions,
        string loginBaseUrl,
        LoginFlowDiagnostics diagnostics,
        string errorMessage,
        string? authorizeUrl = null,
        PkcePair? pkce = null,
        string? authorizationCode = null) =>
        new()
        {
            LoginOptions = loginOptions,
            LoginBaseUrl = loginBaseUrl,
            Diagnostics = diagnostics,
            AuthorizeUrl = authorizeUrl,
            Pkce = pkce,
            AuthorizationCode = authorizationCode,
            ErrorMessage = errorMessage,
        };

    private static UserLoginPreferences ToPreferences(
        SalesforceLoginOptions loginOptions,
        string? apiVersion = null,
        System.Collections.Generic.IReadOnlyList<string>? cachedSupportedApiVersions = null) => new()
    {
        Tenant = loginOptions.Tenant,
        Environment = loginOptions.Environment,
        SandboxId = loginOptions.SandboxId,
        ApiVersion = apiVersion,
        CachedSupportedApiVersions = cachedSupportedApiVersions ?? Array.Empty<string>(),
    };

    private static void SavePreferencesAfterLogin(
        IUserLoginPreferencesStore preferencesStore,
        SalesforceLoginOptions loginOptions,
        ApiVersionResolution? apiVersionResolution,
        string? preferredApiVersion)
    {
        var apiVersion = apiVersionResolution?.PreferredVersionWasInvalid == true
            ? null
            : preferredApiVersion;

        preferencesStore.Save(new UserLoginPreferences
        {
            Tenant = loginOptions.Tenant,
            Environment = loginOptions.Environment,
            SandboxId = loginOptions.SandboxId,
            ApiVersion = apiVersion,
            CachedSupportedApiVersions = apiVersionResolution?.SupportedVersions ?? Array.Empty<string>(),
            ShowLoginOptionsOnNextUse = false,
        });
    }

    private static string? FetchIdentitySummary(HttpClient httpClient, SessionContext session)
    {
        var identityUrl = session.Id;
        var accessToken = session.AccessToken;
        if (identityUrl is null || accessToken is null
            || string.IsNullOrWhiteSpace(identityUrl)
            || string.IsNullOrWhiteSpace(accessToken))
        {
            return null;
        }

        try
        {
            var identity = new SalesforceIdentityClient(httpClient)
                .GetIdentityAsync(identityUrl, accessToken)
                .GetAwaiter()
                .GetResult();
            return identity.FormatSummary();
        }
        catch (Exception ex)
        {
            return $"Salesforce identity request failed: {ex.Message}";
        }
    }

    private static void UpdateResolvedHostDiagnostics(LoginFlowDiagnostics diagnostics, string resolvedHostKey)
    {
        diagnostics.ResolvedHostKey = resolvedHostKey;
        diagnostics.CredentialTargetAfterSignIn = WindowsCredentialSessionStore.TargetNameFor(resolvedHostKey);
    }

    private static void FinalizeStoredCredentialState(
        LoginFlowDiagnostics diagnostics,
        ISessionCredentialStore credentialStore,
        string hostKey)
    {
        diagnostics.StoredCredentialAfterSignIn = credentialStore.Load(hostKey) is not null;
    }
}
