using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Net;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Host-scoped OAuth session restore, refresh, and persistence.
/// </summary>
public sealed class SessionAuthenticator
{
    private readonly ISessionCredentialStore _credentialStore;
    private readonly SalesforceOAuthClient _oauthClient;
    private readonly SalesforceApiVersionResolver _apiVersionResolver;
    private readonly SalesforceOAuthOptions _oauth;

    public SessionAuthenticator(
        ISessionCredentialStore credentialStore,
        SalesforceOAuthClient oauthClient,
        SalesforceOAuthOptions? oauth = null,
        SalesforceApiVersionResolver? apiVersionResolver = null)
    {
        _credentialStore = credentialStore ?? throw new ArgumentNullException(nameof(credentialStore));
        _oauthClient = oauthClient ?? throw new ArgumentNullException(nameof(oauthClient));
        _oauth = oauth ?? SalesforceOAuthOptions.Default;
        _apiVersionResolver = apiVersionResolver ?? new SalesforceApiVersionResolver(_oauthClient.HttpClient);
    }

    public async Task<ApiVersionResolution> ApplyToSessionAsync(
        OAuthTokenResponse token,
        SessionContext session,
        string? preferredApiVersion = null,
        CancellationToken cancellationToken = default)
    {
        if (token is null)
        {
            throw new ArgumentNullException(nameof(token));
        }

        if (session is null)
        {
            throw new ArgumentNullException(nameof(session));
        }

        token.ApplyTo(session);
        var resolution = await _apiVersionResolver.ResolveForSessionAsync(
                token.InstanceUrl,
                token.AccessToken,
                preferredApiVersion,
                cancellationToken)
            .ConfigureAwait(false);
        session.SetApiVersion(resolution.SelectedVersion);
        return resolution;
    }

    public bool HasStoredRefreshToken(SalesforceLoginOptions loginOptions)
    {
        var normalized = SalesforceLoginOptionsNormalizer.Normalize(loginOptions);
        var hostKey = SessionHostKey.FromLoginOptions(normalized);
        var stored = _credentialStore.Load(hostKey);
        return stored is not null && !string.IsNullOrWhiteSpace(stored.RefreshToken);
    }

    public async Task<(OAuthTokenResponse? Token, string? FailureReason)> TryRestoreSessionAsync(
        SalesforceLoginOptions loginOptions,
        IAuthRequestMutator? authMutator = null,
        ISessionRecoveryTracer? recoveryTracer = null,
        CancellationToken cancellationToken = default)
    {
        var normalized = SalesforceLoginOptionsNormalizer.Normalize(loginOptions);
        var hostKey = SessionHostKey.FromLoginOptions(normalized);
        var stored = _credentialStore.Load(hostKey);
        if (stored is null || string.IsNullOrWhiteSpace(stored.RefreshToken))
        {
            recoveryTracer?.AddStep("Silent refresh skipped: no saved refresh token for this host.");
            SessionFlowTrace.Log($"Silent refresh: no stored credential for hostKey={hostKey}.");
            return (null, "No saved refresh token for this host.");
        }

        var oauthTokenHost = stored.OAuthTokenHost ?? hostKey;
        SessionFlowTrace.Log(
            $"Silent refresh: stored credential found hostKey={hostKey} oauthTokenHost={oauthTokenHost} refreshTokenLength={stored.RefreshToken.Length}");

        var refreshToken = authMutator is null
            ? stored.RefreshToken
            : authMutator.MutateRefreshToken(stored.RefreshToken);
        var refreshCorrupted = refreshToken.StartsWith("CORRUPTED:", StringComparison.Ordinal);

        try
        {
            TlsProtocolBootstrap.EnsureEnabled();
            recoveryTracer?.AddStep(
                $"POST https://{oauthTokenHost}/services/oauth2/token (grant_type=refresh_token, refresh_corrupted={refreshCorrupted})");
            SessionFlowTrace.Log($"Silent refresh: POST https://{oauthTokenHost}/services/oauth2/token");

            var token = await _oauthClient.RefreshAccessTokenAsync(
                    oauthTokenHost,
                    refreshToken,
                    _oauth,
                    cancellationToken)
                .ConfigureAwait(false);

            var reconciliation = SessionLoginReconciler.Reconcile(normalized, token);
            var persistResult = PersistReconciled(
                reconciliation,
                token,
                token.RefreshToken ?? stored.RefreshToken,
                deleteHosts: [hostKey]);

            if (!persistResult.Saved && !string.IsNullOrWhiteSpace(persistResult.ErrorMessage))
            {
                recoveryTracer?.AddStep($"refresh_token grant returned tokens but credential save failed: {persistResult.ErrorMessage}");
                return (null, persistResult.ErrorMessage);
            }

            recoveryTracer?.AddStep("refresh_token grant succeeded; access token updated in session.");
            SessionFlowTrace.Log("Silent refresh: token response received.");
            return (token, null);
        }
        catch (HttpRequestException ex)
        {
            recoveryTracer?.AddStep($"refresh_token grant network error: {ex.Message}");
            SessionFlowTrace.LogException("Silent refresh: HTTP request failed", ex);
            return (null, ex.Message);
        }
        catch (InvalidOperationException ex)
        {
            recoveryTracer?.AddStep($"refresh_token grant failed: {ex.Message}");
            SessionFlowTrace.Log($"Silent refresh: Salesforce rejected refresh ({ex.Message}).");
            _credentialStore.Delete(hostKey);
            return (null, ex.Message);
        }
    }

    public AuthorizationCodeLoginResult CompleteAuthorizationCodeLogin(
        SalesforceLoginOptions loginOptions,
        string authorizationCode,
        string codeVerifier)
    {
        var normalized = SalesforceLoginOptionsNormalizer.Normalize(loginOptions);
        var oauthHostBefore = SessionHostKey.FromLoginOptions(normalized);

        TlsProtocolBootstrap.EnsureEnabled();
        var token = _oauthClient.ExchangeAuthorizationCodeAsync(
                normalized,
                authorizationCode,
                codeVerifier,
                _oauth)
            .GetAwaiter()
            .GetResult();

        var reconciliation = SessionLoginReconciler.Reconcile(normalized, token);
        var persistResult = PersistReconciled(
            reconciliation,
            token,
            token.RefreshToken,
            deleteHosts: [oauthHostBefore]);

        return new AuthorizationCodeLoginResult(token, persistResult, reconciliation.ResolvedLoginOptions);
    }

    public SessionPersistResult TryPersist(
        string hostKey,
        OAuthTokenResponse token,
        string? fallbackRefreshToken = null) =>
        PersistReconciled(
            new SessionLoginReconciliation(
                ResolvedLoginOptions: new SalesforceLoginOptions(),
                CanonicalHostKey: hostKey,
                OAuthTokenHost: hostKey),
            token,
            token.RefreshToken ?? fallbackRefreshToken,
            deleteHosts: []);

    private SessionPersistResult PersistReconciled(
        SessionLoginReconciliation reconciliation,
        OAuthTokenResponse token,
        string? refreshToken,
        IReadOnlyList<string> deleteHosts)
    {
        if (refreshToken is null || string.IsNullOrWhiteSpace(refreshToken))
        {
            return SessionPersistResult.NoRefreshTokenInResponse();
        }

        try
        {
            _credentialStore.Save(new StoredSessionCredentials
            {
                HostKey = reconciliation.CanonicalHostKey,
                OAuthTokenHost = reconciliation.OAuthTokenHost,
                RefreshToken = refreshToken,
                InstanceUrl = token.InstanceUrl,
                Id = token.Id,
            });

            foreach (var hostToDelete in deleteHosts)
            {
                if (!string.Equals(hostToDelete, reconciliation.CanonicalHostKey, StringComparison.OrdinalIgnoreCase))
                {
                    _credentialStore.Delete(hostToDelete);
                }
            }

            return SessionPersistResult.SavedSuccessfully();
        }
        catch (Exception ex)
        {
            return SessionPersistResult.Failed(ex.Message);
        }
    }
}

public sealed record AuthorizationCodeLoginResult(
    OAuthTokenResponse Token,
    SessionPersistResult PersistResult,
    SalesforceLoginOptions ResolvedLoginOptions);
