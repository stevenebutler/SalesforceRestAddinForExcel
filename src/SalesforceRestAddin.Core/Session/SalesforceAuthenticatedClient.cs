using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Sends HTTP requests with bearer auth and mid-call 401 recovery (refresh → interactive → retry).
/// Does not show login options — only <see cref="IInteractiveLoginHandler"/>.
/// </summary>
public sealed class SalesforceAuthenticatedClient
{
    private readonly HttpClient _http;
    private readonly SessionContext _session;
    private readonly IUserLoginPreferencesStore _preferencesStore;
    private readonly SessionAccessTokenRefresher _tokenRefresher;
    private readonly SessionLoginOrchestrator _loginOrchestrator;
    private readonly IAuthRequestMutator _authMutator;
    private readonly ISessionRecoveryTracer _recoveryTracer;

    public SalesforceAuthenticatedClient(
        HttpClient http,
        SessionContext session,
        IUserLoginPreferencesStore preferencesStore,
        SessionAccessTokenRefresher tokenRefresher,
        SessionLoginOrchestrator loginOrchestrator,
        IAuthRequestMutator? authMutator = null,
        ISessionRecoveryTracer? recoveryTracer = null)
    {
        _http = http ?? throw new ArgumentNullException(nameof(http));
        _session = session ?? throw new ArgumentNullException(nameof(session));
        _preferencesStore = preferencesStore ?? throw new ArgumentNullException(nameof(preferencesStore));
        _tokenRefresher = tokenRefresher ?? throw new ArgumentNullException(nameof(tokenRefresher));
        _loginOrchestrator = loginOrchestrator ?? throw new ArgumentNullException(nameof(loginOrchestrator));
        _authMutator = authMutator ?? NullAuthRequestMutator.Instance;
        _recoveryTracer = recoveryTracer ?? NullSessionRecoveryTracer.Instance;
    }

    public async Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken = default)
    {
        if (request is null)
        {
            throw new ArgumentNullException(nameof(request));
        }

        if (_recoveryTracer is SessionRecoveryTrace recoveryTrace)
        {
            recoveryTrace.Clear();
        }

        if (_session.IsJwtAccessTokenMode)
        {
            if (_session.IsExpired())
            {
                throw new SessionRecoveryException(
                    "jwtAccessToken session cannot be refreshed. Update or remove jwtAccessToken in login preferences.");
            }

            var jwtAttempt = await SendOnceAsync(request, isRetryAfterRecovery: false, cancellationToken)
                .ConfigureAwait(false);
            if (!SalesforceApiErrors.IsSessionAuthFailure(jwtAttempt))
            {
                _recoveryTracer.AddStep(
                    $"REST attempt succeeded ({(int)jwtAttempt.StatusCode}); JWT mode (no refresh).");
                return jwtAttempt;
            }

            var jwtStatus = (int)jwtAttempt.StatusCode;
            var jwtBody = await jwtAttempt.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            jwtAttempt.Dispose();
            _recoveryTracer.AddStep($"REST auth failure in JWT mode ({jwtStatus}): {SummarizeBody(jwtBody)}");
            throw new SessionRecoveryException(
                $"jwtAccessToken was rejected by Salesforce and cannot be refreshed: {jwtBody}");
        }

        if (_session.IsExpired())
        {
            _recoveryTracer.AddStep("Access token expired (heuristic); refreshing before REST call.");
            await RefreshOrEscalateInteractiveAsync(cancellationToken).ConfigureAwait(false);
        }

        var firstAttempt = await SendOnceAsync(request, isRetryAfterRecovery: false, cancellationToken)
            .ConfigureAwait(false);
        if (!SalesforceApiErrors.IsSessionAuthFailure(firstAttempt))
        {
            _recoveryTracer.AddStep(
                $"REST attempt succeeded ({(int)firstAttempt.StatusCode}); no mid-call session recovery needed.");
            return firstAttempt;
        }

        var firstStatusCode = (int)firstAttempt.StatusCode;
        var body = await firstAttempt.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        firstAttempt.Dispose();
        _recoveryTracer.AddStep(
            $"REST auth failure ({firstStatusCode}): {SummarizeBody(body)}");

        _recoveryTracer.AddStep("Attempting silent refresh from stored credentials.");
        if (!await RecoverSessionAsync(cancellationToken).ConfigureAwait(false))
        {
            throw new SessionRecoveryException(
                $"Salesforce session recovery failed after auth error: {body}");
        }

        _recoveryTracer.AddStep("Retrying REST call with refreshed access token.");
        var retryAttempt = await SendOnceAsync(CloneRequest(request), isRetryAfterRecovery: true, cancellationToken)
            .ConfigureAwait(false);
        if (SalesforceApiErrors.IsSessionAuthFailure(retryAttempt))
        {
            var retryBody = await retryAttempt.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            retryAttempt.Dispose();
            _recoveryTracer.AddStep($"REST retry still rejected: {SummarizeBody(retryBody)}");
            throw new SessionRecoveryException(
                $"Salesforce session recovery exhausted; auth error persisted: {retryBody}");
        }

        _recoveryTracer.AddStep($"REST retry succeeded ({(int)retryAttempt.StatusCode}).");
        return retryAttempt;
    }

    private static string SummarizeBody(string body) =>
        body.Length <= 160 ? body : body[..160] + "...";

    private async Task<HttpResponseMessage> SendOnceAsync(
        HttpRequestMessage request,
        bool isRetryAfterRecovery,
        CancellationToken cancellationToken)
    {
        var accessToken = _session.AccessToken;
        if (accessToken is null || string.IsNullOrWhiteSpace(accessToken))
        {
            throw new InvalidOperationException("No Salesforce access token is available for this request.");
        }

        var outbound = CloneRequest(request);
        outbound.Headers.Authorization = new AuthenticationHeaderValue(
            "Bearer",
            _authMutator.MutateBearer(accessToken, isRetryAfterRecovery));

        return await _http.SendAsync(outbound, cancellationToken).ConfigureAwait(false);
    }

    private async Task<bool> RecoverSessionAsync(CancellationToken cancellationToken)
    {
        var loginOptions = GetLoginOptionsForRecovery();
        if (await _tokenRefresher.TryRefreshAsync(loginOptions, _session, cancellationToken).ConfigureAwait(false))
        {
            return true;
        }

        return await EscalateInteractiveAsync(loginOptions, cancellationToken).ConfigureAwait(false);
    }

    private async Task RefreshOrEscalateInteractiveAsync(CancellationToken cancellationToken)
    {
        var loginOptions = GetLoginOptionsForRecovery();
        if (await _tokenRefresher.TryRefreshAsync(loginOptions, _session, cancellationToken).ConfigureAwait(false))
        {
            return;
        }

        if (!await EscalateInteractiveAsync(loginOptions, cancellationToken).ConfigureAwait(false))
        {
            throw new SessionRecoveryException("Proactive token refresh failed and interactive recovery did not succeed.");
        }
    }

    private async Task<bool> EscalateInteractiveAsync(
        SalesforceLoginOptions loginOptions,
        CancellationToken cancellationToken)
    {
        _recoveryTracer.AddStep("Escalating to interactive OAuth (WebView).");
        var result = await _loginOrchestrator
            .TryLoginAsync(loginOptions, GetPreferredApiVersion(), forceSignInAgain: true, cancellationToken)
            .ConfigureAwait(false);

        if (result.Cancelled)
        {
            throw new SalesforceLoginCancelledException("Salesforce sign-in was cancelled during session recovery.");
        }

        if (result.Succeeded)
        {
            _recoveryTracer.AddStep("Interactive OAuth succeeded.");
        }

        return result.Succeeded;
    }

    private SalesforceLoginOptions GetLoginOptionsForRecovery()
    {
        var prefs = _preferencesStore.TryLoadValid();
        if (prefs is not null)
        {
            return prefs.ToLoginOptions();
        }

        var fromInstance = SalesforceInstanceHost.TryParseInstanceUrl(_session.InstanceUrl);
        if (fromInstance is not null)
        {
            return fromInstance.ToLoginOptions();
        }

        return new SalesforceLoginOptions();
    }

    private string? GetPreferredApiVersion() => _preferencesStore.TryLoadValid()?.ApiVersion;

    private static HttpRequestMessage CloneRequest(HttpRequestMessage request)
    {
        var clone = new HttpRequestMessage(request.Method, request.RequestUri);

        if (request.Content is not null)
        {
            clone.Content = request.Content;
        }

        foreach (var header in request.Headers)
        {
            clone.Headers.TryAddWithoutValidation(header.Key, header.Value);
        }

        return clone;
    }
}
