using System;
using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Core.Session;

public sealed class SessionLoginAttemptResult
{
    public bool Succeeded { get; init; }

    public bool Cancelled { get; init; }

    public bool UsedSilentRefresh { get; init; }

    public string? FailureReason { get; init; }

    public OAuthTokenResponse? Token { get; init; }

    public SalesforceLoginOptions? ResolvedLoginOptions { get; init; }

    public ApiVersionResolution? ApiVersionResolution { get; init; }

    public SessionLoginDiagnostics Diagnostics { get; init; } = new();
}

/// <summary>
/// Silent refresh first, then interactive OAuth — no UI except via injected handlers.
/// </summary>
public sealed class SessionLoginOrchestrator
{
    private readonly SessionContext _session;
    private readonly SessionAuthenticator _authenticator;
    private readonly IInteractiveLoginHandler _interactiveLoginHandler;
    private readonly IAuthRequestMutator? _authMutator;

    public SessionLoginOrchestrator(
        SessionContext session,
        SessionAuthenticator authenticator,
        IInteractiveLoginHandler interactiveLoginHandler,
        IAuthRequestMutator? authMutator = null)
    {
        _session = session ?? throw new ArgumentNullException(nameof(session));
        _authenticator = authenticator ?? throw new ArgumentNullException(nameof(authenticator));
        _interactiveLoginHandler = interactiveLoginHandler
                                   ?? throw new ArgumentNullException(nameof(interactiveLoginHandler));
        _authMutator = authMutator;
    }

    public async Task<SessionLoginAttemptResult> TryLoginAsync(
        SalesforceLoginOptions loginOptions,
        string? preferredApiVersion,
        bool forceSignInAgain,
        CancellationToken cancellationToken = default,
        Action<string>? reportStatus = null)
    {
        var diagnostics = new SessionLoginDiagnostics
        {
            ForceSignInAgain = forceSignInAgain,
        };

        loginOptions = SalesforceLoginOptionsNormalizer.Normalize(loginOptions);
        diagnostics.HostKey = SessionHostKey.FromLoginOptions(loginOptions);
        SessionFlowTrace.Log(
            $"TryLoginAsync tenant={loginOptions.Tenant} environment={loginOptions.Environment} hostKey={diagnostics.HostKey} forceSignInAgain={forceSignInAgain}");

        if (!forceSignInAgain)
        {
            diagnostics.SilentRefreshAttempted = true;
            diagnostics.StoredCredentialFound = _authenticator.HasStoredRefreshToken(loginOptions);
            diagnostics.OAuthTokenHost = diagnostics.StoredCredentialFound
                ? SessionHostKey.FromLoginOptions(loginOptions)
                : null;

            if (diagnostics.StoredCredentialFound)
            {
                reportStatus?.Invoke("Refreshing session...");
            }

            SessionFlowTrace.Log("Silent refresh: attempting restore from stored credentials.");

            try
            {
                var (restored, failureReason) = await _authenticator
                    .TryRestoreSessionAsync(loginOptions, _authMutator, cancellationToken: cancellationToken)
                    .ConfigureAwait(true);

                if (restored is not null)
                {
                    reportStatus?.Invoke("Loading metadata...");
                    var resolution = await _authenticator
                        .ApplyToSessionAsync(restored, _session, preferredApiVersion, cancellationToken)
                        .ConfigureAwait(true);
                    var resolved = SessionLoginReconciler.Reconcile(loginOptions, restored).ResolvedLoginOptions;
                    diagnostics.SilentRefreshSucceeded = true;
                    SessionFlowTrace.Log("Silent refresh: succeeded.");

                    return new SessionLoginAttemptResult
                    {
                        Succeeded = true,
                        UsedSilentRefresh = true,
                        Token = restored,
                        ResolvedLoginOptions = resolved,
                        ApiVersionResolution = resolution,
                        Diagnostics = diagnostics,
                    };
                }

                diagnostics.SilentRefreshFailureReason = failureReason;
                SessionFlowTrace.Log($"Silent refresh: no session restored ({failureReason ?? "unknown"}).");
            }
            catch (Exception ex)
            {
                diagnostics.SilentRefreshFailureReason = ex.Message;
                SessionFlowTrace.LogException("Silent refresh: unexpected exception — escalating to interactive OAuth", ex);
            }
        }
        else
        {
            SessionFlowTrace.Log("Silent refresh: skipped (force sign-in again).");
        }

        diagnostics.InteractiveLoginUsed = true;
        reportStatus?.Invoke("Signing in...");
        SessionFlowTrace.Log("Interactive OAuth: starting.");

        var interactive = await _interactiveLoginHandler
            .TrySignInAsync(loginOptions, preferredApiVersion, cancellationToken)
            .ConfigureAwait(true);

        if (interactive.Cancelled)
        {
            SessionFlowTrace.Log("Interactive OAuth: cancelled by user.");
            return new SessionLoginAttemptResult
            {
                Cancelled = true,
                Diagnostics = diagnostics,
            };
        }

        if (!interactive.Succeeded || interactive.Token is null)
        {
            diagnostics.InteractiveFailureReason = interactive.ErrorMessage;
            SessionFlowTrace.Log($"Interactive OAuth: failed ({interactive.ErrorMessage ?? "unknown"}).");
            return new SessionLoginAttemptResult
            {
                FailureReason = interactive.ErrorMessage ?? "Interactive sign-in failed.",
                Diagnostics = diagnostics,
            };
        }

        SessionFlowTrace.Log("Interactive OAuth: token received; applying to session.");
        reportStatus?.Invoke("Loading metadata...");
        var apiResolution = await _authenticator
            .ApplyToSessionAsync(interactive.Token, _session, preferredApiVersion, cancellationToken)
            .ConfigureAwait(true);

        SessionFlowTrace.Log("Interactive OAuth: session established.");
        return new SessionLoginAttemptResult
        {
            Succeeded = true,
            Token = interactive.Token,
            ResolvedLoginOptions = interactive.ResolvedLoginOptions
                                     ?? SessionLoginReconciler.Reconcile(loginOptions, interactive.Token)
                                         .ResolvedLoginOptions,
            ApiVersionResolution = apiResolution,
            Diagnostics = diagnostics,
        };
    }
}
