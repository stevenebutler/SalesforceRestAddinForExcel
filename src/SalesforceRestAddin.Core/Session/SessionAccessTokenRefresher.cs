using System;
using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.OAuth;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Silent access-token refresh from stored credentials with single-flight coalescing.
/// </summary>
public sealed class SessionAccessTokenRefresher
{
    private readonly SessionAuthenticator _authenticator;
    private readonly IAuthRequestMutator _authMutator;
    private readonly ISessionRecoveryTracer _recoveryTracer;
    private readonly object _inflightLock = new();
    private Task<bool>? _inflightRefresh;

    public SessionAccessTokenRefresher(
        SessionAuthenticator authenticator,
        IAuthRequestMutator? authMutator = null,
        ISessionRecoveryTracer? recoveryTracer = null)
    {
        _authenticator = authenticator ?? throw new ArgumentNullException(nameof(authenticator));
        _authMutator = authMutator ?? NullAuthRequestMutator.Instance;
        _recoveryTracer = recoveryTracer ?? NullSessionRecoveryTracer.Instance;
    }

    public Task<bool> TryRefreshAsync(
        SalesforceLoginOptions loginOptions,
        SessionContext session,
        CancellationToken cancellationToken = default)
    {
        lock (_inflightLock)
        {
            // Only coalesce concurrent callers; never reuse a completed refresh from a prior REST call.
            if (_inflightRefresh is { IsCompleted: false } inFlight)
            {
                _recoveryTracer.AddStep("Joining in-flight silent refresh.");
                return inFlight;
            }

            _inflightRefresh = RefreshCoreAsync(loginOptions, session, cancellationToken);
            return _inflightRefresh;
        }
    }

    private async Task<bool> RefreshCoreAsync(
        SalesforceLoginOptions loginOptions,
        SessionContext session,
        CancellationToken cancellationToken)
    {
        try
        {
            var (token, failureReason) = await _authenticator
                .TryRestoreSessionAsync(loginOptions, _authMutator, _recoveryTracer, cancellationToken)
                .ConfigureAwait(false);

            if (token is null)
            {
                if (!string.IsNullOrWhiteSpace(failureReason))
                {
                    _recoveryTracer.AddStep($"Silent refresh failed: {failureReason}");
                }

                return false;
            }

            token.ApplyTo(session);
            return true;
        }
        finally
        {
            lock (_inflightLock)
            {
                _inflightRefresh = null;
            }
        }
    }
}
