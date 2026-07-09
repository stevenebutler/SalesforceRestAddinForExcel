using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.OAuth;

namespace SalesforceRestAddin.Core.Session;

public sealed class InteractiveLoginResult
{
    public bool Succeeded { get; init; }

    public bool Cancelled { get; init; }

    public string? ErrorMessage { get; init; }

    public OAuthTokenResponse? Token { get; init; }

    public SalesforceLoginOptions? ResolvedLoginOptions { get; init; }

    public static InteractiveLoginResult Success(OAuthTokenResponse token, SalesforceLoginOptions resolvedLoginOptions) =>
        new()
        {
            Succeeded = true,
            Token = token,
            ResolvedLoginOptions = resolvedLoginOptions,
        };

    public static InteractiveLoginResult CancelledResult() =>
        new() { Cancelled = true };

    public static InteractiveLoginResult Failed(string message) =>
        new() { ErrorMessage = message };
}

/// <summary>
/// Embedded WebView2 OAuth sign-in when silent refresh cannot restore a session.
/// </summary>
public interface IInteractiveLoginHandler
{
    Task<InteractiveLoginResult> TrySignInAsync(
        SalesforceLoginOptions loginOptions,
        string? preferredApiVersion,
        CancellationToken cancellationToken = default);
}
