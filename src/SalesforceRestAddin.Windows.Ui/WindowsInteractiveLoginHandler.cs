using System;
using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.Net;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class WindowsInteractiveLoginHandler : IInteractiveLoginHandler
{
    private readonly SessionAuthenticator _authenticator;
    private readonly SalesforceOAuthOptions _oauth;

    public WindowsInteractiveLoginHandler(
        SessionAuthenticator authenticator,
        SalesforceOAuthOptions? oauth = null)
    {
        _authenticator = authenticator ?? throw new ArgumentNullException(nameof(authenticator));
        _oauth = oauth ?? SalesforceOAuthOptions.Default;
    }

    public Task<InteractiveLoginResult> TrySignInAsync(
        SalesforceLoginOptions loginOptions,
        string? preferredApiVersion,
        CancellationToken cancellationToken = default) =>
        WpfUiThread.RunAsync(() => RunInteractive(loginOptions, cancellationToken));

    private InteractiveLoginResult RunInteractive(
        SalesforceLoginOptions loginOptions,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        loginOptions = SalesforceLoginOptionsNormalizer.Normalize(loginOptions);
        var pkce = PkcePair.Create();
        var authorizeUrl = SalesforceLoginUrlBuilder.BuildAuthorizeUrl(loginOptions, pkce, _oauth);

        SessionFlowTrace.Log("Interactive OAuth: opening sign-in window.");
        var oauthWindow = new OAuthWebViewWindow(authorizeUrl, _oauth);
        oauthWindow.ShowDialog();
        SessionFlowTrace.Log("Interactive OAuth: sign-in window closed.");

        var errorMessage = oauthWindow.ErrorMessage;
        if (errorMessage is not null && !string.IsNullOrWhiteSpace(errorMessage))
        {
            return InteractiveLoginResult.Failed(errorMessage);
        }

        var authorizationCode = oauthWindow.AuthorizationCode;
        if (authorizationCode is null || string.IsNullOrWhiteSpace(authorizationCode))
        {
            return InteractiveLoginResult.CancelledResult();
        }

        try
        {
            TlsProtocolBootstrap.EnsureEnabled();
            SessionFlowTrace.Log("Interactive OAuth: exchanging authorization code for tokens.");
            var loginResult = _authenticator.CompleteAuthorizationCodeLogin(
                loginOptions,
                authorizationCode,
                pkce.Verifier);
            SessionFlowTrace.Log("Interactive OAuth: token exchange succeeded.");
            return InteractiveLoginResult.Success(
                loginResult.Token,
                loginResult.ResolvedLoginOptions);
        }
        catch (Exception ex)
        {
            SessionFlowTrace.LogException("Interactive OAuth: token exchange failed", ex);
            return InteractiveLoginResult.Failed(ExceptionChainText.Format(ex));
        }
    }
}
