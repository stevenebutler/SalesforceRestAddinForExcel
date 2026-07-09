using System;
using System.Collections.Generic;
using System.Net.Http;

namespace SalesforceRestAddin.Core.OAuth;

/// <summary>
/// Exchanges Salesforce OAuth authorization codes for access tokens (PKCE public client).
/// </summary>
public sealed class SalesforceOAuthClient
{
    private readonly HttpClient _http;

    public SalesforceOAuthClient(HttpClient http)
    {
        _http = http ?? throw new ArgumentNullException(nameof(http));
    }

    internal HttpClient HttpClient => _http;

    public async Task<OAuthTokenResponse> ExchangeAuthorizationCodeAsync(
        SalesforceLoginOptions login,
        string authorizationCode,
        string codeVerifier,
        SalesforceOAuthOptions? oauth = null,
        CancellationToken cancellationToken = default)
    {
        if (login is null)
        {
            throw new ArgumentNullException(nameof(login));
        }

        if (string.IsNullOrWhiteSpace(authorizationCode))
        {
            throw new ArgumentException("Authorization code is required.", nameof(authorizationCode));
        }

        if (string.IsNullOrWhiteSpace(codeVerifier))
        {
            throw new ArgumentException("PKCE code verifier is required.", nameof(codeVerifier));
        }

        oauth ??= SalesforceOAuthOptions.Default;
        login = SalesforceLoginOptionsNormalizer.Normalize(login);
        var tokenUrl = SalesforceLoginUrlBuilder.BuildTokenUrl(login);

        using var content = new FormUrlEncodedContent(BuildTokenRequestBody(
            authorizationCode,
            codeVerifier,
            oauth));

        using var response = await _http.PostAsync(tokenUrl, content, cancellationToken)
            .ConfigureAwait(false);
        var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"Salesforce token exchange failed ({(int)response.StatusCode}): {json}");
        }

        return OAuthTokenResponse.ParseJson(json);
    }

    public async Task<OAuthTokenResponse> RefreshAccessTokenAsync(
        SalesforceLoginOptions login,
        string refreshToken,
        SalesforceOAuthOptions? oauth = null,
        CancellationToken cancellationToken = default)
    {
        login = SalesforceLoginOptionsNormalizer.Normalize(login);
        var tokenHost = new Uri(SalesforceLoginUrlBuilder.BuildLoginBaseUrl(login)).Host;
        return await RefreshAccessTokenAsync(tokenHost, refreshToken, oauth, cancellationToken)
            .ConfigureAwait(false);
    }

    public async Task<OAuthTokenResponse> RefreshAccessTokenAsync(
        string oauthTokenHost,
        string refreshToken,
        SalesforceOAuthOptions? oauth = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(oauthTokenHost))
        {
            throw new ArgumentException("OAuth token host is required.", nameof(oauthTokenHost));
        }

        if (string.IsNullOrWhiteSpace(refreshToken))
        {
            throw new ArgumentException("Refresh token is required.", nameof(refreshToken));
        }

        oauth ??= SalesforceOAuthOptions.Default;
        var tokenUrl = $"https://{oauthTokenHost.Trim().ToLowerInvariant()}/services/oauth2/token";

        using var content = new FormUrlEncodedContent(BuildRefreshRequestBody(refreshToken, oauth));
        using var response = await _http.PostAsync(tokenUrl, content, cancellationToken)
            .ConfigureAwait(false);
        var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"Salesforce token refresh failed ({(int)response.StatusCode}): {json}");
        }

        return OAuthTokenResponse.ParseJson(json);
    }

    private static IEnumerable<KeyValuePair<string, string>> BuildRefreshRequestBody(
        string refreshToken,
        SalesforceOAuthOptions oauth)
    {
        yield return new KeyValuePair<string, string>("grant_type", "refresh_token");
        yield return new KeyValuePair<string, string>("refresh_token", refreshToken);
        yield return new KeyValuePair<string, string>("client_id", oauth.ClientId);

        var refreshSecret = oauth.ClientSecret;
        if (refreshSecret is not null && !string.IsNullOrWhiteSpace(refreshSecret))
        {
            yield return new KeyValuePair<string, string>("client_secret", refreshSecret);
        }
    }

    private static IEnumerable<KeyValuePair<string, string>> BuildTokenRequestBody(
        string authorizationCode,
        string codeVerifier,
        SalesforceOAuthOptions oauth)
    {
        yield return new KeyValuePair<string, string>("grant_type", "authorization_code");
        yield return new KeyValuePair<string, string>("code", authorizationCode);
        yield return new KeyValuePair<string, string>("client_id", oauth.ClientId);
        yield return new KeyValuePair<string, string>("redirect_uri", oauth.RedirectUri);
        yield return new KeyValuePair<string, string>("code_verifier", codeVerifier);

        var tokenSecret = oauth.ClientSecret;
        if (tokenSecret is not null && !string.IsNullOrWhiteSpace(tokenSecret))
        {
            yield return new KeyValuePair<string, string>("client_secret", tokenSecret);
        }
    }
}
