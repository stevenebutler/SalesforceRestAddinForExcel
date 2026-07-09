using System.Net;
using SalesforceRestAddin.Core.OAuth;
using TUnit.Mocks;

namespace SalesforceRestAddin.Tests;

public sealed class SalesforceOAuthClientTests
{
    [Test]
    public async Task ExchangeAuthorizationCodeAsync_PostsPkceFormAndParsesToken()
    {
        using var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        client.Handler.OnPost("/services/oauth2/token")
            .RespondWithJson(
                """
                {
                  "access_token": "00Daccess",
                  "refresh_token": "5Arefresh",
                  "instance_url": "https://kjr.my.salesforce.com",
                  "id": "https://login.salesforce.com/id/00Dxx/005yy",
                  "token_type": "Bearer",
                  "issued_at": "1710000000000"
                }
                """);

        var login = new SalesforceLoginOptions { Tenant = "kjr" };
        var oauthClient = new SalesforceOAuthClient(client);
        var token = await oauthClient.ExchangeAuthorizationCodeAsync(
            login,
            authorizationCode: "auth-code",
            codeVerifier: "pkce-verifier");

        await Assert.That(token.AccessToken).IsEqualTo("00Daccess");
        await Assert.That(token.RefreshToken).IsEqualTo("5Arefresh");
        await Assert.That(token.InstanceUrl).IsEqualTo("https://kjr.my.salesforce.com");
        await Assert.That(client.Handler.Requests.Count()).IsEqualTo(1);
        await Assert.That(client.Handler.Requests[0].Method).IsEqualTo(HttpMethod.Post);
    }

    [Test]
    public async Task RefreshAccessTokenAsync_PostsRefreshGrant()
    {
        using var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        client.Handler.OnPost("/services/oauth2/token")
            .RespondWithJson(
                """
                {
                  "access_token": "refreshed",
                  "instance_url": "https://kjr.my.salesforce.com",
                  "id": "https://login.salesforce.com/id/00Dxx/005yy",
                  "issued_at": "1710000000000"
                }
                """);

        var oauthClient = new SalesforceOAuthClient(client);
        var token = await oauthClient.RefreshAccessTokenAsync(
            new SalesforceLoginOptions { Tenant = "kjr" },
            refreshToken: "saved-refresh");

        await Assert.That(token.AccessToken).IsEqualTo("refreshed");
    }

    [Test]
    public async Task ExchangeAuthorizationCodeAsync_ThrowsOnErrorResponse()
    {
        using var client = Mock.HttpClient("https://login.salesforce.com");
        client.Handler.OnPost("/services/oauth2/token")
            .RespondWithJson(
                """{"error":"invalid_grant","error_description":"bad code"}""",
                HttpStatusCode.BadRequest);

        var oauthClient = new SalesforceOAuthClient(client);

        await Assert.That(async () => await oauthClient.ExchangeAuthorizationCodeAsync(
                new SalesforceLoginOptions(),
                authorizationCode: "bad",
                codeVerifier: "verifier"))
            .Throws<InvalidOperationException>()
            .WithMessageContaining("invalid_grant");
    }
}
