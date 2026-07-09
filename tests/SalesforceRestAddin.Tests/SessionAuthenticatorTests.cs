using System.Net;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;
using TUnit.Mocks;

namespace SalesforceRestAddin.Tests;

public sealed class SessionAuthenticatorTests
{
    [Test]
    public async Task TryRestoreSessionAsync_UsesStoredOAuthTokenHost()
    {
        using var client = Mock.HttpClient("https://login.salesforce.com");
        client.Handler.OnPost("/services/oauth2/token")
            .RespondWithJson(
                """
                {
                  "access_token": "refreshed",
                  "instance_url": "https://kjr.my.salesforce.com",
                  "id": "https://login.salesforce.com/id/00Dxx/005yy"
                }
                """);
        client.Handler.OnGet("/services/data/")
            .RespondWithJson("""
                [
                  {"label":"Winter '26","url":"/services/data/v66.0/","version":"66.0"}
                ]
                """);

        var store = new InMemorySessionCredentialStore();
        store.Save(new StoredSessionCredentials
        {
            HostKey = "kjr.my.salesforce.com",
            OAuthTokenHost = "login.salesforce.com",
            RefreshToken = "saved-refresh",
        });

        var authenticator = new SessionAuthenticator(store, new SalesforceOAuthClient(client));
        var login = new SalesforceLoginOptions { Tenant = "kjr" };
        var (token, failureReason) = await authenticator.TryRestoreSessionAsync(login);

        await Assert.That(token).IsNotNull();
        await Assert.That(failureReason).IsNull();
        await Assert.That(client.Handler.Requests[0].RequestUri!.Host).IsEqualTo("login.salesforce.com");
        await Assert.That(store.Load("kjr.my.salesforce.com")).IsNotNull();
    }

    [Test]
    public async Task CompleteAuthorizationCodeLogin_RekeysGenericCredentialToMyDomain()
    {
        using var client = Mock.HttpClient("https://login.salesforce.com");
        client.Handler.OnPost("/services/oauth2/token")
            .RespondWithJson(
                """
                {
                  "access_token": "access",
                  "refresh_token": "refresh",
                  "instance_url": "https://kjr.my.salesforce.com",
                  "id": "https://login.salesforce.com/id/00Dxx/005yy"
                }
                """);

        var store = new InMemorySessionCredentialStore();
        var authenticator = new SessionAuthenticator(store, new SalesforceOAuthClient(client));
        var result = authenticator.CompleteAuthorizationCodeLogin(
            new SalesforceLoginOptions { Environment = SalesforceEnvironment.Production },
            authorizationCode: "code",
            codeVerifier: "verifier");

        await Assert.That(result.ResolvedLoginOptions.Tenant).IsEqualTo("kjr");
        await Assert.That(store.Load("kjr.my.salesforce.com")!.OAuthTokenHost).IsEqualTo("login.salesforce.com");
        await Assert.That(store.Load("login.salesforce.com")).IsNull();
        await Assert.That(result.PersistResult.Saved).IsTrue();
    }

    [Test]
    public async Task TryRestoreSessionAsync_UsesStoredRefreshTokenPerHost()
    {
        using var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        client.Handler.OnPost("/services/oauth2/token")
            .RespondWithJson(
                """
                {
                  "access_token": "refreshed",
                  "instance_url": "https://kjr.my.salesforce.com",
                  "id": "https://login.salesforce.com/id/00Dxx/005yy"
                }
                """);
        client.Handler.OnGet("/services/data/")
            .RespondWithJson("""
                [
                  {"label":"Winter '26","url":"/services/data/v66.0/","version":"66.0"},
                  {"label":"Spring '26","url":"/services/data/v67.0/","version":"67.0"}
                ]
                """);

        var store = new InMemorySessionCredentialStore();
        store.Save(new StoredSessionCredentials
        {
            HostKey = "kjr.my.salesforce.com",
            RefreshToken = "saved-refresh",
        });

        var authenticator = new SessionAuthenticator(store, new SalesforceOAuthClient(client));
        var login = new SalesforceLoginOptions { Tenant = "kjr" };
        var (token, failureReason) = await authenticator.TryRestoreSessionAsync(login);

        await Assert.That(token).IsNotNull();
        await Assert.That(failureReason).IsNull();
        await Assert.That(token!.AccessToken).IsEqualTo("refreshed");
        await authenticator.ApplyToSessionAsync(token, SessionContext.Current);
        await Assert.That(SessionContext.Current.IsLoggedIn).IsTrue();
        await Assert.That(SessionContext.Current.ApiVersion).IsEqualTo("67.0");
    }

    [Test]
    public async Task TryRestoreSessionAsync_DeletesStoredCredentialWhenRefreshFails()
    {
        using var client = Mock.HttpClient("https://login.salesforce.com");
        client.Handler.OnPost("/services/oauth2/token")
            .RespondWithJson("""{"error":"invalid_grant"}""", HttpStatusCode.BadRequest);

        var store = new InMemorySessionCredentialStore();
        store.Save(new StoredSessionCredentials
        {
            HostKey = "login.salesforce.com",
            RefreshToken = "revoked",
        });

        var authenticator = new SessionAuthenticator(store, new SalesforceOAuthClient(client));
        var (token, _) = await authenticator.TryRestoreSessionAsync(new SalesforceLoginOptions());

        await Assert.That(token).IsNull();
        await Assert.That(store.Load("login.salesforce.com")).IsNull();
    }
}
