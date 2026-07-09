using System.Net;
using System.Net.Http;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests.DataPlane;

internal static class SalesforceTestClients
{
    public static (SalesforceDataClient client, SequentialMockHttpHandler handler) Create(
        Action<SequentialMockHttpHandler>? configure = null)
    {
        var handler = new SequentialMockHttpHandler();
        configure?.Invoke(handler);
        var http = new HttpClient(handler) { BaseAddress = new Uri("https://example.my.salesforce.com") };
        var session = new SessionContext
        {
            AccessToken = "token",
            InstanceUrl = "https://example.my.salesforce.com",
            Id = "https://login.salesforce.com/id/00D/005",
            IssuedAtUnixMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
        };
        session.SetApiVersion("66.0");
        var creds = new InMemorySessionCredentialStore();
        var prefs = new InMemoryUserLoginPreferencesStore();
        var authenticator = new SessionAuthenticator(creds, new SalesforceOAuthClient(http), SalesforceOAuthOptions.Default);
        var orchestrator = new SessionLoginOrchestrator(session, authenticator, new FakeInteractiveLoginHandler());
        var authClient = new SalesforceAuthenticatedClient(
            http,
            session,
            prefs,
            new SessionAccessTokenRefresher(authenticator),
            orchestrator);
        return (new SalesforceDataClient(authClient, session), handler);
    }
}
