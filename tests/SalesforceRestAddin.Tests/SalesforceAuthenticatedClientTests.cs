using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class SalesforceAuthenticatedClientTests
{
    [Test]
    public async Task ForbiddenThenRefreshThenSuccess()
    {
        var handler = new SequentialMockHttpHandler();
        handler.EnqueueForIdentity(HttpStatusCode.Forbidden, SalesforceMockResponses.MissingOAuthToken403);
        handler.EnqueueForTokenEndpoint(HttpStatusCode.OK, SalesforceMockResponses.RefreshSuccess);
        handler.EnqueueForIdentity(HttpStatusCode.OK, SalesforceMockResponses.IdentitySuccess);

        using var http = new HttpClient(handler) { BaseAddress = new Uri("https://kjr.my.salesforce.com") };
        var mutator = new TestAuthRequestMutator { CorruptBearer = true };
        var stack = CreateAuthenticatedStack(http, mutator);
        SeedLoggedInSession(stack.session);
        stack.creds.Save(StoredCredentials());

        using var request = new HttpRequestMessage(
            HttpMethod.Get,
            "https://login.salesforce.com/id/00Dxx/005yy?version=latest");
        using var response = await stack.client.SendAsync(request);

        await Assert.That(response.IsSuccessStatusCode).IsTrue();
        await Assert.That(handler.Requests.Count).IsEqualTo(3);
    }

    [Test]
    public async Task UnauthorizedThenRefreshThenSuccess()
    {
        var handler = CreateRecoveryHandler();
        using var http = new HttpClient(handler) { BaseAddress = new Uri("https://kjr.my.salesforce.com") };
        var stack = CreateAuthenticatedStack(http);
        SeedLoggedInSession(stack.session);
        stack.creds.Save(StoredCredentials());

        using var request = new HttpRequestMessage(
            HttpMethod.Get,
            "https://login.salesforce.com/id/00Dxx/005yy?version=latest");
        using var response = await stack.client.SendAsync(request);

        await Assert.That(response.IsSuccessStatusCode).IsTrue();
        await Assert.That(handler.Requests.Count).IsEqualTo(3);
        await Assert.That(handler.Requests[1].RequestUri!.AbsolutePath).Contains("oauth2/token");
    }

    [Test]
    public async Task CorruptBearer_ClearedAfterRefresh()
    {
        var handler = CreateRecoveryHandler();
        using var http = new HttpClient(handler) { BaseAddress = new Uri("https://kjr.my.salesforce.com") };
        var mutator = new TestAuthRequestMutator { CorruptBearer = true };
        var stack = CreateAuthenticatedStack(http, mutator);
        SeedLoggedInSession(stack.session);
        stack.creds.Save(StoredCredentials());

        using var request = new HttpRequestMessage(
            HttpMethod.Get,
            "https://login.salesforce.com/id/00Dxx/005yy?version=latest");
        using var response = await stack.client.SendAsync(request);

        await Assert.That(response.IsSuccessStatusCode).IsTrue();
        await Assert.That(handler.Requests[2].Headers.Authorization!.Parameter).IsEqualTo("refreshed-access");
    }

    [Test]
    public async Task RefreshFails_EscalatesInteractive()
    {
        var handler = new SequentialMockHttpHandler();
        handler.EnqueueForIdentity(HttpStatusCode.Unauthorized, SalesforceMockResponses.InvalidSessionId401);
        handler.EnqueueForTokenEndpoint(HttpStatusCode.BadRequest, SalesforceMockResponses.InvalidGrant400);
        handler.EnqueueForIdentity(HttpStatusCode.OK, SalesforceMockResponses.IdentitySuccess);

        using var http = new HttpClient(handler) { BaseAddress = new Uri("https://kjr.my.salesforce.com") };
        var stack = CreateAuthenticatedStack(http);
        SeedLoggedInSession(stack.session);
        stack.prefs.Stored = ValidPrefs();
        stack.creds.Save(StoredCredentials());
        stack.interactive.NextResult = InteractiveLoginResult.CancelledResult();

        using var request = new HttpRequestMessage(
            HttpMethod.Get,
            "https://login.salesforce.com/id/00Dxx/005yy?version=latest");

        await Assert.That(async () => await stack.client.SendAsync(request))
            .Throws<SalesforceLoginCancelledException>();
        await Assert.That(stack.interactive.SignInCount).IsEqualTo(1);
    }

    [Test]
    public async Task ForceFailureAfterRecovery()
    {
        var handler = new SequentialMockHttpHandler();
        handler.EnqueueForIdentity(HttpStatusCode.Unauthorized, SalesforceMockResponses.InvalidSessionId401);
        handler.EnqueueForTokenEndpoint(HttpStatusCode.OK, SalesforceMockResponses.RefreshSuccess);
        handler.EnqueueForIdentity(HttpStatusCode.Unauthorized, SalesforceMockResponses.InvalidSessionId401);

        using var http = new HttpClient(handler) { BaseAddress = new Uri("https://kjr.my.salesforce.com") };
        var mutator = new TestAuthRequestMutator { CorruptBearer = true, ForceFailureAfterRecovery = true };
        var stack = CreateAuthenticatedStack(http, mutator);
        SeedLoggedInSession(stack.session);
        stack.creds.Save(StoredCredentials());

        using var request = new HttpRequestMessage(
            HttpMethod.Get,
            "https://login.salesforce.com/id/00Dxx/005yy?version=latest");

        await Assert.That(async () => await stack.client.SendAsync(request))
            .Throws<SessionRecoveryException>()
            .WithMessageContaining("auth error persisted");
    }

    [Test]
    public async Task JwtMode_Unauthorized_DoesNotRefresh()
    {
        var handler = new SequentialMockHttpHandler();
        handler.EnqueueForIdentity(HttpStatusCode.Unauthorized, SalesforceMockResponses.InvalidSessionId401);

        using var http = new HttpClient(handler) { BaseAddress = new Uri("https://kjr.my.salesforce.com") };
        var stack = CreateAuthenticatedStack(http);
        stack.session.AccessToken = "jwt-debug-token";
        stack.session.InstanceUrl = "https://kjr.my.salesforce.com";
        stack.session.IsJwtAccessTokenMode = true;
        stack.session.SetApiVersion("66.0");
        stack.creds.Save(StoredCredentials());

        using var request = new HttpRequestMessage(
            HttpMethod.Get,
            "https://login.salesforce.com/id/00Dxx/005yy?version=latest");

        var ex = await Assert.That(async () => await stack.client.SendAsync(request))
            .Throws<SessionRecoveryException>();
        await Assert.That(ex!.Message).Contains("jwtAccessToken");
        await Assert.That(handler.Requests.Count).IsEqualTo(1);
        await Assert.That(handler.Requests.Any(r => r.RequestUri!.AbsolutePath.Contains("oauth2/token"))).IsFalse();
        await Assert.That(stack.interactive.SignInCount).IsEqualTo(0);
        await Assert.That(stack.creds.Load("kjr.my.salesforce.com")!.RefreshToken).IsEqualTo("saved-refresh");
    }

    [Test]
    public async Task SequentialRecovery_TwoCalls_TwoRefreshPosts()
    {
        var handler = new SequentialMockHttpHandler();
        handler.EnqueueForIdentity(HttpStatusCode.Unauthorized, SalesforceMockResponses.InvalidSessionId401);
        handler.EnqueueForTokenEndpoint(HttpStatusCode.OK, SalesforceMockResponses.RefreshSuccess);
        handler.EnqueueForIdentity(HttpStatusCode.OK, SalesforceMockResponses.IdentitySuccess);
        handler.EnqueueForIdentity(HttpStatusCode.Unauthorized, SalesforceMockResponses.InvalidSessionId401);
        handler.EnqueueForTokenEndpoint(HttpStatusCode.OK, SalesforceMockResponses.RefreshSuccess);
        handler.EnqueueForIdentity(HttpStatusCode.OK, SalesforceMockResponses.IdentitySuccess);

        using var http = new HttpClient(handler) { BaseAddress = new Uri("https://kjr.my.salesforce.com") };
        var mutator = new TestAuthRequestMutator { CorruptBearer = true };
        var stack = CreateAuthenticatedStack(http, mutator);
        SeedLoggedInSession(stack.session);
        stack.creds.Save(StoredCredentials());

        using var request = new HttpRequestMessage(
            HttpMethod.Get,
            "https://login.salesforce.com/id/00Dxx/005yy?version=latest");

        using var response1 = await stack.client.SendAsync(Clone(request));
        await Assert.That(response1.IsSuccessStatusCode).IsTrue();

        stack.session.AccessToken = "stale-access";

        using var response2 = await stack.client.SendAsync(Clone(request));
        await Assert.That(response2.IsSuccessStatusCode).IsTrue();

        var refreshPosts = handler.Requests.Count(r =>
            r.Method == HttpMethod.Post && r.RequestUri!.AbsolutePath.Contains("oauth2/token"));
        await Assert.That(refreshPosts).IsEqualTo(2);
    }

    [Test]
    public async Task CorruptRefresh_EscalatesInteractive()
    {
        string? refreshRequestBody = null;
        var handler = new SequentialMockHttpHandler();
        handler.EnqueueForIdentity(HttpStatusCode.Unauthorized, SalesforceMockResponses.InvalidSessionId401);
        handler.EnqueueForRoute("oauth2/token", req =>
        {
            refreshRequestBody = req.Content!.ReadAsStringAsync().GetAwaiter().GetResult();
            return new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = new StringContent(
                    SalesforceMockResponses.InvalidGrant400,
                    System.Text.Encoding.UTF8,
                    "application/json"),
            };
        });
        handler.EnqueueForIdentity(HttpStatusCode.OK, SalesforceMockResponses.IdentitySuccess);

        using var http = new HttpClient(handler) { BaseAddress = new Uri("https://kjr.my.salesforce.com") };
        var mutator = new TestAuthRequestMutator { CorruptBearer = true, CorruptRefresh = true };
        var stack = CreateAuthenticatedStack(http, mutator);
        SeedLoggedInSession(stack.session);
        stack.prefs.Stored = ValidPrefs();
        stack.creds.Save(StoredCredentials());
        stack.interactive.NextResult = InteractiveLoginResult.CancelledResult();

        using var request = new HttpRequestMessage(
            HttpMethod.Get,
            "https://login.salesforce.com/id/00Dxx/005yy?version=latest");

        await Assert.That(async () => await stack.client.SendAsync(request))
            .Throws<SalesforceLoginCancelledException>();
        await Assert.That(stack.interactive.SignInCount).IsEqualTo(1);
        await Assert.That(refreshRequestBody).IsNotNull();
        await Assert.That(refreshRequestBody!).Contains("CORRUPTED%3A");
    }

    [Test]
    public async Task Concurrent401_SingleRefresh()
    {
        var handler = CreateRecoveryHandler(extraIdentityFailures: 1);
        using var http = new HttpClient(handler) { BaseAddress = new Uri("https://kjr.my.salesforce.com") };
        var stack = CreateAuthenticatedStack(http);
        SeedLoggedInSession(stack.session);
        stack.creds.Save(StoredCredentials());

        using var request1 = new HttpRequestMessage(
            HttpMethod.Get,
            "https://login.salesforce.com/id/00Dxx/005yy?version=latest");
        using var request2 = new HttpRequestMessage(
            HttpMethod.Get,
            "https://login.salesforce.com/id/00Dxx/005yy?version=latest");

        var task1 = stack.client.SendAsync(Clone(request1));
        var task2 = stack.client.SendAsync(Clone(request2));
        await Task.WhenAll(task1, task2);

        var refreshPosts = handler.Requests.Count(r =>
            r.Method == HttpMethod.Post && r.RequestUri!.AbsolutePath.Contains("oauth2/token"));
        await Assert.That(refreshPosts).IsLessThanOrEqualTo(2);
        await Assert.That(refreshPosts).IsGreaterThanOrEqualTo(1);
    }

    private static SequentialMockHttpHandler CreateRecoveryHandler(int extraIdentityFailures = 0)
    {
        var handler = new SequentialMockHttpHandler();
        handler.EnqueueForIdentity(HttpStatusCode.Unauthorized, SalesforceMockResponses.InvalidSessionId401);
        handler.EnqueueForTokenEndpoint(HttpStatusCode.OK, SalesforceMockResponses.RefreshSuccess);
        handler.EnqueueForTokenEndpoint(HttpStatusCode.OK, SalesforceMockResponses.RefreshSuccess);
        handler.EnqueueForIdentity(HttpStatusCode.OK, SalesforceMockResponses.IdentitySuccess);

        for (var i = 0; i < extraIdentityFailures; i++)
        {
            handler.EnqueueForIdentity(HttpStatusCode.Unauthorized, SalesforceMockResponses.InvalidSessionId401);
            handler.EnqueueForIdentity(HttpStatusCode.OK, SalesforceMockResponses.IdentitySuccess);
        }

        return handler;
    }

    private static HttpRequestMessage Clone(HttpRequestMessage request) =>
        new(request.Method, request.RequestUri);

    private static void SeedLoggedInSession(SessionContext session)
    {
        session.AccessToken = "stale-access";
        session.RefreshToken = "saved-refresh";
        session.InstanceUrl = "https://kjr.my.salesforce.com";
        session.Id = "https://login.salesforce.com/id/00Dxx/005yy";
        session.IssuedAtUnixMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        session.SetApiVersion("66.0");
    }

    private static UserLoginPreferences ValidPrefs() =>
        new()
        {
            Tenant = "kjr",
            Environment = SalesforceEnvironment.Production,
        };

    private static StoredSessionCredentials StoredCredentials() =>
        new()
        {
            HostKey = "kjr.my.salesforce.com",
            RefreshToken = "saved-refresh",
        };

    private static AuthenticatedClientTestStack CreateAuthenticatedStack(
        HttpClient http,
        TestAuthRequestMutator? mutator = null)
    {
        var session = new SessionContext();
        var creds = new InMemorySessionCredentialStore();
        var prefs = new InMemoryUserLoginPreferencesStore { Stored = ValidPrefs() };
        var interactive = new FakeInteractiveLoginHandler();
        var authenticator = new SessionAuthenticator(creds, new SalesforceOAuthClient(http));
        var refresher = new SessionAccessTokenRefresher(authenticator, mutator);
        var orchestrator = new SessionLoginOrchestrator(session, authenticator, interactive);
        var client = new SalesforceAuthenticatedClient(
            http,
            session,
            prefs,
            refresher,
            orchestrator,
            mutator);
        return new AuthenticatedClientTestStack(client, creds, prefs, interactive, session);
    }

    private sealed record AuthenticatedClientTestStack(
        SalesforceAuthenticatedClient client,
        InMemorySessionCredentialStore creds,
        InMemoryUserLoginPreferencesStore prefs,
        FakeInteractiveLoginHandler interactive,
        SessionContext session);
}
