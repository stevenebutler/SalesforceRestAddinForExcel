using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;
using TUnit.Mocks;

namespace SalesforceRestAddin.Tests;

public sealed class SessionGateTests
{
    [Test]
    public async Task AlreadyLoggedIn_NoUi()
    {
        var stack = CreateGate();
        stack.session.AccessToken = "token";
        stack.session.InstanceUrl = "https://kjr.my.salesforce.com";
        stack.session.Id = "https://login.salesforce.com/id/00D/005";
        stack.session.IssuedAtUnixMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        stack.session.SetApiVersion("66.0");

        await stack.gate.EnsureLoggedInAsync();

        await Assert.That(stack.presenter.ShowCount).IsEqualTo(0);
        await Assert.That(stack.interactive.SignInCount).IsEqualTo(0);
        await Assert.That(stack.session.IsLoggedIn).IsTrue();
    }

    [Test]
    public async Task ReturningUser_SilentRefresh_NoOptionsDialog()
    {
        using var client = CreateSalesforceMockClient();
        var stack = CreateGate(client);
        stack.prefs.Stored = ValidPrefs();
        stack.creds.Save(StoredCredentials());

        await stack.gate.EnsureLoggedInAsync();

        await Assert.That(stack.presenter.ShowCount).IsEqualTo(0);
        await Assert.That(stack.interactive.SignInCount).IsEqualTo(0);
        await Assert.That(stack.session.IsLoggedIn).IsTrue();
        await Assert.That(stack.prefs.Stored!.ShowLoginOptionsOnNextUse).IsFalse();
    }

    [Test]
    public async Task SilentRefresh_ReportsStatusPhases()
    {
        using var client = CreateSalesforceMockClient();
        var stack = CreateGate(client);
        stack.prefs.Stored = ValidPrefs();
        stack.creds.Save(StoredCredentials());
        var statuses = new List<string>();

        await stack.gate.EnsureLoggedInAsync(reportStatus: statuses.Add);

        await Assert.That(stack.session.IsLoggedIn).IsTrue();
        await Assert.That(statuses).IsEquivalentTo(new[]
        {
            "Refreshing session...",
            "Loading metadata...",
        });
    }

    [Test]
    public async Task InteractiveLogin_ReportsSigningInThenMetadata()
    {
        using var client = CreateSalesforceMockClient(includeTokenRefresh: false);
        var stack = CreateGate(client);
        stack.prefs.Stored = null;
        stack.presenter.NextResult = PresenterResult();
        stack.interactive.NextResult = InteractiveSuccess();
        var statuses = new List<string>();

        await stack.gate.EnsureLoggedInAsync(reportStatus: statuses.Add);

        await Assert.That(stack.session.IsLoggedIn).IsTrue();
        await Assert.That(statuses).IsEquivalentTo(new[]
        {
            "Signing in...",
            "Loading metadata...",
        });
    }

    [Test]
    public async Task ExcelRestart_Simulated_NoOptionsIfNotLoggedOut()
    {
        using var client = CreateSalesforceMockClient();
        var stack = CreateGate(client);
        stack.prefs.Stored = ValidPrefs(showLoginOptionsOnNextUse: false);
        stack.creds.Save(StoredCredentials());

        await stack.gate.EnsureLoggedInAsync();

        await Assert.That(stack.presenter.ShowCount).IsEqualTo(0);
        await Assert.That(stack.session.IsLoggedIn).IsTrue();
    }

    [Test]
    public async Task FirstUse_NoValidPrefs_ShowsOptions()
    {
        using var client = CreateSalesforceMockClient(includeTokenRefresh: false);
        var stack = CreateGate(client);
        stack.prefs.Stored = null;
        stack.presenter.NextResult = PresenterResult();
        stack.interactive.NextResult = InteractiveSuccess();

        await stack.gate.EnsureLoggedInAsync();

        await Assert.That(stack.presenter.ShowCount).IsEqualTo(1);
        await Assert.That(stack.session.IsLoggedIn).IsTrue();
    }

    [Test]
    public async Task InvalidPrefsFile_TreatedAsMissing()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, """{"tenant":"kjr"}""");
        var store = new JsonUserLoginPreferencesStore(path);

        await Assert.That(store.TryLoadValid()).IsNull();

        File.Delete(path);
    }

    [Test]
    public async Task ValidPrefs_RequiresEnvironment()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, """{"tenant":"kjr","environment":"Production"}""");

        var loaded = new JsonUserLoginPreferencesStore(path).TryLoadValid();

        await Assert.That(loaded).IsNotNull();
        await Assert.That(loaded!.Tenant).IsEqualTo("kjr");

        File.Delete(path);
    }

    [Test]
    public async Task AfterLogout_ShowsOptionsThenSilent()
    {
        using var client = CreateSalesforceMockClient();
        var stack = CreateGate(client);
        stack.prefs.Stored = ValidPrefs();
        stack.creds.Save(StoredCredentials());
        stack.presenter.NextResult = PresenterResult();

        stack.gate.Logout();

        await Assert.That(stack.prefs.Stored!.ShowLoginOptionsOnNextUse).IsTrue();
        await Assert.That(stack.creds.Load("kjr.my.salesforce.com")).IsNotNull();

        await stack.gate.EnsureLoggedInAsync();

        await Assert.That(stack.presenter.ShowCount).IsEqualTo(1);
        await Assert.That(stack.interactive.SignInCount).IsEqualTo(0);
        await Assert.That(stack.session.IsLoggedIn).IsTrue();
    }

    [Test]
    public async Task T_GATE_02_Logout_Keeps_Credentials_Forces_Options()
    {
        using var client = CreateSalesforceMockClient();
        var stack = CreateGate(client);
        stack.prefs.Stored = ValidPrefs();
        stack.creds.Save(StoredCredentials());
        stack.session.AccessToken = "token";
        stack.session.Id = "https://login.salesforce.com/id/00D/005";
        stack.session.InstanceUrl = "https://kjr.my.salesforce.com";
        stack.session.DisplayName = "Test User";
        stack.session.IssuedAtUnixMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        stack.session.SetApiVersion("66.0");

        var cleared = stack.gate.Logout();

        await Assert.That(cleared).IsTrue();
        await Assert.That(stack.session.IsLoggedIn).IsFalse();
        await Assert.That(stack.session.DisplayName).IsNull();
        await Assert.That(stack.session.InstanceUrl).IsNull();
        await Assert.That(stack.prefs.Stored!.ShowLoginOptionsOnNextUse).IsTrue();
        await Assert.That(stack.creds.Load("kjr.my.salesforce.com")!.RefreshToken).IsEqualTo("saved-refresh");
    }

    [Test]
    public async Task AfterLogout_PersistsFlag_ReloadPrefs()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        var store = new JsonUserLoginPreferencesStore(path);
        store.Save(ValidPrefs());

        using var client = CreateSalesforceMockClient();
        var stack = CreateGate(client, store);

        stack.gate.Logout();

        var reloaded = new JsonUserLoginPreferencesStore(path).TryLoadValid();
        await Assert.That(reloaded!.ShowLoginOptionsOnNextUse).IsTrue();

        File.Delete(path);
    }

    [Test]
    public async Task OptionsCancel_Throws()
    {
        var stack = CreateGate();
        stack.prefs.Stored = null;
        stack.presenter.NextResult = null;

        await Assert.That(() => stack.gate.EnsureLoggedInAsync())
            .Throws<SalesforceLoginCancelledException>();
    }

    [Test]
    public async Task InteractiveCancel_Throws()
    {
        var stack = CreateGate();
        stack.prefs.Stored = null;
        stack.presenter.NextResult = PresenterResult();
        stack.interactive.NextResult = InteractiveLoginResult.CancelledResult();

        await Assert.That(() => stack.gate.EnsureLoggedInAsync())
            .Throws<SalesforceLoginCancelledException>();
    }

    [Test]
    public async Task AfterLogout_ForceSignIn_SkipsSilent()
    {
        using var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        client.Handler.OnPost("/services/oauth2/token")
            .RespondWithJson(SalesforceMockResponses.RefreshSuccess);
        client.Handler.OnGet("/services/data/")
            .RespondWithJson("""[{"version":"66.0","url":"/services/data/v66.0/"}]""");

        var stack = CreateGate(client);
        stack.prefs.Stored = ValidPrefs(showLoginOptionsOnNextUse: true);
        stack.creds.Save(StoredCredentials());
        stack.presenter.NextResult = PresenterResult(forceSignInAgain: true);
        stack.interactive.NextResult = InteractiveSuccess();

        await stack.gate.EnsureLoggedInAsync();

        await Assert.That(stack.interactive.SignInCount).IsEqualTo(1);
        await Assert.That(client.Handler.Requests.Count(r => r.Method == HttpMethod.Post)).IsEqualTo(0);
    }

    [Test]
    public async Task SilentFails_EscalatesInteractive()
    {
        using var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        client.Handler.OnPost("/services/oauth2/token")
            .RespondWithJson(SalesforceMockResponses.InvalidGrant400, HttpStatusCode.BadRequest);
        client.Handler.OnGet("/services/data/")
            .RespondWithJson("""[{"version":"66.0","url":"/services/data/v66.0/"}]""");

        var stack = CreateGate(client);
        stack.prefs.Stored = ValidPrefs();
        stack.creds.Save(StoredCredentials());
        stack.interactive.NextResult = InteractiveSuccess();

        await stack.gate.EnsureLoggedInAsync();

        await Assert.That(stack.interactive.SignInCount).IsEqualTo(1);
        await Assert.That(stack.session.IsLoggedIn).IsTrue();
    }

    [Test]
    public async Task SilentRefresh_NetworkError_EscalatesInteractive()
    {
        var handler = new SequentialMockHttpHandler();
        handler.EnqueueForRoute(
            "oauth2/token",
            _ => throw new HttpRequestException(
                "An error occurred while sending the request.",
                inner: new System.Net.Sockets.SocketException(10061, "Connection refused")));
        handler.EnqueueForRoute(
            "default",
            _ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""[{"version":"66.0","url":"/services/data/v66.0/"}]"""),
            });

        using var client = new HttpClient(handler) { BaseAddress = new Uri("https://kjr.my.salesforce.com") };
        var stack = CreateGate(client);
        stack.prefs.Stored = ValidPrefs();
        stack.creds.Save(StoredCredentials());
        stack.interactive.NextResult = InteractiveSuccess();

        await stack.gate.EnsureLoggedInAsync();

        await Assert.That(stack.interactive.SignInCount).IsEqualTo(1);
        await Assert.That(stack.session.IsLoggedIn).IsTrue();
        await Assert.That(stack.gate.LastDiagnostics!.SilentRefreshFailureReason)
            .Contains("An error occurred while sending the request.");
    }

    [Test]
    public async Task LoginFailure_ThrowsFailedNotCancelled()
    {
        var stack = CreateGate();
        stack.prefs.Stored = ValidPrefs();
        stack.creds.Save(StoredCredentials());
        stack.interactive.NextResult = InteractiveLoginResult.Failed("token exchange failed");

        await Assert.That(() => stack.gate.EnsureLoggedInAsync())
            .Throws<SalesforceLoginFailedException>();
    }

    [Test]
    public async Task EnsureLoggedIn_LoadsDisplayName_FromIdentity()
    {
        using var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        client.Handler.OnPost("/services/oauth2/token")
            .RespondWithJson(SalesforceMockResponses.RefreshSuccess);
        client.Handler.OnGet("/services/data/")
            .RespondWithJson("""[{"version":"66.0","url":"/services/data/v66.0/"}]""");
        client.Handler.OnGet("/id/00Dxx/005yy?version=latest")
            .RespondWithJson(SalesforceMockResponses.IdentitySuccess);

        var stack = CreateGate(client, identityClient: new SalesforceRestAddin.Core.Rest.SalesforceIdentityClient(client));
        stack.prefs.Stored = ValidPrefs();
        stack.creds.Save(StoredCredentials());

        await stack.gate.EnsureLoggedInAsync();

        await Assert.That(stack.session.IsLoggedIn).IsTrue();
        await Assert.That(stack.session.DisplayName).IsEqualTo("Test User");
    }

    [Test]
    public async Task JwtAccessToken_HydratesSession_WithoutRefreshOrInteractive()
    {
        var stack = CreateGate();
        stack.prefs.Stored = new UserLoginPreferences
        {
            Tenant = "kjr",
            Environment = SalesforceEnvironment.Production,
            ApiVersion = "66.0",
            JwtAccessToken = "jwt-debug-token",
            ShowLoginOptionsOnNextUse = false,
        };
        stack.creds.Save(StoredCredentials());

        await stack.gate.EnsureLoggedInAsync();

        await Assert.That(stack.session.IsLoggedIn).IsTrue();
        await Assert.That(stack.session.IsJwtAccessTokenMode).IsTrue();
        await Assert.That(stack.session.AccessToken).IsEqualTo("jwt-debug-token");
        await Assert.That(stack.session.InstanceUrl).IsEqualTo("https://kjr.my.salesforce.com");
        await Assert.That(stack.session.RefreshToken).IsNull();
        await Assert.That(stack.session.ApiVersion).IsEqualTo("66.0");
        await Assert.That(stack.interactive.SignInCount).IsEqualTo(0);
        await Assert.That(stack.presenter.ShowCount).IsEqualTo(0);
        await Assert.That(stack.creds.Load("kjr.my.salesforce.com")!.RefreshToken).IsEqualTo("saved-refresh");
        await Assert.That(stack.prefs.Stored!.JwtAccessToken).IsEqualTo("jwt-debug-token");
        await Assert.That(stack.gate.LastDiagnostics!.SilentRefreshAttempted).IsFalse();
    }

    [Test]
    public async Task JwtAccessToken_WithoutTenant_Fails()
    {
        var stack = CreateGate();
        stack.prefs.Stored = new UserLoginPreferences
        {
            Environment = SalesforceEnvironment.Production,
            ApiVersion = "66.0",
            JwtAccessToken = "jwt-debug-token",
        };

        var ex = await Assert.That(() => stack.gate.EnsureLoggedInAsync())
            .Throws<SalesforceLoginFailedException>();
        await Assert.That(ex!.Message).Contains("tenant");
        await Assert.That(stack.interactive.SignInCount).IsEqualTo(0);
    }

    [Test]
    public async Task JwtAccessToken_SandboxWithoutSandboxId_Fails()
    {
        var stack = CreateGate();
        stack.prefs.Stored = new UserLoginPreferences
        {
            Tenant = "kjr",
            Environment = SalesforceEnvironment.Sandbox,
            ApiVersion = "66.0",
            JwtAccessToken = "jwt-debug-token",
        };

        var ex = await Assert.That(() => stack.gate.EnsureLoggedInAsync())
            .Throws<SalesforceLoginFailedException>();
        await Assert.That(ex!.Message).Contains("sandboxId");
        await Assert.That(stack.interactive.SignInCount).IsEqualTo(0);
    }

    [Test]
    public async Task Logout_PreservesJwtAccessToken_InPreferences()
    {
        var stack = CreateGate();
        stack.prefs.Stored = new UserLoginPreferences
        {
            Tenant = "kjr",
            Environment = SalesforceEnvironment.Production,
            ApiVersion = "66.0",
            JwtAccessToken = "jwt-debug-token",
        };
        await stack.gate.EnsureLoggedInAsync();

        stack.gate.Logout();

        await Assert.That(stack.session.IsLoggedIn).IsFalse();
        await Assert.That(stack.session.IsJwtAccessTokenMode).IsFalse();
        await Assert.That(stack.prefs.Stored!.JwtAccessToken).IsEqualTo("jwt-debug-token");
        await Assert.That(stack.prefs.Stored.ShowLoginOptionsOnNextUse).IsTrue();
    }

    private static HttpClient CreateSalesforceMockClient(bool includeTokenRefresh = true)
    {
        var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        if (includeTokenRefresh)
        {
            client.Handler.OnPost("/services/oauth2/token")
                .RespondWithJson(SalesforceMockResponses.RefreshSuccess);
        }

        client.Handler.OnGet("/services/data/")
            .RespondWithJson("""[{"version":"66.0","url":"/services/data/v66.0/"}]""");
        return client;
    }

    private static InteractiveLoginResult InteractiveSuccess()
    {
        var issuedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        return InteractiveLoginResult.Success(
            OAuthTokenResponse.ParseJson(
                $$"""
                {
                  "access_token": "interactive-access",
                  "refresh_token": "interactive-refresh",
                  "instance_url": "https://kjr.my.salesforce.com",
                  "id": "https://login.salesforce.com/id/00Dxx/005yy",
                  "issued_at": "{{issuedAt}}"
                }
                """),
            new SalesforceLoginOptions { Tenant = "kjr", Environment = SalesforceEnvironment.Production });
    }

    private static LoginOptionsPresenterResult PresenterResult(bool forceSignInAgain = false) =>
        new()
        {
            LoginOptions = new SalesforceLoginOptions
            {
                Tenant = "kjr",
                Environment = SalesforceEnvironment.Production,
            },
            ForceSignInAgain = forceSignInAgain,
        };

    private static UserLoginPreferences ValidPrefs(bool showLoginOptionsOnNextUse = false) =>
        new()
        {
            Tenant = "kjr",
            Environment = SalesforceEnvironment.Production,
            ShowLoginOptionsOnNextUse = showLoginOptionsOnNextUse,
        };

    private static StoredSessionCredentials StoredCredentials() =>
        new()
        {
            HostKey = "kjr.my.salesforce.com",
            RefreshToken = "saved-refresh",
        };

    private static GateTestStack CreateGate(
        HttpClient? http = null,
        IUserLoginPreferencesStore? preferencesStore = null,
        SalesforceRestAddin.Core.Rest.SalesforceIdentityClient? identityClient = null)
    {
        http ??= new HttpClient();
        var session = new SessionContext();
        var creds = new InMemorySessionCredentialStore();
        var prefs = preferencesStore as InMemoryUserLoginPreferencesStore ?? new InMemoryUserLoginPreferencesStore();
        var presenter = new FakeLoginOptionsPresenter();
        var interactive = new FakeInteractiveLoginHandler();
        var authenticator = new SessionAuthenticator(creds, new SalesforceOAuthClient(http));
        var orchestrator = new SessionLoginOrchestrator(session, authenticator, interactive);
        var gate = new SessionGate(
            session,
            preferencesStore ?? prefs,
            presenter,
            orchestrator,
            identityClient);
        return new GateTestStack(gate, presenter, interactive, prefs, creds, http, session);
    }

    private sealed record GateTestStack(
        SessionGate gate,
        FakeLoginOptionsPresenter presenter,
        FakeInteractiveLoginHandler interactive,
        InMemoryUserLoginPreferencesStore prefs,
        InMemorySessionCredentialStore creds,
        HttpClient http,
        SessionContext session);
}
