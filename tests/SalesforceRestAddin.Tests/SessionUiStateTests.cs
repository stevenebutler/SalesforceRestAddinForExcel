using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class SessionUiStateTests
{
    [Test]
    public async Task NoTarget_ShowsNoneAndNotSignedIn_LogoutDisabled_OptionsEnabled()
    {
        var prefs = new InMemoryUserLoginPreferencesStore { Stored = null };
        var session = new SessionContext();

        var state = SessionUiState.From(prefs, session);

        await Assert.That(state.HasTarget).IsFalse();
        await Assert.That(state.LogoutEnabled).IsFalse();
        await Assert.That(state.OptionsEnabled).IsTrue();
        await Assert.That(state.GroupLabel).IsEqualTo(
            "Instance: (none)  User: (not signed in)");
    }

    [Test]
    public async Task AfterLogoutFlag_TreatedAsNoTarget()
    {
        var prefs = new InMemoryUserLoginPreferencesStore
        {
            Stored = new UserLoginPreferences
            {
                Tenant = "kjr",
                Environment = SalesforceEnvironment.Production,
                ShowLoginOptionsOnNextUse = true,
            },
        };
        var session = new SessionContext();

        var state = SessionUiState.From(prefs, session);

        await Assert.That(state.HasTarget).IsFalse();
        await Assert.That(state.LogoutEnabled).IsFalse();
        await Assert.That(state.GroupLabel).IsEqualTo(
            "Instance: (none)  User: (not signed in)");
    }

    [Test]
    public async Task CommittedTarget_NotLoggedIn_ShowsHostAndNotSignedIn()
    {
        var prefs = new InMemoryUserLoginPreferencesStore
        {
            Stored = new UserLoginPreferences
            {
                Tenant = "kjr",
                Environment = SalesforceEnvironment.Production,
                ShowLoginOptionsOnNextUse = false,
            },
        };
        var session = new SessionContext();

        var state = SessionUiState.From(prefs, session);

        await Assert.That(state.HasTarget).IsTrue();
        await Assert.That(state.LogoutEnabled).IsTrue();
        await Assert.That(state.IsLoggedIn).IsFalse();
        await Assert.That(state.OptionsEnabled).IsTrue();
        await Assert.That(state.InstanceHost).IsEqualTo("kjr.my.salesforce.com");
        await Assert.That(state.GroupLabel).IsEqualTo(
            "Instance: kjr.my.salesforce.com  User: (not signed in)");
    }

    [Test]
    public async Task LoggedIn_WithDisplayName_ShowsHostAndUser()
    {
        var prefs = new InMemoryUserLoginPreferencesStore
        {
            Stored = new UserLoginPreferences
            {
                Tenant = "kjr",
                Environment = SalesforceEnvironment.Production,
                ShowLoginOptionsOnNextUse = false,
            },
        };
        var session = new SessionContext
        {
            AccessToken = "token",
            Id = "https://login.salesforce.com/id/00D/005",
            InstanceUrl = "https://kjr.my.salesforce.com",
            DisplayName = "Ada Lovelace",
            IssuedAtUnixMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
        };
        session.SetApiVersion("66.0");

        var state = SessionUiState.From(prefs, session);

        await Assert.That(state.HasTarget).IsTrue();
        await Assert.That(state.LogoutEnabled).IsTrue();
        await Assert.That(state.IsLoggedIn).IsTrue();
        await Assert.That(state.GroupLabel).IsEqualTo(
            "Instance: kjr.my.salesforce.com  User: Ada Lovelace");
    }

    [Test]
    public async Task LoggedIn_WithoutDisplayName_ShowsHostAndUnknown()
    {
        var prefs = new InMemoryUserLoginPreferencesStore { Stored = null };
        var session = new SessionContext
        {
            AccessToken = "token",
            Id = "https://login.salesforce.com/id/00D/005",
            InstanceUrl = "https://kjr.my.salesforce.com",
            IssuedAtUnixMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
        };
        session.SetApiVersion("66.0");

        var state = SessionUiState.From(prefs, session);

        await Assert.That(state.HasTarget).IsTrue();
        await Assert.That(state.GroupLabel).IsEqualTo(
            "Instance: kjr.my.salesforce.com  User: (unknown)");
    }

    [Test]
    public async Task PreferLiveInstanceHost_WhenLoggedIn()
    {
        var prefs = new InMemoryUserLoginPreferencesStore
        {
            Stored = new UserLoginPreferences
            {
                Tenant = "kjr",
                Environment = SalesforceEnvironment.Production,
                ShowLoginOptionsOnNextUse = false,
            },
        };
        var session = new SessionContext
        {
            AccessToken = "token",
            Id = "https://login.salesforce.com/id/00D/005",
            InstanceUrl = "https://other.my.salesforce.com",
            DisplayName = "User",
            IssuedAtUnixMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
        };
        session.SetApiVersion("66.0");

        var state = SessionUiState.From(prefs, session);

        await Assert.That(state.InstanceHost).IsEqualTo("other.my.salesforce.com");
        await Assert.That(state.GroupLabel).IsEqualTo(
            "Instance: other.my.salesforce.com  User: User");
    }
}
