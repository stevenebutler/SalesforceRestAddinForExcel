using SalesforceRestAddin.Core;

namespace SalesforceRestAddin.Tests;

public class SessionContextTests
{
    [Test]
    public async Task IsLoggedIn_False_When_Session_Is_Empty()
    {
        var session = new SessionContext();
        await Assert.That(session.IsLoggedIn).IsFalse();
    }

    [Test]
    public async Task IsLoggedIn_True_When_Token_And_InstanceUrl_Present()
    {
        var session = new SessionContext
        {
            AccessToken = "00D...token",
            InstanceUrl = "https://kjr.my.salesforce.com",
            IssuedAtUnixMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
        };
        session.SetApiVersion("66.0");

        await Assert.That(session.IsLoggedIn).IsTrue();
    }

    [Test]
    public async Task IsLoggedIn_False_When_Token_Expired()
    {
        var session = new SessionContext
        {
            AccessToken = "00D...token",
            InstanceUrl = "https://kjr.my.salesforce.com",
            IssuedAtUnixMs = DateTimeOffset.UtcNow.AddHours(-3).ToUnixTimeMilliseconds(),
        };
        session.SetApiVersion("66.0");

        await Assert.That(session.IsLoggedIn).IsFalse();
    }

    [Test]
    public async Task ApiVersion_Throws_UntilResolved()
    {
        var session = new SessionContext();

        await Assert.That(() => _ = session.ApiVersion).Throws<InvalidOperationException>();
        await Assert.That(session.IsApiVersionResolved).IsFalse();

        session.SetApiVersion("67.0");

        await Assert.That(session.ApiVersion).IsEqualTo("67.0");
        await Assert.That(session.IsApiVersionResolved).IsTrue();
    }

    [Test]
    public async Task Invalidate_Clears_Session_Fields()
    {
        var session = new SessionContext
        {
            AccessToken = "token",
            RefreshToken = "refresh",
            InstanceUrl = "https://example.my.salesforce.com",
            Id = "id",
            DisplayName = "Test User",
            IsJwtAccessTokenMode = true,
            IssuedAtUnixMs = 1,
        };
        session.SetApiVersion("66.0");

        session.Invalidate();

        await Assert.That(session.AccessToken).IsNull();
        await Assert.That(session.RefreshToken).IsNull();
        await Assert.That(session.InstanceUrl).IsNull();
        await Assert.That(session.Id).IsNull();
        await Assert.That(session.DisplayName).IsNull();
        await Assert.That(session.IssuedAtUnixMs).IsNull();
        await Assert.That(session.IsJwtAccessTokenMode).IsFalse();
        await Assert.That(session.IsApiVersionResolved).IsFalse();
        await Assert.That(() => _ = session.ApiVersion).Throws<InvalidOperationException>();
        await Assert.That(session.IsLoggedIn).IsFalse();
    }
}
