using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;
using TUnit.Mocks;

namespace SalesforceRestAddin.Tests;

public sealed class SalesforceApiVersionResolverTests
{
    [Test]
    public async Task ResolveForSessionAsync_UsesPreferredVersionWhenSupported()
    {
        using var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        client.Handler.OnGet("/services/data/")
            .RespondWithJson("""
                [
                  {"label":"Winter '26","url":"/services/data/v66.0/","version":"66.0"},
                  {"label":"Spring '26","url":"/services/data/v67.0/","version":"67.0"}
                ]
                """);

        var resolver = new SalesforceApiVersionResolver(client);
        var resolution = await resolver.ResolveForSessionAsync(
            "https://kjr.my.salesforce.com",
            "token",
            preferredVersion: "66.0");

        await Assert.That(resolution.SelectedVersion).IsEqualTo("66.0");
        await Assert.That(resolution.SupportedVersions).IsEquivalentTo(new[] { "67.0", "66.0" });
    }
}

public sealed class SessionAuthenticatorApiVersionTests
{
    [Test]
    public async Task ApplyToSessionAsync_UsesPreferredApiVersionWhenSupported()
    {
        using var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        client.Handler.OnGet("/services/data/")
            .RespondWithJson("""
                [
                  {"label":"Winter '26","url":"/services/data/v66.0/","version":"66.0"},
                  {"label":"Spring '26","url":"/services/data/v67.0/","version":"67.0"}
                ]
                """);

        var authenticator = new SessionAuthenticator(
            new InMemorySessionCredentialStore(),
            new SalesforceOAuthClient(client));

        var session = new SessionContext();
        var resolution = await authenticator.ApplyToSessionAsync(
            new OAuthTokenResponse
            {
                AccessToken = "token",
                InstanceUrl = "https://kjr.my.salesforce.com",
                Id = "https://login.salesforce.com/id/00D/005",
            },
            session,
            preferredApiVersion: "66.0");

        await Assert.That(resolution.SelectedVersion).IsEqualTo("66.0");
        await Assert.That(session.ApiVersion).IsEqualTo("66.0");
    }
}
