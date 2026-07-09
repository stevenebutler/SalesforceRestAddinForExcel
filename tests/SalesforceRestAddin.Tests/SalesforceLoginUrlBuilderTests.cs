using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class SalesforceLoginUrlBuilderTests
{
    [Test]
    public async Task ProductionTenant_IsNormalizedToLowercase()
    {
        var options = new SalesforceLoginOptions
        {
            Tenant = "KJR",
            Environment = SalesforceEnvironment.Production,
        };

        await Assert.That(SalesforceLoginUrlBuilder.BuildLoginBaseUrl(options))
            .IsEqualTo("https://kjr.my.salesforce.com");
        await Assert.That(SessionHostKey.FromLoginOptions(options))
            .IsEqualTo("kjr.my.salesforce.com");
    }

    [Test]
    public async Task Production_WithTenant_BuildsMyDomainUrl()
    {
        var options = new SalesforceLoginOptions
        {
            Tenant = "kjr",
            Environment = SalesforceEnvironment.Production,
        };

        await Assert.That(SalesforceLoginUrlBuilder.BuildLoginBaseUrl(options))
            .IsEqualTo("https://kjr.my.salesforce.com");
    }

    [Test]
    public async Task Production_WithoutTenant_UsesLoginHost()
    {
        var options = new SalesforceLoginOptions
        {
            Environment = SalesforceEnvironment.Production,
        };

        await Assert.That(SalesforceLoginUrlBuilder.BuildLoginBaseUrl(options))
            .IsEqualTo("https://login.salesforce.com");
    }

    [Test]
    public async Task Sandbox_WithTenantAndSandboxId_BuildsMyDomainSandboxUrl()
    {
        var options = new SalesforceLoginOptions
        {
            Tenant = "kjr",
            SandboxId = "dev",
            Environment = SalesforceEnvironment.Sandbox,
        };

        await Assert.That(SalesforceLoginUrlBuilder.BuildLoginBaseUrl(options))
            .IsEqualTo("https://kjr--dev.sandbox.my.salesforce.com");
    }

    [Test]
    public async Task Sandbox_WithoutTenant_UsesTestHost()
    {
        var options = new SalesforceLoginOptions
        {
            Environment = SalesforceEnvironment.Sandbox,
        };

        await Assert.That(SalesforceLoginUrlBuilder.BuildLoginBaseUrl(options))
            .IsEqualTo("https://test.salesforce.com");
    }

    [Test]
    public async Task BuildAuthorizeUrl_IncludesPkceAndClientId()
    {
        var login = new SalesforceLoginOptions { Tenant = "kjr" };
        var pkce = new PkcePair { Verifier = "verifier", Challenge = "challenge-value" };

        var url = SalesforceLoginUrlBuilder.BuildAuthorizeUrl(login, pkce);

        await Assert.That(url).StartsWith("https://kjr.my.salesforce.com/services/oauth2/authorize?");
        await Assert.That(url).Contains("client_id=PlatformCLI");
        await Assert.That(url).Contains("code_challenge=challenge-value");
        await Assert.That(url).Contains("code_challenge_method=S256");
    }

    [Test]
    public async Task OAuthRedirectParser_ExtractsAuthorizationCode()
    {
        const string redirectUri = "http://localhost:1717/OauthRedirect";
        const string callback = "http://localhost:1717/OauthRedirect?code=abc123&state=x";

        var ok = OAuthRedirectParser.TryGetAuthorizationCode(callback, redirectUri, out var code, out var error);

        await Assert.That(ok).IsTrue();
        await Assert.That(code).IsEqualTo("abc123");
        await Assert.That(error).IsNull();
    }
}
