using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class SalesforceInstanceHostTests
{
    [Test]
    public async Task TryParse_ProductionMyDomainUrl_WithOrWithoutScheme()
    {
        var withScheme = SalesforceInstanceHost.TryParse("https://kjr.my.salesforce.com/lightning/page/home");
        var withoutScheme = SalesforceInstanceHost.TryParse("kjr.my.salesforce.com");

        await Assert.That(withScheme).IsNotNull();
        await Assert.That(withScheme!.Tenant).IsEqualTo("kjr");
        await Assert.That(withScheme.Environment).IsEqualTo(SalesforceEnvironment.Production);
        await Assert.That(withScheme.SandboxId).IsNull();

        await Assert.That(withoutScheme!.HostKey).IsEqualTo("kjr.my.salesforce.com");
    }

    [Test]
    public async Task TryParse_NamedSandboxHost()
    {
        var host = SalesforceInstanceHost.TryParse("https://kjr--kjr2026.sandbox.my.salesforce.com");

        await Assert.That(host).IsNotNull();
        await Assert.That(host!.Tenant).IsEqualTo("kjr");
        await Assert.That(host.SandboxId).IsEqualTo("kjr2026");
        await Assert.That(host.Environment).IsEqualTo(SalesforceEnvironment.Sandbox);
    }

    [Test]
    public async Task TryParse_SandboxMyDomainWithoutNamedSandbox()
    {
        var host = SalesforceInstanceHost.TryParse("https://kjr.sandbox.my.salesforce.com");

        await Assert.That(host).IsNotNull();
        await Assert.That(host!.Tenant).IsEqualTo("kjr");
        await Assert.That(host.SandboxId).IsNull();
        await Assert.That(host.Environment).IsEqualTo(SalesforceEnvironment.Sandbox);
    }

    [Test]
    public async Task TryParse_GenericLoginHosts()
    {
        var production = SalesforceInstanceHost.TryParse("login.salesforce.com");
        var sandbox = SalesforceInstanceHost.TryParse("https://test.salesforce.com");

        await Assert.That(production!.IsGenericLoginHost).IsTrue();
        await Assert.That(production.HasResolvableTenant).IsFalse();
        await Assert.That(sandbox!.Environment).IsEqualTo(SalesforceEnvironment.Sandbox);
    }

    [Test]
    public async Task TryParse_ReturnsNull_ForUnknownHost()
    {
        await Assert.That(SalesforceInstanceHost.TryParse("https://example.com")).IsNull();
        await Assert.That(SalesforceInstanceHost.TryParse(string.Empty)).IsNull();
    }

    [Test]
    public async Task ToLoginOptions_MapsParsedHost()
    {
        var options = SalesforceInstanceHost.TryParse("kjr--dev.sandbox.my.salesforce.com")!.ToLoginOptions();

        await Assert.That(options.Tenant).IsEqualTo("kjr");
        await Assert.That(options.SandboxId).IsEqualTo("dev");
        await Assert.That(options.Environment).IsEqualTo(SalesforceEnvironment.Sandbox);
    }
}

public sealed class SalesforceLoginOptionsNormalizerTests
{
    [Test]
    public async Task Normalize_PastedUrl_SetsTenantAndEnvironment()
    {
        var options = SalesforceLoginOptionsNormalizer.Normalize(new SalesforceLoginOptions
        {
            Tenant = "https://kjr--dev.sandbox.my.salesforce.com",
            Environment = SalesforceEnvironment.Production,
        });

        await Assert.That(options.Tenant).IsEqualTo("kjr");
        await Assert.That(options.SandboxId).IsEqualTo("dev");
        await Assert.That(options.Environment).IsEqualTo(SalesforceEnvironment.Sandbox);
    }

    [Test]
    public async Task Normalize_GenericLoginUrl_ClearsTenant()
    {
        var options = SalesforceLoginOptionsNormalizer.Normalize(new SalesforceLoginOptions
        {
            Tenant = "https://login.salesforce.com",
        });

        await Assert.That(options.Tenant).IsNull();
        await Assert.That(options.Environment).IsEqualTo(SalesforceEnvironment.Production);
    }

    [Test]
    public async Task Normalize_TestLoginUrl_SelectsSandboxEnvironment()
    {
        var options = SalesforceLoginOptionsNormalizer.Normalize(new SalesforceLoginOptions
        {
            Tenant = "test.salesforce.com",
            Environment = SalesforceEnvironment.Production,
        });

        await Assert.That(options.Tenant).IsNull();
        await Assert.That(options.Environment).IsEqualTo(SalesforceEnvironment.Sandbox);
    }

    [Test]
    public async Task Normalize_PlainTenant_PreservesEnvironmentAndSandboxId()
    {
        var options = SalesforceLoginOptionsNormalizer.Normalize(new SalesforceLoginOptions
        {
            Tenant = "KJR",
            Environment = SalesforceEnvironment.Sandbox,
            SandboxId = "DEV",
        });

        await Assert.That(options.Tenant).IsEqualTo("kjr");
        await Assert.That(options.SandboxId).IsEqualTo("dev");
        await Assert.That(options.Environment).IsEqualTo(SalesforceEnvironment.Sandbox);
    }

    [Test]
    public async Task Normalize_EmptyTenant_KeepsSandboxIdWhenProvided()
    {
        var options = SalesforceLoginOptionsNormalizer.Normalize(new SalesforceLoginOptions
        {
            Environment = SalesforceEnvironment.Sandbox,
            SandboxId = "dev",
        });

        await Assert.That(options.Tenant).IsNull();
        await Assert.That(options.SandboxId).IsEqualTo("dev");
    }
}

public sealed class SessionLoginReconcilerTests
{
    [Test]
    public async Task Reconcile_GenericLogin_ResolvesTenantFromInstance()
    {
        var reconciliation = SessionLoginReconciler.Reconcile(
            new SalesforceLoginOptions { Environment = SalesforceEnvironment.Production },
            new OAuthTokenResponse
            {
                AccessToken = "token",
                InstanceUrl = "https://kjr.my.salesforce.com",
                Id = "https://login.salesforce.com/id/00D/005",
            });

        await Assert.That(reconciliation.OAuthTokenHost).IsEqualTo("login.salesforce.com");
        await Assert.That(reconciliation.CanonicalHostKey).IsEqualTo("kjr.my.salesforce.com");
        await Assert.That(reconciliation.ResolvedLoginOptions.Tenant).IsEqualTo("kjr");
        await Assert.That(reconciliation.ResolvedLoginOptions.Environment)
            .IsEqualTo(SalesforceEnvironment.Production);
    }

    [Test]
    public async Task Reconcile_GenericSandboxLogin_ResolvesNamedSandboxFromInstance()
    {
        var reconciliation = SessionLoginReconciler.Reconcile(
            new SalesforceLoginOptions { Environment = SalesforceEnvironment.Sandbox },
            new OAuthTokenResponse
            {
                AccessToken = "token",
                InstanceUrl = "https://kjr--kjr2026.sandbox.my.salesforce.com",
                Id = "https://test.salesforce.com/id/00D/005",
            });

        await Assert.That(reconciliation.OAuthTokenHost).IsEqualTo("test.salesforce.com");
        await Assert.That(reconciliation.CanonicalHostKey).IsEqualTo("kjr--kjr2026.sandbox.my.salesforce.com");
        await Assert.That(reconciliation.ResolvedLoginOptions.Tenant).IsEqualTo("kjr");
        await Assert.That(reconciliation.ResolvedLoginOptions.SandboxId).IsEqualTo("kjr2026");
    }

    [Test]
    public async Task Reconcile_TenantSpecificLogin_KeepsSameOAuthHost()
    {
        var reconciliation = SessionLoginReconciler.Reconcile(
            new SalesforceLoginOptions { Tenant = "kjr" },
            new OAuthTokenResponse
            {
                AccessToken = "token",
                InstanceUrl = "https://kjr.my.salesforce.com",
                Id = "https://login.salesforce.com/id/00D/005",
            });

        await Assert.That(reconciliation.OAuthTokenHost).IsEqualTo("kjr.my.salesforce.com");
        await Assert.That(reconciliation.CanonicalHostKey).IsEqualTo("kjr.my.salesforce.com");
    }
}

public sealed class SalesforceLoginUrlBuilderNormalizationTests
{
    [Test]
    public async Task BuildLoginBaseUrl_AcceptsTenantFieldUrl()
    {
        var url = SalesforceLoginUrlBuilder.BuildLoginBaseUrl(new SalesforceLoginOptions
        {
            Tenant = "https://kjr--dev.sandbox.my.salesforce.com",
        });

        await Assert.That(url).IsEqualTo("https://kjr--dev.sandbox.my.salesforce.com");
    }

    [Test]
    public async Task SessionHostKey_MatchesParsedTenantUrl()
    {
        var hostKey = SessionHostKey.FromLoginOptions(new SalesforceLoginOptions
        {
            Tenant = "kjr.my.salesforce.com",
        });

        await Assert.That(hostKey).IsEqualTo("kjr.my.salesforce.com");
    }
}
