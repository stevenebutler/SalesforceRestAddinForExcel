using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class SessionHostKeyTests
{
    [Test]
    public async Task ProductionTenant_UsesMyDomainHost()
    {
        var key = SessionHostKey.FromLoginOptions(new SalesforceLoginOptions { Tenant = "kjr" });
        await Assert.That(key).IsEqualTo("kjr.my.salesforce.com");
    }

    [Test]
    public async Task SandboxTenantAndId_UsesDistinctHostFromProduction()
    {
        var production = SessionHostKey.FromLoginOptions(new SalesforceLoginOptions { Tenant = "kjr" });
        var sandbox = SessionHostKey.FromLoginOptions(new SalesforceLoginOptions
        {
            Tenant = "kjr",
            SandboxId = "dev",
            Environment = SalesforceEnvironment.Sandbox,
        });

        await Assert.That(production).IsNotEqualTo(sandbox);
        await Assert.That(sandbox).IsEqualTo("kjr--dev.sandbox.my.salesforce.com");
    }
}
