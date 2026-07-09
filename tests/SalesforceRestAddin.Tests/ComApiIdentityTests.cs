using SalesforceRestAddin.Core;

namespace SalesforceRestAddin.Tests;

public class ComApiIdentityTests
{
    [Test]
    public async Task Automation_ClassId_Matches_Legacy_VSTO_Value()
    {
        await Assert.That(ComApiIdentity.AutomationClassId)
            .IsEqualTo("e5fcccc8-a685-4980-a79a-eab37f2c7caf");
    }

    [Test]
    public async Task Automation_InterfaceId_Matches_Legacy_VSTO_Value()
    {
        await Assert.That(ComApiIdentity.AutomationInterfaceId)
            .IsEqualTo("bacb17f7-9b85-4137-861d-c9a2d899b564");
    }

    [Test]
    public async Task Automation_EventsId_Matches_Legacy_VSTO_Value()
    {
        await Assert.That(ComApiIdentity.AutomationEventsId)
            .IsEqualTo("6894737e-cfa7-410c-be80-03ad4de47656");
    }

    [Test]
    public async Task ComAddIn_ProgId_Matches_ConnectorAdaptor()
    {
        await Assert.That(ComApiIdentity.ComAddInProgId)
            .IsEqualTo("ForceConnector.NextGen");
    }
}
