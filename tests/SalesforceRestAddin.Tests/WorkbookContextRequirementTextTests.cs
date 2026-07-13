using SalesforceRestAddin.Core;

namespace SalesforceRestAddin.Tests;

public sealed class WorkbookContextRequirementTextTests
{
    [Test]
    public async Task Workbook_And_Worksheet_Use_Create_Or_Open_Message()
    {
        await Assert.That(WorkbookContextRequirementText.GetMessage(WorkbookContextRequirement.Workbook))
            .IsEqualTo("Create or open a workbook, then try again.");

        await Assert.That(WorkbookContextRequirementText.GetMessage(WorkbookContextRequirement.Worksheet))
            .IsEqualTo("Create or open a workbook, then try again.");
    }

    [Test]
    public async Task Salesforce_Connector_Table_Uses_Table_Message()
    {
        var message = WorkbookContextRequirementText.GetMessage(
            WorkbookContextRequirement.SalesforceConnectorTable);

        await Assert.That(message)
            .Contains("Salesforce Connector");
        await Assert.That(message)
            .IsEqualTo("Open a workbook containing a Salesforce Connector table, then try again.");
    }
}
