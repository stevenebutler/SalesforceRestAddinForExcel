using SalesforceRestAddin.Installer;

namespace SalesforceRestAddin.Tests;

public sealed class ExcelStartupRegistrationTests
{
    [Test]
    [Arguments("\"Z:\\Salesforce Development\\SalesforceRestAddin-AddIn64-packed.xll\"")]
    [Arguments("Z:\\Salesforce Development\\SalesforceRestAddin-AddIn64-packed.xll")]
    [Arguments("/R \"Z:\\Salesforce Development\\SalesforceRestAddin-AddIn64-packed.xll\"")]
    [Arguments("z:\\build\\salesforcerestaddin64-packed.XLL /R")]
    public async Task Salesforce_Xll_Matcher_Recognizes_Development_Registrations(string value)
    {
        await Assert.That(ExcelStartupRegistration.IsSalesforceRestAddinXll(value)).IsTrue();
    }

    [Test]
    [Arguments("\"Z:\\SalesforceRestAddinForExcel\\OtherAddIn.xll\"")]
    [Arguments("\"C:\\Addins\\ForceConnector.xll\"")]
    [Arguments("\"C:\\Addins\\SalesforceRestAddin.txt\"")]
    public async Task Salesforce_Xll_Matcher_Preserves_Unrelated_Addins(string value)
    {
        await Assert.That(ExcelStartupRegistration.IsSalesforceRestAddinXll(value)).IsFalse();
    }

    [Test]
    public async Task Legacy_ForceConnector_Key_Is_Recognized_Without_Matching_Compatibility_ProgId()
    {
        await Assert.That(ExcelStartupRegistration.IsLegacyForceConnectorAddinKey("ForceConnector")).IsTrue();
        await Assert.That(ExcelStartupRegistration.IsLegacyForceConnectorAddinKey("ForceConnector.NextGen")).IsFalse();
    }

    [Test]
    public async Task Open_Value_Names_Include_Only_Open_And_Numbered_Open_Values()
    {
        await Assert.That(ExcelStartupRegistration.IsOpenValueName("OPEN")).IsTrue();
        await Assert.That(ExcelStartupRegistration.IsOpenValueName("open12")).IsTrue();
        await Assert.That(ExcelStartupRegistration.IsOpenValueName("OPENX")).IsFalse();
    }
}
