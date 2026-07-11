using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class WizardObjectListFilterTests
{
    [Test]
    public async Task T_WIZ_OBJ_01_Classifies_Standard_Custom_System_And_Hidden_System()
    {
        var standard = new SObjectSummary { Name = "Account", Label = "Account", Queryable = true, Custom = false };
        var custom = new SObjectSummary { Name = "Custom__c", Label = "Custom Object", Queryable = true, Custom = true };
        var system = new SObjectSummary { Name = "ApexClass", Label = "Apex Class", Queryable = true, Custom = false };
        var hidden = new SObjectSummary
        {
            Name = "Hidden__c",
            Label = "Hidden Object",
            Queryable = true,
            Custom = true,
            DeprecatedAndHidden = true,
        };

        await Assert.That(WizardObjectListFilter.GetKind(standard)).IsEqualTo(WizardObjectKind.Standard);
        await Assert.That(WizardObjectListFilter.GetKind(custom)).IsEqualTo(WizardObjectKind.Custom);
        await Assert.That(WizardObjectListFilter.GetKind(system)).IsEqualTo(WizardObjectKind.System);
        await Assert.That(WizardObjectListFilter.GetKind(hidden)).IsEqualTo(WizardObjectKind.HiddenSystem);
    }

    [Test]
    public async Task T_WIZ_OBJ_02_Extracts_Namespace_With_Underscores_And_SFL5_Special_Case()
    {
        var b3f = new SObjectSummary { Name = "b3f__FormPage__c", Label = "Form Page", Queryable = true, Custom = true };
        var linkedIn = new SObjectSummary
        {
            Name = "RK_LinkedIn__LinkedInConnectorPreferences__c",
            Label = "LinkedIn Connector Preferences",
            Queryable = true,
            Custom = true,
        };
        var sfl5 = new SObjectSummary { Name = "SFL5_Assign__c", Label = "Resource Assignment KJRA", Queryable = true, Custom = true };
        var sfdc = new SObjectSummary { Name = "SFDC_Assign__c", Label = "Resource Assignment", Queryable = true, Custom = true };
        var plainCustom = new SObjectSummary { Name = "Budget_Plan__c", Label = "Budget Plan", Queryable = true, Custom = true };

        await Assert.That(WizardObjectListFilter.GetNamespace(b3f)).IsEqualTo("b3f");
        await Assert.That(WizardObjectListFilter.GetNamespace(linkedIn)).IsEqualTo("RK_LinkedIn");
        await Assert.That(WizardObjectListFilter.GetNamespace(sfl5)).IsEqualTo("SFL5");
        await Assert.That(WizardObjectListFilter.GetNamespace(sfdc)).IsEqualTo("SFDC");
        await Assert.That(WizardObjectListFilter.GetNamespace(plainCustom)).IsNull();
    }

    [Test]
    public async Task T_WIZ_OBJ_03_Maps_Group_For_Standard_Custom_And_Namespaced_Objects()
    {
        var standard = new SObjectSummary { Name = "Account", Label = "Account", Queryable = true, Custom = false };
        var custom = new SObjectSummary { Name = "Budget_Plan__c", Label = "Budget Plan", Queryable = true, Custom = true };
        var namespaced = new SObjectSummary { Name = "b3f__FormPage__c", Label = "Form Page", Queryable = true, Custom = true };
        var sfdc = new SObjectSummary { Name = "SFDC_Assign__c", Label = "Resource Assignment", Queryable = true, Custom = true };

        await Assert.That(WizardObjectListFilter.GetGroup(standard)).IsEqualTo(string.Empty);
        await Assert.That(WizardObjectListFilter.GetGroup(custom)).IsEqualTo(" ");
        await Assert.That(WizardObjectListFilter.GetGroup(namespaced)).IsEqualTo("b3f");
        await Assert.That(WizardObjectListFilter.GetGroup(sfdc)).IsEqualTo("SFDC");
    }

    [Test]
    public async Task T_WIZ_OBJ_04_Filter_Excludes_Hidden_System_Objects_And_Respects_Category_Checkboxes()
    {
        var objects =
            new[]
            {
                new SObjectSummary { Name = "Account", Label = "Account", Queryable = true, Custom = false },
                new SObjectSummary { Name = "Budget_Plan__c", Label = "Budget Plan", Queryable = true, Custom = true },
                new SObjectSummary { Name = "ApexClass", Label = "Apex Class", Queryable = true, Custom = false },
                new SObjectSummary { Name = "Hidden__c", Label = "Hidden", Queryable = true, Custom = true, DeprecatedAndHidden = true },
            };

        var filtered = WizardObjectListFilter.FilterAndSort(
            objects,
            showStandardObjects: true,
            showCustomObjects: true,
            showSystemObjects: false);

        await Assert.That(filtered.Select(o => o.ApiName).SequenceEqual(["Account", "Budget_Plan__c"])).IsTrue();

        var withSystem = WizardObjectListFilter.FilterAndSort(
            objects,
            showStandardObjects: true,
            showCustomObjects: true,
            showSystemObjects: true);

        await Assert.That(withSystem.Select(o => o.ApiName).SequenceEqual(["Account", "ApexClass", "Budget_Plan__c"])).IsTrue();
    }

    [Test]
    public async Task T_WIZ_OBJ_05_Default_Sort_Order_Is_Group_Then_Label()
    {
        var objects =
            new[]
            {
                new SObjectSummary { Name = "ApexClass", Label = "Beta", Queryable = true, Custom = false },
                new SObjectSummary { Name = "Account", Label = "Alpha", Queryable = true, Custom = false },
                new SObjectSummary { Name = "Budget_Plan__c", Label = "Budget Plan", Queryable = true, Custom = true },
                new SObjectSummary { Name = "b3f__FormPage__c", Label = "Form Page", Queryable = true, Custom = true },
            };

        var filtered = WizardObjectListFilter.FilterAndSort(
            objects,
            showStandardObjects: true,
            showCustomObjects: true,
            showSystemObjects: true);

        await Assert.That(filtered.Select(o => o.ApiName).SequenceEqual(
            ["Account", "ApexClass", "Budget_Plan__c", "b3f__FormPage__c"])).IsTrue();
    }
}
