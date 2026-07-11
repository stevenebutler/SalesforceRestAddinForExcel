using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class ObjectPickerSortStateTests
{
    [Test]
    public async Task T_WIZ_SORT_01_Default_Order_Sorts_By_Group_Then_Label()
    {
        var rows = new[]
        {
            new ObjectPickerRow(
                new SObjectSummary { Name = "ApexClass", Label = "Beta", Queryable = true, Custom = false },
                string.Empty,
                "Beta",
                "ApexClass",
                WizardObjectKind.System),
            new ObjectPickerRow(
                new SObjectSummary { Name = "Account", Label = "Alpha", Queryable = true, Custom = false },
                string.Empty,
                "Alpha",
                "Account",
                WizardObjectKind.Standard),
            new ObjectPickerRow(
                new SObjectSummary { Name = "Budget_Plan__c", Label = "Budget Plan", Queryable = true, Custom = true },
                " ",
                "Budget Plan",
                "Budget_Plan__c",
                WizardObjectKind.Custom),
            new ObjectPickerRow(
                new SObjectSummary { Name = "b3f__FormPage__c", Label = "Form Page", Queryable = true, Custom = true },
                "b3f",
                "Form Page",
                "b3f__FormPage__c",
                WizardObjectKind.Custom),
        };

        var sorted = new ObjectPickerSortState().Sort(rows).Select(row => row.ApiName).ToArray();

        await Assert.That(sorted.SequenceEqual(["Account", "ApexClass", "Budget_Plan__c", "b3f__FormPage__c"])).IsTrue();
    }

    [Test]
    public async Task T_WIZ_SORT_02_Clicking_Another_Column_Moves_It_To_The_Front_Of_The_Sort_Order()
    {
        var state = new ObjectPickerSortState();

        state.Promote(ObjectPickerSortColumn.ApiName);
        await Assert.That(state.Keys.Select(k => k.Column).SequenceEqual(
            [ObjectPickerSortColumn.ApiName, ObjectPickerSortColumn.Group, ObjectPickerSortColumn.Label])).IsTrue();

        state.Promote(ObjectPickerSortColumn.Label);
        await Assert.That(state.Keys.Select(k => k.Column).SequenceEqual(
            [ObjectPickerSortColumn.Label, ObjectPickerSortColumn.ApiName, ObjectPickerSortColumn.Group])).IsTrue();
    }

    [Test]
    public async Task T_WIZ_SORT_03_Clicking_The_Same_Column_Again_Toggles_Direction()
    {
        var state = new ObjectPickerSortState();

        state.Promote(ObjectPickerSortColumn.Group);

        await Assert.That(state.Keys[0]).IsEqualTo(
            new ObjectPickerSortKey(ObjectPickerSortColumn.Group, ObjectPickerSortDirection.Descending));
        await Assert.That(state.Keys.Select(k => k.Column).SequenceEqual(
            [ObjectPickerSortColumn.Group, ObjectPickerSortColumn.Label])).IsTrue();
    }
}
