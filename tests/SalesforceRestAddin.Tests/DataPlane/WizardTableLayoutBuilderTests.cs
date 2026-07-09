using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class WizardTableLayoutBuilderTests
{
    [Test]
    public async Task T_WIZ_01_Field_Order_Id_Required_Name_Standard_Custom_ReadOnly()
    {
        var describe = new SObjectDescribe
        {
            Name = "Account",
            Label = "Account",
            Fields =
            [
                new FieldDescriptor
                {
                    Name = "Custom__c",
                    Label = "Custom Field",
                    Type = "string",
                    Createable = true,
                    Updateable = true,
                    Nillable = true,
                    Custom = true,
                },
                new FieldDescriptor
                {
                    Name = "Industry",
                    Label = "Industry",
                    Type = "picklist",
                    Createable = true,
                    Updateable = true,
                    Nillable = true,
                },
                new FieldDescriptor
                {
                    Name = "CreatedDate",
                    Label = "Created Date",
                    Type = "datetime",
                    Createable = false,
                    Updateable = false,
                    Nillable = true,
                },
                new FieldDescriptor
                {
                    Name = "Name",
                    Label = "Account Name",
                    Type = "string",
                    Createable = true,
                    Updateable = true,
                    Nillable = true,
                },
                new FieldDescriptor
                {
                    Name = "Id",
                    Label = "Account ID",
                    Type = "id",
                    Createable = false,
                    Updateable = false,
                    Nillable = false,
                },
                new FieldDescriptor
                {
                    Name = "ExternalId__c",
                    Label = "External Id",
                    Type = "string",
                    Createable = true,
                    Updateable = true,
                    Nillable = false,
                    Custom = true,
                },
            ],
        };

        var ordered = WizardTableLayoutBuilder.OrderFieldsForWizard(describe);
        var names = ordered.Select(f => f.Name).ToArray();

        await Assert.That(names).IsEquivalentTo(
        [
            "Id",
            "ExternalId__c",
            "Name",
            "Industry",
            "Custom__c",
            "CreatedDate",
        ]);
    }

    [Test]
    public async Task T_WIZ_01_Account_Describe_Id_Field_Ordered_First()
    {
        var describe = DescribeFixtures.LoadAccountDescribe();
        var ordered = WizardTableLayoutBuilder.OrderFieldsForWizard(describe);
        await Assert.That(ordered[0].IsId).IsTrue();
    }

    [Test]
    public async Task T_WIZ_02_Default_Where_Clause()
    {
        await Assert.That(WizardTableLayoutBuilder.DefaultWhereClause()).IsEqualTo("Id != null");
    }

    [Test]
    public async Task T_WIZ_03_Header_Row_Aligns_Id_With_Object_Column()
    {
        var describe = DescribeFixtures.LoadAccountDescribe();
        var fields = WizardTableLayoutBuilder.OrderFieldsForWizard(describe)
            .Where(f => f.IsId || string.Equals(f.Name, "Name", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var labels = WizardTableLayoutBuilder.BuildHeaderRowLabels(fields);
        var apiNames = fields.Select(f => (string?)f.Name).ToArray();

        var snapshot = new ForceTableSnapshot
        {
            ObjectApiName = "Account",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = labels,
            HeaderApiNames = apiNames,
            Body = new object?[0, labels.Length],
        };

        var result = ForceTableBinder.Bind(snapshot, describe);

        await Assert.That(result.Succeeded).IsTrue();
        await Assert.That(result.Binding!.IdColumnIndex).IsEqualTo(0);
        await Assert.That(labels[0]).IsEqualTo("Account ID");
    }
}
