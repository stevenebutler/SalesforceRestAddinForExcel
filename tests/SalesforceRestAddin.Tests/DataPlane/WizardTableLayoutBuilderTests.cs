using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class WizardTableLayoutBuilderTests
{
    [Test]
    public async Task T_WIZ_01_Field_Order_Id_Name_Required_Standard_Custom_ReadOnly()
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
                    Nillable = false,
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
                new FieldDescriptor
                {
                    Name = "Website",
                    Label = "Website",
                    Type = "url",
                    Createable = true,
                    Updateable = true,
                    Nillable = true,
                },
                new FieldDescriptor
                {
                    Name = "RoCustom__c",
                    Label = "RO Custom",
                    Type = "string",
                    Createable = false,
                    Updateable = false,
                    Nillable = true,
                    Custom = true,
                },
            ],
        };

        var ordered = WizardTableLayoutBuilder.OrderFieldsForWizard(describe);
        var names = ordered.Select(f => f.Name).ToArray();

        // Within Standard bucket, Industry (index 1) before Website (index 6) — describe order.
        await Assert.That(names).IsEquivalentTo(
        [
            "Id",
            "Name",
            "ExternalId__c",
            "Industry",
            "Website",
            "Custom__c",
            "CreatedDate",
            "RoCustom__c",
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
    public async Task T_WIZ_01b_Order_Preserves_Describe_Index_Within_Bucket()
    {
        var describe = new SObjectDescribe
        {
            Name = "Account",
            Label = "Account",
            Fields =
            [
                new FieldDescriptor
                {
                    Name = "Zebra__c",
                    Label = "Zebra",
                    Type = "string",
                    Createable = true,
                    Updateable = true,
                    Nillable = true,
                    Custom = true,
                },
                new FieldDescriptor
                {
                    Name = "Id",
                    Label = "Account ID",
                    Type = "id",
                    Nillable = false,
                },
                new FieldDescriptor
                {
                    Name = "Alpha__c",
                    Label = "Alpha",
                    Type = "string",
                    Createable = true,
                    Updateable = true,
                    Nillable = true,
                    Custom = true,
                },
            ],
        };

        var ordered = WizardTableLayoutBuilder.OrderFieldsForWizard(describe);
        var customNames = ordered
            .Where(f => WizardTableLayoutBuilder.GetFieldBucket(f) == 5)
            .Select(f => f.Name)
            .ToArray();

        await Assert.That(customNames).IsEquivalentTo(["Zebra__c", "Alpha__c"]);
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

    [Test]
    public async Task T_WIZ_04_Header_Fill_Rgb_Distinct_Per_Bucket()
    {
        var fills = Enumerable.Range(0, WizardTableLayoutBuilder.BucketCount)
            .Select(WizardTableLayoutBuilder.HeaderFillRgb)
            .ToArray();

        await Assert.That(fills[0]).IsEqualTo(((byte)189, (byte)215, (byte)238));
        await Assert.That(fills[1]).IsEqualTo(((byte)255, (byte)230, (byte)153));
        await Assert.That(fills[2]).IsEqualTo(((byte)248, (byte)203, (byte)173));
        await Assert.That(fills[3]).IsEqualTo(((byte)230, (byte)160, (byte)120));
        await Assert.That(fills[4]).IsEqualTo(((byte)198, (byte)239, (byte)206));
        await Assert.That(fills[5]).IsEqualTo(((byte)226, (byte)213, (byte)241));
        await Assert.That(fills[6]).IsEqualTo(((byte)217, (byte)217, (byte)217));
        await Assert.That(fills[7]).IsEqualTo(((byte)160, (byte)160, (byte)160));
        await Assert.That(fills.Select(f => (f.R, f.G, f.B)).Distinct().Count()).IsEqualTo(8);
    }

    [Test]
    public async Task T_WIZ_04b_GetFieldBucket_Eight_Buckets()
    {
        var id = new FieldDescriptor { Name = "Id", Label = "Account ID", Type = "id", Nillable = false };
        var nameRequired = new FieldDescriptor
        {
            Name = "Name",
            Label = "Account Name",
            Type = "string",
            Createable = true,
            Updateable = true,
            Nillable = false,
        };
        var requiredStandard = new FieldDescriptor
        {
            Name = "OwnerId",
            Label = "Owner ID",
            Type = "reference",
            Createable = true,
            Updateable = true,
            Nillable = false,
        };
        var requiredCustom = new FieldDescriptor
        {
            Name = "ExternalId__c",
            Label = "External Id",
            Type = "string",
            Createable = true,
            Updateable = true,
            Nillable = false,
            Custom = true,
        };
        var standard = new FieldDescriptor
        {
            Name = "Industry",
            Label = "Industry",
            Type = "picklist",
            Createable = true,
            Updateable = true,
            Nillable = true,
        };
        var custom = new FieldDescriptor
        {
            Name = "Custom__c",
            Label = "Custom Field",
            Type = "string",
            Createable = true,
            Updateable = true,
            Nillable = true,
            Custom = true,
        };
        var readOnlyStandard = new FieldDescriptor
        {
            Name = "CreatedDate",
            Label = "Created Date",
            Type = "datetime",
            Createable = false,
            Updateable = false,
            Nillable = true,
        };
        var readOnlyCustom = new FieldDescriptor
        {
            Name = "RoCustom__c",
            Label = "RO Custom",
            Type = "string",
            Createable = false,
            Updateable = false,
            Nillable = true,
            Custom = true,
        };

        await Assert.That(WizardTableLayoutBuilder.GetFieldBucket(id)).IsEqualTo(0);
        await Assert.That(WizardTableLayoutBuilder.GetFieldBucket(nameRequired)).IsEqualTo(1);
        await Assert.That(WizardTableLayoutBuilder.GetFieldBucket(requiredStandard)).IsEqualTo(2);
        await Assert.That(WizardTableLayoutBuilder.GetFieldBucket(requiredCustom)).IsEqualTo(3);
        await Assert.That(WizardTableLayoutBuilder.GetFieldBucket(standard)).IsEqualTo(4);
        await Assert.That(WizardTableLayoutBuilder.GetFieldBucket(custom)).IsEqualTo(5);
        await Assert.That(WizardTableLayoutBuilder.GetFieldBucket(readOnlyStandard)).IsEqualTo(6);
        await Assert.That(WizardTableLayoutBuilder.GetFieldBucket(readOnlyCustom)).IsEqualTo(7);
    }

    [Test]
    public async Task T_WIZ_04c_NameField_Flag_Classifies_As_Name_Bucket()
    {
        var subject = new FieldDescriptor
        {
            Name = "Subject",
            Label = "Subject",
            Type = "string",
            Createable = true,
            Updateable = true,
            Nillable = false,
            NameField = true,
        };

        await Assert.That(WizardTableLayoutBuilder.IsNameField(subject)).IsTrue();
        await Assert.That(WizardTableLayoutBuilder.GetFieldBucket(subject)).IsEqualTo(1);
    }

    [Test]
    public async Task T_WIZ_04d_FormatFieldListDisplay_Has_No_Brackets()
    {
        var field = new FieldDescriptor
        {
            Name = "Name",
            Label = "Account Name",
            Type = "string",
            Createable = true,
            Updateable = true,
            Nillable = false,
        };

        var display = WizardTableLayoutBuilder.FormatFieldListDisplay(field);
        await Assert.That(display).IsEqualTo("Account Name (Name)");
        await Assert.That(display.Contains('[')).IsFalse();
    }

    [Test]
    public async Task T_WIZ_04e_Parse_NameField_From_Describe_Json()
    {
        const string json = """
            {
              "name": "Case",
              "label": "Case",
              "fields": [
                {
                  "name": "Id",
                  "label": "Case ID",
                  "type": "id",
                  "createable": false,
                  "updateable": false,
                  "nillable": false
                },
                {
                  "name": "Subject",
                  "label": "Subject",
                  "type": "string",
                  "createable": true,
                  "updateable": true,
                  "nillable": false,
                  "nameField": true
                }
              ]
            }
            """;

        var describe = SObjectDescribeJsonParser.Parse(json);
        var subject = describe.Fields.Single(f => f.Name == "Subject");
        await Assert.That(subject.NameField).IsTrue();
        await Assert.That(WizardTableLayoutBuilder.GetFieldBucket(subject)).IsEqualTo(1);
    }

    [Test]
    public async Task T_WIZ_04f_Bucket_Legend_Labels_Match_Eight_Buckets()
    {
        var labels = Enumerable.Range(0, WizardTableLayoutBuilder.BucketCount)
            .Select(WizardTableLayoutBuilder.BucketLegendLabel)
            .ToArray();

        await Assert.That(labels).IsEquivalentTo(
        [
            "Id",
            "Name",
            "Required (standard)",
            "Required (custom)",
            "Standard",
            "Custom",
            "Read-only (standard)",
            "Read-only (custom)",
        ]);
    }

    [Test]
    public async Task T_WIZ_04g_Bucket_Display_Labels_Are_One_Based()
    {
        await Assert.That(WizardTableLayoutBuilder.BucketDisplayLabel(0)).IsEqualTo("1 Id");
        await Assert.That(WizardTableLayoutBuilder.BucketDisplayLabel(1)).IsEqualTo("2 Name");
        await Assert.That(WizardTableLayoutBuilder.BucketDisplayLabel(7)).IsEqualTo("8 Read-only (custom)");
    }

    [Test]
    public async Task T_WIZ_05_Criteria_Row_Values_Are_Contiguous_Triplets()
    {
        var describe = DescribeFixtures.LoadAccountDescribe();
        var id = describe.Fields.First(f => f.IsId);
        var name = describe.Fields.First(f => string.Equals(f.Name, "Name", StringComparison.OrdinalIgnoreCase));
        var criteria = new[]
        {
            new WizardCriteriaClause { Field = id, Operator = "not equals", Value = string.Empty },
            new WizardCriteriaClause { Field = name, Operator = "equals", Value = "Acme" },
        };

        var values = WizardTableLayoutBuilder.BuildCriteriaRowValues(criteria);

        await Assert.That(values.GetLength(0)).IsEqualTo(1);
        await Assert.That(values.GetLength(1)).IsEqualTo(6);
        await Assert.That(values[0, 0]).IsEqualTo(id.Label);
        await Assert.That(values[0, 1]).IsEqualTo("not equals");
        await Assert.That(values[0, 2]).IsEqualTo(string.Empty);
        await Assert.That(values[0, 3]).IsEqualTo(name.Label);
        await Assert.That(values[0, 4]).IsEqualTo("equals");
        await Assert.That(values[0, 5]).IsEqualTo("Acme");
    }
}
