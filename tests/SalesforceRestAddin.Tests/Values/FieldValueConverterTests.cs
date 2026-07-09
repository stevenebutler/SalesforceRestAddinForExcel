using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Core.Values;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.Values;

public sealed class FieldValueConverterTests
{
    private static FieldDescriptor NameField => DescribeFixtures.LoadAccountDescribe().Fields.First(f => f.Name == "Name");

    [Test]
    public async Task T_VAL_01_Empty_Cell_Is_Null_On_Update_And_Create()
    {
        var updateEmpty = FieldValueConverter.ToSalesforceValue(NameField, "", ConnectorOptions.Default, forCreate: false);
        await Assert.That(updateEmpty).IsNull();

        var createEmpty = FieldValueConverter.ToSalesforceValue(NameField, null, ConnectorOptions.Default, forCreate: true);
        await Assert.That(createEmpty).IsNull();

        var dateField = new FieldDescriptor
        {
            Name = "PO_date__c",
            Label = "PO Date",
            Type = "date",
            Createable = true,
            Updateable = true,
        };
        var clearedDate = FieldValueConverter.ToSalesforceValue(dateField, "", ConnectorOptions.Default, forCreate: false);
        await Assert.That(clearedDate).IsNull();
    }

    [Test]
    public async Task T_VAL_02_Date_Cell_Iso_Format()
    {
        var field = new FieldDescriptor { Name = "CloseDate", Label = "Close Date", Type = "date", Createable = true, Updateable = true };
        var value = FieldValueConverter.ToSalesforceValue(field, new DateTime(2024, 3, 15), ConnectorOptions.Default, forCreate: true);
        await Assert.That(value).IsEqualTo("2024-03-15");
    }

    [Test]
    public async Task T_VAL_06_Picklist_Passthrough()
    {
        var field = DescribeFixtures.LoadAccountDescribe().Fields.First(f => f.Name == "Industry");
        var value = FieldValueConverter.ToSalesforceValue(field, "Technology", ConnectorOptions.Default, forCreate: true);
        await Assert.That(value).IsEqualTo("Technology");
    }

    [Test]
    public async Task T_VAL_10_Non_Updateable_Omitted()
    {
        var idField = DescribeFixtures.LoadAccountDescribe().Fields.First(f => f.Name == "Id");
        var value = FieldValueConverter.ToSalesforceValue(idField, "001", ConnectorOptions.Default, forCreate: false);
        await Assert.That(value).IsNull();
    }

    [Test]
    public async Task T_VAL_04_Reference_Name_To_Id()
    {
        var field = DescribeFixtures.LoadAccountDescribe().Fields.First(f => f.Name == "ParentId");
        var resolver = new DictionaryReferenceResolver(new Dictionary<(string, string), string>
        {
            [("Account", "Parent")] = "001XX",
        });
        var display = FieldValueConverter.ToDisplayValue(
            field,
            "001XX",
            new ConnectorOptions { UseReference = true },
            resolver,
            "Parent");
        await Assert.That(display).IsEqualTo("Parent");
    }

    [Test]
    public async Task CreateColumnFormat_Date_Uses_Legacy_Excel_Format()
    {
        var date = FieldValueConverter.CreateColumnFormat(new FieldDescriptor
        {
            Name = "CreatedDate",
            Label = "Created Date",
            Type = "datetime",
        });
        await Assert.That(date.Kind).IsEqualTo(ColumnFormatKind.DateTime);
        await Assert.That(date.ExcelFormat).IsEqualTo("yyyy-MM-dd HH:mm:ss");

        var day = FieldValueConverter.CreateColumnFormat(new FieldDescriptor
        {
            Name = "CloseDate",
            Label = "Close Date",
            Type = "date",
        });
        await Assert.That(day.Kind).IsEqualTo(ColumnFormatKind.Date);
        await Assert.That(day.ExcelFormat).IsEqualTo("yyyy-MM-dd");
    }

    [Test]
    public async Task CreateColumnFormat_Text_And_Currency()
    {
        var text = FieldValueConverter.CreateColumnFormat(new FieldDescriptor
        {
            Name = "Name",
            Label = "Name",
            Type = "string",
        });
        await Assert.That(text.ExcelFormat).IsEqualTo("@");

        var currency = FieldValueConverter.CreateColumnFormat(new FieldDescriptor
        {
            Name = "Amount",
            Label = "Amount",
            Type = "currency",
        });
        await Assert.That(currency.Kind).IsEqualTo(ColumnFormatKind.Number);
        await Assert.That(currency.ExcelFormat).IsEqualTo("$#,##0_);($#,##0)");
    }
}
