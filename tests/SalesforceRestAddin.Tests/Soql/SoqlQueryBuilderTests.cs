using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Soql;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.Soql;

public sealed class SoqlQueryBuilderTests
{
    private static FieldCatalog Catalog => new(DescribeFixtures.LoadAccountDescribe());

    [Test]
    public async Task T_SOQL_01_Equals_Operator()
    {
        var result = await ParseAsync("Account Name", "equals", "Acme");
        await Assert.That(result.WhereClause).IsEqualTo("Name = 'Acme'");
    }

    [Test]
    public async Task T_SOQL_02_Contains_Operator()
    {
        var result = await ParseAsync("Account Name", "contains", "Acme");
        await Assert.That(result.WhereClause).IsEqualTo("Name like '%Acme%'");
    }

    [Test]
    public async Task T_SOQL_03_Not_Equals_Operator()
    {
        var result = await ParseAsync("Account Name", "not equals", "Acme");
        await Assert.That(result.WhereClause).IsEqualTo("Name != 'Acme'");
    }

    [Test]
    public async Task T_SOQL_04_Begins_With_Operator()
    {
        var result = await ParseAsync("Account Name", "begins with", "Ac");
        await Assert.That(result.WhereClause).IsEqualTo("Name like 'Ac%'");
    }

    [Test]
    public async Task T_SOQL_05_Ends_With_Operator()
    {
        var result = await ParseAsync("Account Name", "ends with", "me");
        await Assert.That(result.WhereClause).IsEqualTo("Name like '%me'");
    }

    [Test]
    public async Task T_SOQL_06_Comma_Separated_Values_Or_Group()
    {
        var result = await ParseAsync("Account Name", "equals", "a,b");
        await Assert.That(result.WhereClause).IsEqualTo("(Name = 'a' or Name = 'b')");
    }

    [Test]
    public async Task T_SOQL_07_Empty_Text_Value()
    {
        var result = await ParseAsync("Account Name", "equals", "");
        await Assert.That(result.WhereClause).IsEqualTo("Name = ''");
    }

    [Test]
    public async Task T_SOQL_11_Soql_Special_Chars_Escaped()
    {
        var result = await ParseAsync("Account Name", "equals", "O'Brien");
        await Assert.That(result.WhereClause).IsEqualTo("Name = 'O\\'Brien'");
    }

    [Test]
    public async Task T_SOQL_12_Reference_UseReference_Resolves_Id()
    {
        var resolver = new DictionaryReferenceResolver(new Dictionary<(string, string), string>
        {
            [("Account", "Parent Co")] = "001XX000003DGbQ",
        });
        var criteria = new object?[] { "Parent Account", "equals", "Parent Co" };
        var result = await SoqlCriteriaParser.ParseAsync(
            criteria,
            Catalog,
            new ConnectorOptions { UseReference = true },
            resolver);

        await Assert.That(result.WhereClause).IsEqualTo("ParentId = '001XX000003DGbQ'");
    }

    [Test]
    public async Task T_SOQL_12b_Legacy_Record_Id_Label_Resolves_Id()
    {
        var result = await ParseAsync("Record Id", "not equals", "");

        await Assert.That(result.WhereClause).IsEqualTo("Id != ''");
    }

    [Test]
    public async Task T_SOQL_13_Select_Field_Order_Matches_Binding()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var soql = SoqlQueryBuilder.BuildSelectQuery(binding, "Id != null");
        await Assert.That(soql).IsEqualTo("SELECT Id, Name, Industry FROM Account WHERE Id != null");
    }

    [Test]
    public async Task T_SOQL_14_Reference_Range_Is_Batched()
    {
        var criteria = new object?[] { "Parent Account", "in", "AccountIds" };
        var result = await SoqlCriteriaParser.ParseAsync(
            criteria,
            Catalog,
            ConnectorOptions.Default,
            null,
            criteriaReferenceIds: new Dictionary<int, IReadOnlyList<string>>
            {
                [2] = ["001000000000001", "001000000000002"],
            });

        await Assert.That(result.WhereClause).IsEqualTo(string.Empty);
        await Assert.That(result.ReferenceJoinField).IsEqualTo("ParentId");
        await Assert.That(result.ReferenceJoinIds).IsEquivalentTo(new[] { "001000000000001", "001000000000002" });
    }

    [Test]
    public async Task T_SOQL_15_Reference_In_Requires_Excel_Range()
    {
        var result = await ParseAsync("Parent Account", "in", "001000000000001,001000000000002");

        await Assert.That(result.Succeeded).IsFalse();
        await Assert.That(result.Errors[0].Message).Contains("Excel range or named range");
    }

    [Test]
    public async Task T_SOQL_15a_On_Is_Rejected_With_In_Guidance()
    {
        var result = await ParseAsync("Parent Account", "on", "CustomerIds");

        await Assert.That(result.Succeeded).IsFalse();
        await Assert.That(result.Errors[0].Message).IsEqualTo(
            "ON is not supported - use IN with a range reference to select multiple items.");
    }

    [Test]
    public async Task T_SOQL_15b_Multipicklist_Excludes_Is_Preserved()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "Regions__c",
            Label = "Regions",
            Type = "multipicklist",
        });

        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "Regions", "excludes", "APAC" },
            catalog,
            ConnectorOptions.Default,
            null);

        await Assert.That(result.WhereClause).IsEqualTo("Regions__c excludes ('APAC')");
    }

    [Test]
    public async Task T_SOQL_15c_Reference_Range_Is_Split_Into_Batches()
    {
        var batches = SoqlQueryBuilder.BuildReferenceInBatches(
            ["001000000000001", "001000000000002", "001000000000003"],
            "ParentId",
            batchSize: 2);

        await Assert.That(batches).IsEquivalentTo(new[]
        {
            "ParentId IN ('001000000000001', '001000000000002')",
            "ParentId IN ('001000000000003')",
        });
    }

    [Test]
    public async Task T_SOQL_16_Count_Query_Same_Where()
    {
        var count = SoqlQueryBuilder.BuildCountQuery("Account", "Name = 'Acme'");
        await Assert.That(count).IsEqualTo("SELECT COUNT(Id) FROM Account WHERE Name = 'Acme'");
    }

    [Test]
    public async Task T_SOQL_08_Empty_Date_Uses_Null()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "CloseDate",
            Label = "Close Date",
            Type = "date",
        });
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "Close Date", "equals", "" },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo("CloseDate = null");
    }

    [Test]
    public async Task Double_Criteria_Is_Unquoted_With_Decimal()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "Batch_Import_Number__c",
            Label = "Batch Import Number",
            Type = "double",
        });
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "Batch Import Number", "equals", "0" },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo("Batch_Import_Number__c = 0.0");
    }

    [Test]
    public async Task Currency_Criteria_Preserves_Decimal()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "Amount",
            Label = "Amount",
            Type = "currency",
        });
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "Amount", "equals", "12.5" },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo("Amount = 12.5");
    }

    [Test]
    public async Task Boolean_Criteria_Is_Unquoted_TrueFalse()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "IsActive",
            Label = "Active",
            Type = "boolean",
        });
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "Active", "equals", "true" },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo("IsActive = TRUE");
    }

    [Test]
    public async Task Int_Criteria_Is_Unquoted()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "NumberOfEmployees",
            Label = "Employees",
            Type = "int",
        });
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "Employees", "equals", "10" },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo("NumberOfEmployees = 10");
    }

    [Test]
    public async Task Date_Criteria_Is_Unquoted_Iso()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "CloseDate",
            Label = "Close Date",
            Type = "date",
        });
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "Close Date", "equals", "2024-03-15" },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo("CloseDate = 2024-03-15");
    }

    [Test]
    public async Task Date_Criteria_From_Excel_DateTime_Cell()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "PO_date__c",
            Label = "PO Date",
            Type = "date",
        });
        var excelDate = new DateTime(2021, 7, 26);
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "PO Date", "greater than", excelDate },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo("PO_date__c > 2021-07-26");
    }

    [Test]
    public async Task Date_Criteria_From_Excel_OaDate_Serial()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "PO_date__c",
            Label = "PO Date",
            Type = "date",
        });
        var expected = new DateTime(2021, 7, 26);
        var serial = expected.ToOADate();
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "PO Date", "greater than", serial },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo("PO_date__c > 2021-07-26");
    }

    [Test]
    public async Task DateTime_Criteria_From_Excel_OaDate_Serial_With_Fraction()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "CreatedDate",
            Label = "Created Date",
            Type = "datetime",
        });
        var serial = new DateTime(2021, 7, 26, 14, 30, 0, DateTimeKind.Local).ToOADate();
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "Created Date", "equals", serial },
            catalog,
            ConnectorOptions.Default,
            null);

        await Assert.That(result.WhereClause!.StartsWith("CreatedDate = ")).IsTrue();
        await Assert.That(result.WhereClause).Contains("2021-07-26T");
        await Assert.That(result.WhereClause).Contains("Z");
        await Assert.That(result.WhereClause).DoesNotContain("'");
    }

    [Test]
    public async Task Date_Criteria_From_OaDate_Stringified_Serial()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "PO_date__c",
            Label = "PO Date",
            Type = "date",
        });
        // Same serial the failing Salesforce query logged (Excel Value2.ToString()).
        var serialText = new DateTime(2021, 7, 26).ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "PO Date", "greater than", serialText },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo("PO_date__c > 2021-07-26");
    }

    [Test]
    public async Task Date_Criteria_From_Logged_OaDate_Serial_44368()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "PO_date__c",
            Label = "PO Date",
            Type = "date",
        });
        var expected = DateTime.FromOADate(44368);
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "PO Date", "greater than", 44368d },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo(
            $"PO_date__c > {expected:yyyy-MM-dd}");
    }

    [Test]
    public async Task Date_Criteria_Iso_With_Space_Separator()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "CreatedDate",
            Label = "Created Date",
            Type = "datetime",
        });
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "Created Date", "equals", "2021-07-26 14:30:00" },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.Succeeded).IsTrue();
        await Assert.That(result.WhereClause).Contains("CreatedDate = 2021-07-26T");
        await Assert.That(result.WhereClause).Contains("Z");
    }

    [Test]
    public async Task Date_Criteria_Locale_String_Is_Rejected()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "PO_date__c",
            Label = "PO Date",
            Type = "date",
        });
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "PO Date", "greater than", "26/07/2021" },
            catalog,
            ConnectorOptions.Default,
            null);

        await Assert.That(result.Succeeded).IsFalse();
        await Assert.That(result.Errors[0].Message).Contains("Invalid date value");
        await Assert.That(result.Errors[0].Message).Contains("26/07/2021");
        await Assert.That(result.Errors[0].Message).Contains("Excel date cell");
    }

    [Test]
    public async Task Date_Criteria_Invalid_String_Aborts_With_Error()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "PO_date__c",
            Label = "PO Date",
            Type = "date",
        });
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "PO Date", "equals", "not-a-date" },
            catalog,
            ConnectorOptions.Default,
            null);

        await Assert.That(result.Succeeded).IsFalse();
        await Assert.That(result.Errors[0].Message).Contains("Invalid date value");
        await Assert.That(result.Errors[0].Message).Contains("not-a-date");
    }

    [Test]
    public async Task Date_Criteria_Relative_Today_Passthrough()
    {
        var catalog = CatalogWith(new FieldDescriptor
        {
            Name = "PO_date__c",
            Label = "PO Date",
            Type = "date",
        });
        var result = await SoqlCriteriaParser.ParseAsync(
            new object?[] { "PO Date", "equals", "TODAY" },
            catalog,
            ConnectorOptions.Default,
            null);
        await Assert.That(result.WhereClause).IsEqualTo("PO_date__c = TODAY");
    }

    private static FieldCatalog CatalogWith(params FieldDescriptor[] fields) =>
        new(new SObjectDescribe
        {
            Name = "Account",
            Label = "Account",
            Fields = fields,
        });

    private static Task<SoqlCriteriaParseResult> ParseAsync(string field, string op, string value)
    {
        var criteria = new object?[] { field, op, value };
        return SoqlCriteriaParser.ParseAsync(criteria, Catalog, ConnectorOptions.Default, null);
    }
}
