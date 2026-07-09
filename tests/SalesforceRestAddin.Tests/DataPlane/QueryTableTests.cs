using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class QueryTableTests
{
    [Test]
    public async Task T_QTD_03_Zero_Rows_Marks_NF()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var result = QueryTable.ProjectRecords(binding, Array.Empty<Dictionary<string, object?>>(), new QueryTableInput
        {
            Snapshot = binding.Snapshot,
        });

        await Assert.That(result.Projection!.Values[0, 0]).IsEqualTo("#N/F");
    }

    [Test]
    public async Task T_QTD_07_Projector_Output_Is_Rectangular()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var records = new List<Dictionary<string, object?>>
        {
            new() { ["Id"] = "001A", ["Name"] = "A", ["Industry"] = "Tech" },
        };
        var result = QueryTable.ProjectRecords(binding, records, new QueryTableInput { Snapshot = binding.Snapshot });
        await Assert.That(result.Projection!.Values.GetLength(0)).IsEqualTo(1);
        await Assert.That(result.Projection.Values.GetLength(1)).IsEqualTo(3);
    }

    [Test]
    public async Task T_QTD_08_Plan_Includes_Clear_Body_Instruction()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(5), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var records = new List<Dictionary<string, object?>> { new() { ["Id"] = "001", ["Name"] = "A", ["Industry"] = "T" } };
        var result = QueryTable.ProjectRecords(binding, records, new QueryTableInput { Snapshot = binding.Snapshot });
        await Assert.That(result.ClearBody).IsNotNull();
        await Assert.That(result.ClearBody!.StartRow).IsEqualTo(binding.Snapshot.StartRow + 2);
        // Smaller result must still clear the full previous body (bugs.md #5 / FR-QTD-2).
        await Assert.That(result.ClearBody.EndRow).IsEqualTo(binding.Snapshot.StartRow + 2 + 5 - 1);
        await Assert.That(result.ClearBody.StartColumn).IsEqualTo(binding.Snapshot.StartColumn);
        await Assert.That(result.ClearBody.EndColumn)
            .IsEqualTo(binding.Snapshot.StartColumn + binding.Snapshot.ColumnCount - 1);
    }

    [Test]
    public async Task T_QTD_08b_Zero_Rows_Clears_Full_Existing_Body()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(4), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var result = QueryTable.ProjectRecords(
            binding,
            Array.Empty<Dictionary<string, object?>>(),
            new QueryTableInput { Snapshot = binding.Snapshot });

        await Assert.That(result.ClearBody).IsNotNull();
        await Assert.That(result.ClearBody!.StartRow).IsEqualTo(3);
        await Assert.That(result.ClearBody.EndRow).IsEqualTo(6);
        await Assert.That(result.Projection!.Values[0, 0]).IsEqualTo("#N/F");
    }

    [Test]
    public async Task T_QTD_08c_Larger_Result_Clear_Covers_Projection()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(1), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var records = new List<Dictionary<string, object?>>
        {
            new() { ["Id"] = "001", ["Name"] = "A", ["Industry"] = "T" },
            new() { ["Id"] = "002", ["Name"] = "B", ["Industry"] = "T" },
            new() { ["Id"] = "003", ["Name"] = "C", ["Industry"] = "T" },
        };
        var result = QueryTable.ProjectRecords(binding, records, new QueryTableInput { Snapshot = binding.Snapshot });

        await Assert.That(result.ClearBody!.EndRow).IsEqualTo(binding.Snapshot.StartRow + 2 + 3 - 1);
    }

    [Test]
    public async Task T_PRJ_02_Column_Formats_Match_Column_Count_And_Types()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var records = new List<Dictionary<string, object?>>
        {
            new() { ["Id"] = "001A", ["Name"] = "A", ["Industry"] = "Tech" },
        };
        var result = QueryTable.ProjectRecords(binding, records, new QueryTableInput { Snapshot = binding.Snapshot });
        var formats = result.Projection!.ColumnFormats!;

        await Assert.That(formats.Count).IsEqualTo(result.Projection.Values.GetLength(1));
        await Assert.That(formats[0].Kind).IsEqualTo(ColumnFormatKind.Text); // Id
        await Assert.That(formats[0].ExcelFormat).IsEqualTo("@");
        await Assert.That(formats[1].ExcelFormat).IsEqualTo("@"); // Name string
    }

    [Test]
    public async Task T_PRJ_Offset_Table_Projection_Uses_StartColumn()
    {
        var binding = ForceTableBinder.Bind(
            ForceTableSnapshots.ValidAccountTableAt(startRow: 8, startColumn: 5),
            DescribeFixtures.LoadAccountDescribe()).Binding!;
        var records = new List<Dictionary<string, object?>>
        {
            new() { ["Id"] = "001A", ["Name"] = "A", ["Industry"] = "Tech" },
        };
        var result = QueryTable.ProjectRecords(binding, records, new QueryTableInput { Snapshot = binding.Snapshot });

        await Assert.That(result.Projection!.StartRow).IsEqualTo(10);
        await Assert.That(result.Projection.StartColumn).IsEqualTo(5);
        await Assert.That(result.ClearBody!.StartColumn).IsEqualTo(5);
        await Assert.That(result.ClearBody.EndColumn).IsEqualTo(7);
    }

    [Test]
    public async Task ProjectRecords_DateTime_Display_And_Format()
    {
        var describe = new SObjectDescribe
        {
            Name = "Account",
            Label = "Account",
            Fields = new[]
            {
                new FieldDescriptor { Name = "Id", Label = "Account ID", Type = "id", Createable = false, Updateable = false },
                new FieldDescriptor { Name = "CreatedDate", Label = "Created Date", Type = "datetime", Createable = false, Updateable = false },
            },
        };
        var snapshot = new ForceTableSnapshot
        {
            ObjectApiName = "Account",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = new object?[] { "Account ID", "Created Date" },
            HeaderApiNames = new string?[] { "Id", "CreatedDate" },
            Body = new object?[0, 2],
            IdColumnIndex = 0,
        };
        var binding = ForceTableBinder.Bind(snapshot, describe).Binding!;
        var result = QueryTable.ProjectRecords(
            binding,
            new[] { new Dictionary<string, object?> { ["Id"] = "001A", ["CreatedDate"] = "2007-11-05T17:45:00.000+0000" } },
            new QueryTableInput { Snapshot = snapshot });

        await Assert.That(result.Projection!.ColumnFormats!.Count).IsEqualTo(2);
        await Assert.That(result.Projection.ColumnFormats[1].Kind).IsEqualTo(ColumnFormatKind.DateTime);
        await Assert.That(result.Projection.ColumnFormats[1].ExcelFormat).IsEqualTo("yyyy-MM-dd HH:mm:ss");
        await Assert.That(result.Projection.Values[0, 1]).IsAssignableTo<DateTime>();
    }
}
