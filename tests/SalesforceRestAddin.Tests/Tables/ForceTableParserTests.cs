using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.TableParsing;

public sealed class ForceTableParserTests
{
    private static SObjectDescribe AccountDescribe => DescribeFixtures.LoadAccountDescribe();

    [Test]
    public async Task T_TBL_01_Valid_Snapshot_Binding_Succeeds_IdColumn_Correct()
    {
        var snapshot = ForceTableSnapshots.ValidAccountTable();
        var result = ForceTableBinder.Bind(snapshot, AccountDescribe);

        await Assert.That(result.Succeeded).IsTrue();
        await Assert.That(result.Binding!.IdColumnIndex).IsEqualTo(0);
        await Assert.That(result.Binding!.Columns.Count).IsEqualTo(3);
    }

    [Test]
    public async Task T_TBL_02_A1_Value_When_No_Comment_Resolved()
    {
        var snapshot = ForceTableSnapshots.AccountFromA1Value();
        var result = ForceTableBinder.Bind(snapshot, AccountDescribe);

        await Assert.That(result.Succeeded).IsTrue();
        await Assert.That(result.Binding!.Snapshot.ObjectApiName).IsEqualTo("Account");
    }

    [Test]
    public async Task T_TBL_03_A1_Empty_Or_Space_Validation_Error_Cites_A1()
    {
        var snapshot = new ForceTableSnapshot
        {
            ObjectApiName = "Bad Name",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = new object?[] { "Account" },
            HeaderApiNames = new string?[] { null },
            Body = new object?[0, 1],
        };

        var result = ForceTableBinder.Bind(snapshot, AccountDescribe);

        await Assert.That(result.Succeeded).IsFalse();
        await Assert.That(result.Errors[0].Cell.Row).IsEqualTo(1);
        await Assert.That(result.Errors[0].Cell.Column).IsEqualTo(1);
    }

    [Test]
    public async Task T_TBL_04_Unknown_Label_Stops_At_Gap_Error_Identifies_Column()
    {
        var snapshot = new ForceTableSnapshot
        {
            ObjectApiName = "Account",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = new object?[] { "Account ID", "Account Name", "Unknown Field", "Industry" },
            HeaderApiNames = new string?[] { "Id", "Name", null, "Industry" },
            Body = new object?[0, 4],
        };

        var result = ForceTableBinder.Bind(snapshot, AccountDescribe);

        await Assert.That(result.Succeeded).IsFalse();
        await Assert.That(result.Errors[0].Cell.Column).IsEqualTo(3);
    }

    [Test]
    public async Task T_TBL_05_Row2_Comment_Api_Name_Maps_Field()
    {
        var snapshot = ForceTableSnapshots.WithCustomFieldComment();
        var result = ForceTableBinder.Bind(snapshot, AccountDescribe);

        await Assert.That(result.Succeeded).IsTrue();
        await Assert.That(result.Binding!.Columns[1].Field.Name).IsEqualTo("Custom__c");
    }

    [Test]
    public async Task T_TBL_05b_Record_Id_Header_Without_Comment_Maps_To_Id()
    {
        var snapshot = new ForceTableSnapshot
        {
            ObjectApiName = "Account",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = new object?[] { "Record Id", "Account Name" },
            HeaderApiNames = new string?[] { null, "Name" },
            Body = new object?[0, 2],
        };

        var result = ForceTableBinder.Bind(snapshot, AccountDescribe);

        await Assert.That(result.Succeeded).IsTrue();
        await Assert.That(result.Binding!.Columns[0].Field.Name).IsEqualTo("Id");
        await Assert.That(result.Binding.IdColumnIndex).IsEqualTo(0);
    }

    [Test]
    public async Task T_TBL_06_No_Id_Column_Validation_Error()
    {
        var snapshot = ForceTableSnapshots.MissingIdColumn();
        var result = ForceTableBinder.Bind(snapshot, AccountDescribe);

        await Assert.That(result.Succeeded).IsFalse();
        await Assert.That(result.Errors.Any(e => e.Message.Contains("Id", StringComparison.OrdinalIgnoreCase))).IsTrue();
    }

    [Test]
    public async Task T_TBL_07_Hidden_Row_Excluded_From_Visible_Selection()
    {
        var snapshot = new ForceTableSnapshot
        {
            ObjectApiName = "Account",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = ForceTableSnapshots.ValidAccountTable().HeaderLabels,
            HeaderApiNames = ForceTableSnapshots.ValidAccountTable().HeaderApiNames,
            Body = ForceTableSnapshots.ValidAccountTable().Body,
            IdColumnIndex = 0,
            HiddenRowIndices = new[] { 1 },
        };

        var visible = HiddenSelectionFilter.VisibleBodyRows(snapshot, new[] { 0, 1, 2 }).ToList();

        await Assert.That(visible).IsEquivalentTo(new[] { 0, 2 });
    }

    [Test]
    public async Task T_TBL_08_Hidden_Column_Excluded_From_Visible_Selection()
    {
        var snapshot = new ForceTableSnapshot
        {
            ObjectApiName = "Account",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = ForceTableSnapshots.ValidAccountTable().HeaderLabels,
            HeaderApiNames = ForceTableSnapshots.ValidAccountTable().HeaderApiNames,
            Body = ForceTableSnapshots.ValidAccountTable().Body,
            IdColumnIndex = 0,
            HiddenColumnIndices = new[] { 2 },
        };

        var visible = HiddenSelectionFilter.VisibleColumns(snapshot, 0, 2).ToList();

        await Assert.That(visible).IsEquivalentTo(new[] { 0, 1 });
    }

    [Test]
    public async Task T_TBL_09_Id_Not_First_Column_Binds_Successfully()
    {
        var snapshot = ForceTableSnapshots.IdNotFirstColumn();
        var result = ForceTableBinder.Bind(snapshot, AccountDescribe);

        await Assert.That(result.Succeeded).IsTrue();
        await Assert.That(result.Binding!.IdColumnIndex).IsEqualTo(1);
        await Assert.That(result.Binding.Columns[1].Field.IsId).IsTrue();
    }

    [Test]
    public async Task T_TBL_13_Bind_Preserves_StartRow_And_StartColumn()
    {
        var snapshot = ForceTableSnapshots.ValidAccountTableAt(startRow: 10, startColumn: 5);
        var result = ForceTableBinder.Bind(snapshot, AccountDescribe);

        await Assert.That(result.Succeeded).IsTrue();
        await Assert.That(result.Binding!.Snapshot.StartRow).IsEqualTo(10);
        await Assert.That(result.Binding.Snapshot.StartColumn).IsEqualTo(5);
    }

    [Test]
    public async Task T_TBL_14_Missing_Object_Cites_Anchor_Address()
    {
        var snapshot = new ForceTableSnapshot
        {
            ObjectApiName = "",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = new object?[] { "Account ID" },
            HeaderApiNames = new string?[] { "Id" },
            Body = new object?[0, 1],
            StartRow = 3,
            StartColumn = 5,
        };

        var result = ForceTableBinder.Bind(snapshot, AccountDescribe);

        await Assert.That(result.Succeeded).IsFalse();
        await Assert.That(result.Errors[0].Message).Contains("E3");
    }
}
