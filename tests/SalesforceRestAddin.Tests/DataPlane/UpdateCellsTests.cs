using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class UpdateCellsTests
{
    [Test]
    public async Task T_USC_02_No_Updateable_Columns_Error()
    {
        var snapshot = ForceTableSnapshots.ValidAccountTable();
        var binding = ForceTableBinder.Bind(snapshot, DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0 },
            StartColumnIndex = 0,
            EndColumnIndex = 0,
        };

        var result = UpdateCells.Plan(binding, selection, ConnectorOptions.Default);
        await Assert.That(result.ErrorSummary).IsEqualTo("No updatable columns selected.");
    }

    [Test]
    public async Task T_USC_01_Patch_Contains_Id_And_Updateable_Fields()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0 },
            StartColumnIndex = 0,
            EndColumnIndex = 2,
        };

        var result = UpdateCells.Plan(binding, selection, ConnectorOptions.Default);
        await Assert.That(result.ErrorSummary).IsNull();
        await Assert.That(result.RecordsProcessed).IsEqualTo(1);
    }

    [Test]
    public async Task BuildRecords_Includes_Attributes_Type_For_Composite_Patch()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0 },
            StartColumnIndex = 0,
            EndColumnIndex = 2,
        };

        var records = UpdateCells.BuildRecords(binding, selection, ConnectorOptions.Default, out var hasUpdateable);

        await Assert.That(hasUpdateable).IsTrue();
        await Assert.That(records.Count).IsEqualTo(1);
        var attributes = records[0]["attributes"] as Dictionary<string, object?>;
        await Assert.That(attributes).IsNotNull();
        await Assert.That(attributes!["type"]).IsEqualTo("Account");
        await Assert.That(records[0]["Id"]).IsNotNull();
    }

    [Test]
    public async Task BuildRecords_Cleared_Cell_Sends_Null()
    {
        var snapshot = ForceTableSnapshots.ValidAccountTable(1);
        snapshot.Body[0, 1] = null; // clear Name
        var binding = ForceTableBinder.Bind(snapshot, DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0 },
            StartColumnIndex = 0,
            EndColumnIndex = 2,
        };

        var records = UpdateCells.BuildRecords(binding, selection, ConnectorOptions.Default, out _);

        await Assert.That(records.Count).IsEqualTo(1);
        await Assert.That(records[0].ContainsKey("Name")).IsTrue();
        await Assert.That(records[0]["Name"]).IsNull();
    }

    [Test]
    public async Task T_USC_10_MultiArea_Per_Row_Fields_Only()
    {
        var binding = ForceTableBinder.Bind(
            ForceTableSnapshots.ValidAccountTable(2),
            DescribeFixtures.LoadAccountDescribe()).Binding!;

        // Row 0: Name only (col 1); row 1: Industry only (col 2).
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0, 1 },
            StartColumnIndex = 1,
            EndColumnIndex = 2,
            IsMultiArea = true,
            ColumnsByBodyRow = new Dictionary<int, IReadOnlyList<int>>
            {
                [0] = new[] { 1 },
                [1] = new[] { 2 },
            },
        };

        var plan = UpdateCells.Plan(binding, selection, ConnectorOptions.Default);
        await Assert.That(plan.ErrorSummary).IsNull();

        var records = UpdateCells.BuildRecords(binding, selection, ConnectorOptions.Default, out _);
        await Assert.That(records.Count).IsEqualTo(2);

        await Assert.That(records[0].ContainsKey("Name")).IsTrue();
        await Assert.That(records[0].ContainsKey("Industry")).IsFalse();
        await Assert.That(records[0]["Name"]).IsEqualTo("Acme 0");

        await Assert.That(records[1].ContainsKey("Industry")).IsTrue();
        await Assert.That(records[1].ContainsKey("Name")).IsFalse();
        await Assert.That(records[1]["Industry"]).IsEqualTo("Technology");
    }

    [Test]
    public async Task T_USC_11_Overlapping_Areas_Union_Columns_On_Same_Row()
    {
        var binding = ForceTableBinder.Bind(
            ForceTableSnapshots.ValidAccountTable(1),
            DescribeFixtures.LoadAccountDescribe()).Binding!;

        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0 },
            StartColumnIndex = 1,
            EndColumnIndex = 2,
            IsMultiArea = true,
            ColumnsByBodyRow = new Dictionary<int, IReadOnlyList<int>>
            {
                [0] = new[] { 1, 2 },
            },
        };

        var records = UpdateCells.BuildRecords(binding, selection, ConnectorOptions.Default, out _);
        await Assert.That(records.Count).IsEqualTo(1);
        await Assert.That(records[0].ContainsKey("Name")).IsTrue();
        await Assert.That(records[0].ContainsKey("Industry")).IsTrue();
    }

    [Test]
    public async Task T_USC_12_Column_Outside_Table_Aborted()
    {
        var binding = ForceTableBinder.Bind(
            ForceTableSnapshots.ValidAccountTable(1),
            DescribeFixtures.LoadAccountDescribe()).Binding!;

        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0 },
            StartColumnIndex = 1,
            EndColumnIndex = 5,
            IsMultiArea = true,
            ColumnsByBodyRow = new Dictionary<int, IReadOnlyList<int>>
            {
                [0] = new[] { 1, 5 },
            },
        };

        var plan = UpdateCells.Plan(binding, selection, ConnectorOptions.Default);
        await Assert.That(plan.ErrorSummary).IsEqualTo(UpdateCells.OutsideTableMessage);
    }

    [Test]
    public async Task T_USC_06_Default_Omits_Hidden_Row_From_Patch()
    {
        var baseSnapshot = ForceTableSnapshots.ValidAccountTable(3);
        var snapshot = new ForceTableSnapshot
        {
            ObjectApiName = baseSnapshot.ObjectApiName,
            CriteriaRow = baseSnapshot.CriteriaRow,
            HeaderLabels = baseSnapshot.HeaderLabels,
            HeaderApiNames = baseSnapshot.HeaderApiNames,
            Body = baseSnapshot.Body,
            IdColumnIndex = baseSnapshot.IdColumnIndex,
            HiddenRowIndices = new[] { 1 },
        };
        var binding = ForceTableBinder.Bind(snapshot, DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0, 1, 2 },
            StartColumnIndex = 0,
            EndColumnIndex = 2,
        };

        var records = UpdateCells.BuildRecords(binding, selection, ConnectorOptions.Default, out _);

        await Assert.That(records.Count).IsEqualTo(2);
        await Assert.That(records[0]["Id"]?.ToString()).Contains("000");
        await Assert.That(records[1]["Id"]?.ToString()).Contains("002");
    }

    [Test]
    public async Task T_USC_06b_IncludeHiddenCells_Sends_Hidden_Row()
    {
        var baseSnapshot = ForceTableSnapshots.ValidAccountTable(3);
        var snapshot = new ForceTableSnapshot
        {
            ObjectApiName = baseSnapshot.ObjectApiName,
            CriteriaRow = baseSnapshot.CriteriaRow,
            HeaderLabels = baseSnapshot.HeaderLabels,
            HeaderApiNames = baseSnapshot.HeaderApiNames,
            Body = baseSnapshot.Body,
            IdColumnIndex = baseSnapshot.IdColumnIndex,
            HiddenRowIndices = new[] { 1 },
        };
        var binding = ForceTableBinder.Bind(snapshot, DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0, 1, 2 },
            StartColumnIndex = 0,
            EndColumnIndex = 2,
        };

        var records = UpdateCells.BuildRecords(
            binding,
            selection,
            new ConnectorOptions { IncludeHiddenCells = true },
            out _);

        await Assert.That(records.Count).IsEqualTo(3);
    }
}
