using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Tests.TableLimits;

public sealed class SelectionLimitsTests
{
    [Test]
    public async Task T_LIM_01_3501_Rows_Rejected_By_Default()
    {
        var selection = new ForceTableSelection
        {
            BodyRowIndices = Enumerable.Range(0, 3501).ToArray(),
            StartColumnIndex = 0,
            EndColumnIndex = 5,
            IsMultiArea = false,
        };

        var error = SelectionLimits.ValidateSelection(selection, ConnectorOptions.Default);

        await Assert.That(error).IsNotNull();
        await Assert.That(error).Contains("3500");
    }

    [Test]
    public async Task T_LIM_02_21_Columns_On_Update_Rejected()
    {
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0 },
            StartColumnIndex = 0,
            EndColumnIndex = 20,
            IsMultiArea = false,
        };

        var error = SelectionLimits.ValidateSelection(selection, ConnectorOptions.Default);

        await Assert.That(error).IsNotNull();
        await Assert.That(error).Contains("20");
    }

    [Test]
    public async Task T_LIM_03_NoQueryLimit_Allows_3501_Rows()
    {
        var selection = new ForceTableSelection
        {
            BodyRowIndices = Enumerable.Range(0, 3501).ToArray(),
            StartColumnIndex = 0,
            EndColumnIndex = 5,
            IsMultiArea = false,
        };

        var error = SelectionLimits.ValidateSelection(
            selection,
            new ConnectorOptions { NoQueryLimit = true });

        await Assert.That(error).IsNull();
    }

    [Test]
    public async Task T_LIM_04_MultiArea_Selection_Rejected_By_Default()
    {
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0, 1 },
            StartColumnIndex = 0,
            EndColumnIndex = 3,
            IsMultiArea = true,
        };

        var error = SelectionLimits.ValidateSelection(selection, ConnectorOptions.Default);

        await Assert.That(error).IsNotNull();
        await Assert.That(error).Contains("Multiple");
    }

    [Test]
    public async Task T_LIM_05_MultiArea_Allowed_When_Opted_In()
    {
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0, 1 },
            StartColumnIndex = 0,
            EndColumnIndex = 3,
            IsMultiArea = true,
        };

        var error = SelectionLimits.ValidateSelection(
            selection,
            ConnectorOptions.Default,
            allowMultiArea: true);

        await Assert.That(error).IsNull();
    }

    [Test]
    public async Task T_LIM_06_Distinct_Column_Count_For_Sparse_MultiArea()
    {
        // Bounding box 0..3 would be 4 columns; distinct selected cols are 1 and 3 → 2.
        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0, 1 },
            StartColumnIndex = 1,
            EndColumnIndex = 3,
            IsMultiArea = true,
            ColumnsByBodyRow = new Dictionary<int, IReadOnlyList<int>>
            {
                [0] = new[] { 1 },
                [1] = new[] { 3 },
            },
        };

        await Assert.That(selection.ColumnCount).IsEqualTo(2);

        var error = SelectionLimits.ValidateSelection(
            selection,
            ConnectorOptions.Default,
            allowMultiArea: true);

        await Assert.That(error).IsNull();
    }
}
