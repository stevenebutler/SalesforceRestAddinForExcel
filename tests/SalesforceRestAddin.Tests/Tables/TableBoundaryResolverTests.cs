using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Tests.Tables;

public sealed class TableBoundaryResolverTests
{
    [Test]
    public async Task T_TBL_10_Scan_Left_Finds_Block_Start_After_Gap()
    {
        // Columns: 1 blank, 2-4 left table, 5 blank, 6-8 right table. Active in col 8.
        var headers = new Dictionary<int, object?>
        {
            [2] = "Account ID",
            [3] = "Name",
            [4] = "Industry",
            [6] = "Contact ID",
            [7] = "Last Name",
            [8] = "Email",
        };

        var bounds = TableBoundaryResolver.Resolve(headers, activeSheetColumn: 8);

        await Assert.That(bounds.Succeeded).IsTrue();
        await Assert.That(bounds.StartColumn).IsEqualTo(6);
        await Assert.That(bounds.EndColumn).IsEqualTo(8);
        await Assert.That(bounds.ColumnCount).IsEqualTo(3);
    }

    [Test]
    public async Task T_TBL_11_Left_Table_From_Active_In_Left_Block()
    {
        var headers = new Dictionary<int, object?>
        {
            [1] = "Account ID",
            [2] = "Name",
            [3] = "Industry",
            [5] = "Contact ID",
            [6] = "Email",
        };

        var bounds = TableBoundaryResolver.Resolve(headers, activeSheetColumn: 2);

        await Assert.That(bounds.StartColumn).IsEqualTo(1);
        await Assert.That(bounds.EndColumn).IsEqualTo(3);
    }

    [Test]
    public async Task T_TBL_12_No_Header_Anywhere_To_The_Left_Fails()
    {
        var headers = new Dictionary<int, object?>();

        var bounds = TableBoundaryResolver.Resolve(headers, activeSheetColumn: 10);

        await Assert.That(bounds.Succeeded).IsFalse();
        await Assert.That(bounds.ErrorMessage).IsNotNull();
    }

    [Test]
    public async Task Blank_Active_Column_Scans_Left_To_Nearest_Block()
    {
        var headers = new Dictionary<int, object?>
        {
            [2] = "Account ID",
            [3] = "Name",
        };

        // Active in empty space to the right — walk left to the nearest header block.
        var bounds = TableBoundaryResolver.Resolve(headers, activeSheetColumn: 10);

        await Assert.That(bounds.Succeeded).IsTrue();
        await Assert.That(bounds.StartColumn).IsEqualTo(2);
        await Assert.That(bounds.EndColumn).IsEqualTo(3);
    }

    [Test]
    public async Task Dense_Array_Overload_Uses_First_Sheet_Column()
    {
        // Sheet cols 5-7: Id, Name, Industry
        var cells = new object?[] { "Account ID", "Name", "Industry" };
        var bounds = TableBoundaryResolver.Resolve(cells, activeSheetColumn: 6, firstSheetColumn: 5);

        await Assert.That(bounds.StartColumn).IsEqualTo(5);
        await Assert.That(bounds.EndColumn).IsEqualTo(7);
    }
}
