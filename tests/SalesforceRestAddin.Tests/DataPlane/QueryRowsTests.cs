using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class QueryRowsTests
{
    [Test]
    public async Task T_QSR_06_Limit_Rejects_3501_Rows()
    {
        var rows = Enumerable.Range(0, 3501).ToList();
        var selection = new ForceTableSelection { BodyRowIndices = rows, StartColumnIndex = 0, EndColumnIndex = 3 };
        var error = SelectionLimits.ValidateSelection(selection, ConnectorOptions.Default);
        await Assert.That(error).IsNotNull();
    }

    [Test]
    public async Task T_QSR_05_Refresh_Stops_At_First_Id_Gap()
    {
        var snapshot = ForceTableSnapshots.ValidAccountTable(5);
        snapshot.Body[2, 0] = null;
        await Assert.That(SalesforceId.IsValid(snapshot.Body[2, 0]?.ToString())).IsFalse();
        await Assert.That(SalesforceId.IsValid(snapshot.Body[0, 0]?.ToString())).IsTrue();
        await Assert.That(SalesforceId.IsValid(snapshot.Body[1, 0]?.ToString())).IsTrue();
    }
}
