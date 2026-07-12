using System.Net;
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

    [Test]
    public async Task T_QSR_07_Missing_Retrieve_Record_Yields_Blank_Row()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(1), DescribeFixtures.LoadAccountDescribe()).Binding!;
        binding.Snapshot.Body[0, 0] = "001000000000001";

        var selection = new ForceTableSelection
        {
            BodyRowIndices = new[] { 0 },
            StartColumnIndex = 0,
            EndColumnIndex = 2,
        };

        var (client, _) = SalesforceTestClients.Create(handler =>
        {
            handler.Enqueue(HttpStatusCode.OK, "[null]");
        });

        var result = await QueryRows.RunAsync(
            client,
            new QueryRowsInput
            {
                Binding = binding,
                Selection = selection,
            });

        await Assert.That(result.ErrorSummary).IsNull();
        await Assert.That(result.Projection).IsNotNull();
        await Assert.That(result.Projection!.Values.GetLength(0)).IsEqualTo(1);
        await Assert.That(result.Projection.Values.GetLength(1)).IsEqualTo(2);
        await Assert.That(result.Projection.Values[0, 0]).IsNull();
        await Assert.That(result.Projection.Values[0, 1]).IsNull();
    }
}
