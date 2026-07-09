using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class DeleteRecordsTests
{
    [Test]
    public async Task T_DR_02_Success_Marks_Deleted()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection { BodyRowIndices = new[] { 0 }, StartColumnIndex = 0, EndColumnIndex = 3 };
        var (client, handler) = SalesforceTestClients.Create(h =>
        {
            h.Enqueue(System.Net.HttpStatusCode.OK, """[{"id":"001000000000000","success":true,"errors":[]}]""");
        });

        var result = await DeleteRecords.RunAsync(client, new DeleteRecordsInput
        {
            Binding = binding,
            Selection = selection,
        });

        await Assert.That(result.RowOutcomes[0].IdDisplayOverride).IsEqualTo("deleted");
        await Assert.That(result.ErrorSummary).IsNull();
        await Assert.That(handler.Requests[0].RequestUri!.Query).Contains("001000000000000");
    }

    [Test]
    public async Task T_DR_03_Failure_Sets_ErrorSummary_For_Dialog()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection { BodyRowIndices = new[] { 0 }, StartColumnIndex = 0, EndColumnIndex = 3 };
        var (client, _) = SalesforceTestClients.Create(h =>
        {
            h.Enqueue(
                System.Net.HttpStatusCode.OK,
                """[{"id":"001000000000000","success":false,"errors":[{"statusCode":"ENTITY_IS_DELETED","message":"entity is deleted","fields":[]}]}]""");
        });

        var result = await DeleteRecords.RunAsync(client, new DeleteRecordsInput
        {
            Binding = binding,
            Selection = selection,
        });

        await Assert.That(result.RowOutcomes[0].Succeeded).IsFalse();
        await Assert.That(result.ErrorSummary).IsNotNull();
        await Assert.That(result.ErrorSummary!).Contains("Delete failed for 1 of 1 row(s)");
        await Assert.That(result.ErrorSummary!).Contains("Row 3:");
        await Assert.That(result.ErrorSummary!).Contains("Delete Row Failed");
        await Assert.That(result.ErrorSummary!).Contains("entity is deleted");
    }
}
