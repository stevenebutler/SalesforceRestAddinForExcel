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
        await Assert.That(handler.Requests[0].RequestUri!.Query).Contains("001000000000000");
    }
}
