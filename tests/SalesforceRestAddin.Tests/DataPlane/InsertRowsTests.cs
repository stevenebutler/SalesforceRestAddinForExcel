using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class InsertRowsTests
{
    [Test]
    public async Task T_ISR_01_New_Rows_Case_Insensitive()
    {
        await Assert.That(InsertRows.IsNewRow("new")).IsTrue();
        await Assert.That(InsertRows.IsNewRow("NEW")).IsTrue();
        await Assert.That(InsertRows.IsNewRow("New")).IsTrue();
        await Assert.That(InsertRows.IsNewRow("Acme")).IsFalse();
    }

    [Test]
    public async Task T_ISR_04_No_Eligible_Rows_Error()
    {
        var binding = ForceTableBinder.Bind(ForceTableSnapshots.ValidAccountTable(), DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection { BodyRowIndices = new[] { 0 }, StartColumnIndex = 0, EndColumnIndex = 3 };
        var (client, _) = SalesforceTestClients.Create();
        var result = await InsertRows.RunAsync(client, new InsertRowsInput { Binding = binding, Selection = selection });
        await Assert.That(result.ErrorSummary).IsEqualTo("No rows marked New for insert.");
    }

    [Test]
    public async Task T_ISR_03b_Create_Failure_Sets_ErrorSummary_For_Dialog()
    {
        var snapshot = ForceTableSnapshots.ValidAccountTable(1);
        snapshot.Body[0, 0] = "New";
        var binding = ForceTableBinder.Bind(snapshot, DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection { BodyRowIndices = new[] { 0 }, StartColumnIndex = 0, EndColumnIndex = 2 };
        var (client, _) = SalesforceTestClients.Create(h =>
        {
            h.Enqueue(
                System.Net.HttpStatusCode.OK,
                """[{"id":null,"success":false,"errors":[{"statusCode":"REQUIRED_FIELD_MISSING","message":"Required fields are missing: [Name]","fields":["Name"]}]}]""");
        });

        var result = await InsertRows.RunAsync(client, new InsertRowsInput { Binding = binding, Selection = selection });

        await Assert.That(result.RowOutcomes.Count).IsEqualTo(1);
        await Assert.That(result.RowOutcomes[0].Succeeded).IsFalse();
        await Assert.That(result.ErrorSummary).IsNotNull();
        await Assert.That(result.ErrorSummary!).Contains("Insert failed for 1 of 1 row(s)");
        await Assert.That(result.ErrorSummary!).Contains("Row 3:");
        await Assert.That(result.ErrorSummary!).Contains("Required fields are missing: [Name]");
    }

    [Test]
    public async Task T_ISR_03_Success_Writes_Id_Without_ErrorSummary()
    {
        var snapshot = ForceTableSnapshots.ValidAccountTable(1);
        snapshot.Body[0, 0] = "New";
        var binding = ForceTableBinder.Bind(snapshot, DescribeFixtures.LoadAccountDescribe()).Binding!;
        var selection = new ForceTableSelection { BodyRowIndices = new[] { 0 }, StartColumnIndex = 0, EndColumnIndex = 2 };
        var (client, _) = SalesforceTestClients.Create(h =>
        {
            h.Enqueue(
                System.Net.HttpStatusCode.OK,
                """[{"id":"001000000000ABC","success":true,"errors":[]}]""");
        });

        var result = await InsertRows.RunAsync(client, new InsertRowsInput { Binding = binding, Selection = selection });

        await Assert.That(result.RowOutcomes[0].Succeeded).IsTrue();
        await Assert.That(result.RowOutcomes[0].IdWriteback).IsEqualTo("001000000000ABC");
        await Assert.That(result.ErrorSummary).IsNull();
    }
}
