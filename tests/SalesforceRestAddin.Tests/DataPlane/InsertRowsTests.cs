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
}
