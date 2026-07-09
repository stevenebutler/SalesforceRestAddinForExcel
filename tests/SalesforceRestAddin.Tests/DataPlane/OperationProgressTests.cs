using SalesforceRestAddin.Core.DataPlane;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class OperationProgressTests
{
    [Test]
    public async Task Download_Formats_Status_And_Counts()
    {
        var progress = OperationProgress.Download(400, 1020);

        await Assert.That(progress.Message).IsEqualTo("Downloaded 400 / 1020 records");
        await Assert.That(progress.Completed).IsEqualTo(400);
        await Assert.That(progress.Total).IsEqualTo(1020);
    }

    [Test]
    public async Task Status_Has_No_Counts()
    {
        var progress = OperationProgress.Status("Querying Salesforce...");

        await Assert.That(progress.Message).IsEqualTo("Querying Salesforce...");
        await Assert.That(progress.Completed).IsNull();
        await Assert.That(progress.Total).IsNull();
    }
}
