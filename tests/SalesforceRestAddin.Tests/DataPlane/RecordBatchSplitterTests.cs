using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class RecordBatchSplitterTests
{
    [Test]
    public async Task T_BAT_01_450_Ids_Batch_200_Chunks_Correctly()
    {
        var ids = Enumerable.Range(0, 450).Select(i => $"id{i}").ToList();

        var batches = RecordBatchSplitter.Split(ids, 200);

        await Assert.That(batches.Count).IsEqualTo(3);
        await Assert.That(batches[0].Count).IsEqualTo(200);
        await Assert.That(batches[1].Count).IsEqualTo(200);
        await Assert.That(batches[2].Count).IsEqualTo(50);
    }

    [Test]
    public async Task T_BAT_02_Zero_Ids_Empty_Chunk_List()
    {
        var batches = RecordBatchSplitter.Split(Array.Empty<string>(), 200);

        await Assert.That(batches).IsEmpty();
    }

    [Test]
    public async Task T_BAT_03_Custom_CompositeBatchSize_Honored()
    {
        var ids = Enumerable.Range(0, 120).Select(i => $"id{i}").ToList();
        var options = new ConnectorOptions { CompositeBatchSize = 50 };

        var batches = RecordBatchSplitter.Split(ids, options.CompositeBatchSize);

        await Assert.That(batches.Count).IsEqualTo(3);
        await Assert.That(batches[0].Count).IsEqualTo(50);
        await Assert.That(batches[2].Count).IsEqualTo(20);
    }
}
