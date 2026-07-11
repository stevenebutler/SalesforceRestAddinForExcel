using SalesforceRestAddin.Core.DataPlane;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class QueryCompletionStatusTests
{
    [Test]
    public async Task Build_Includes_Query_And_Worksheet_Timing()
    {
        var message = QueryCompletionStatus.Build(
            recordsProcessed: 42,
            queryElapsed: TimeSpan.FromMilliseconds(1200),
            worksheetElapsed: TimeSpan.FromMilliseconds(300));

        await Assert.That(message).IsEqualTo(
            "Query completed: 42 records in 1.5s (query 1.2s, worksheet 0.3s)");
    }
}
