using SalesforceRestAddin.Core.DataPlane;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class QueryCompletionStatusTests
{
    [Test]
    public async Task Build_Uses_Wall_Time_And_Reports_Overlap()
    {
        var message = QueryCompletionStatus.Build(
            recordsProcessed: 42,
            wallElapsed: TimeSpan.FromMilliseconds(1200),
            apiElapsed: TimeSpan.FromMilliseconds(1000),
            worksheetElapsed: TimeSpan.FromMilliseconds(900));

        await Assert.That(message).IsEqualTo(
            "Query completed: 42 records in 1.2s (API 1.0s, worksheet 0.9s, 0.7s overlapped)");
    }

    [Test]
    public async Task Build_Clamps_Negative_Overlap_To_Zero()
    {
        var message = QueryCompletionStatus.Build(
            recordsProcessed: 42,
            wallElapsed: TimeSpan.FromSeconds(3),
            apiElapsed: TimeSpan.FromSeconds(1),
            worksheetElapsed: TimeSpan.FromSeconds(1));

        await Assert.That(message).IsEqualTo(
            "Query completed: 42 records in 3.0s (API 1.0s, worksheet 1.0s, 0.0s overlapped)");
    }
}
