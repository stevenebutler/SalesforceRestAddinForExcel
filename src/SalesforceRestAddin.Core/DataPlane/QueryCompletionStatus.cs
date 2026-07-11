using System;

namespace SalesforceRestAddin.Core.DataPlane;

public static class QueryCompletionStatus
{
    public static string Build(
        int recordsProcessed,
        TimeSpan wallElapsed,
        TimeSpan apiElapsed,
        TimeSpan worksheetElapsed)
    {
        var overlapElapsed = apiElapsed + worksheetElapsed - wallElapsed;
        return
            $"Query completed: {recordsProcessed} records in {FormatElapsed(wallElapsed)} " +
            $"(API {FormatElapsed(apiElapsed)}, worksheet {FormatElapsed(worksheetElapsed)}, " +
            $"{FormatElapsed(overlapElapsed)} overlapped)";
    }

    private static string FormatElapsed(TimeSpan elapsed)
    {
        if (elapsed < TimeSpan.Zero)
        {
            elapsed = TimeSpan.Zero;
        }

        return $"{elapsed.TotalSeconds:0.0}s";
    }
}
