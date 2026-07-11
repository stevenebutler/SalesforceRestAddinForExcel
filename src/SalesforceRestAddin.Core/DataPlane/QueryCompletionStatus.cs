using System;

namespace SalesforceRestAddin.Core.DataPlane;

public static class QueryCompletionStatus
{
    public static string Build(int recordsProcessed, TimeSpan queryElapsed, TimeSpan worksheetElapsed)
    {
        var totalElapsed = queryElapsed + worksheetElapsed;
        return
            $"Query completed: {recordsProcessed} records in {FormatElapsed(totalElapsed)} " +
            $"(query {FormatElapsed(queryElapsed)}, worksheet {FormatElapsed(worksheetElapsed)})";
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
