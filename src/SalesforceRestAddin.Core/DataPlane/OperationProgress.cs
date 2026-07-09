namespace SalesforceRestAddin.Core.DataPlane;

/// <summary>Progress update for long-running data-plane operations (query pagination, etc.).</summary>
public readonly record struct OperationProgress(string Message, int? Completed = null, int? Total = null)
{
    public static OperationProgress Status(string message) => new(message);

    public static OperationProgress Download(int completed, int total) =>
        new($"Downloaded {completed} / {total} records", completed, total);

    public static OperationProgress DownloadUnknownTotal(int completed) =>
        new($"Downloaded {completed} records", completed, null);
}
