namespace SalesforceRestAddin.Core.Rest;

public sealed class QueryResultPage
{
    public required int TotalSize { get; init; }

    public required bool Done { get; init; }

    public string? NextRecordsUrl { get; init; }

    public required IReadOnlyList<Dictionary<string, object?>> Records { get; init; }
}

public sealed class SaveResult
{
    public required bool Success { get; init; }

    public string? Id { get; init; }

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();
}

public sealed class DeleteResult
{
    public required bool Success { get; init; }

    public string? Id { get; init; }

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();
}
