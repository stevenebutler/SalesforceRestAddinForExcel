namespace SalesforceRestAddin.Core.Soql;

public sealed class SoqlCriteriaError
{
    public required Tables.CellRef Cell { get; init; }

    public required string Message { get; init; }
}

public sealed class SoqlCriteriaParseResult
{
    public string? WhereClause { get; init; }

    public IReadOnlyList<SoqlCriteriaError> Errors { get; init; } = Array.Empty<SoqlCriteriaError>();

    public IReadOnlyList<string>? ReferenceJoinIds { get; init; }

    public string? ReferenceJoinField { get; init; }

    public bool IsReferenceJoinMode { get; init; }

    public bool Succeeded => Errors.Count == 0;
}
