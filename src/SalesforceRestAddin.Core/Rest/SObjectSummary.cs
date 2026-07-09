namespace SalesforceRestAddin.Core.Rest;

public sealed class SObjectSummary
{
    public required string Name { get; init; }

    public required string Label { get; init; }

    public bool Queryable { get; init; }

    public bool Custom { get; init; }
}
