namespace SalesforceRestAddin.Core.Rest;

public sealed class SObjectSummary
{
    public required string Name { get; init; }

    public required string Label { get; init; }

    public bool Queryable { get; init; }

    public bool Custom { get; init; }

    public bool CustomSetting { get; init; }

    public bool DeprecatedAndHidden { get; init; }

    public string? AssociateEntityType { get; init; }

    public string? AssociateParentEntity { get; init; }

    public string? KeyPrefix { get; init; }
}
