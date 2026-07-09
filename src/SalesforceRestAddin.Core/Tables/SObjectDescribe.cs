namespace SalesforceRestAddin.Core.Tables;

public sealed class SObjectDescribe
{
    public required string Name { get; init; }

    public required string Label { get; init; }

    public required IReadOnlyList<FieldDescriptor> Fields { get; init; }
}
