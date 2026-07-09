namespace SalesforceRestAddin.Core.Tables;

public sealed class FieldDescriptor
{
    public required string Name { get; init; }

    public required string Label { get; init; }

    public required string Type { get; init; }

    public bool Createable { get; init; }

    public bool Updateable { get; init; }

    public bool Nillable { get; init; }

    public bool Custom { get; init; }

    /// <summary>Salesforce describe <c>nameField</c> — the object's display-name field.</summary>
    public bool NameField { get; init; }

    public int? Length { get; init; }

    public int? Precision { get; init; }

    public int? Scale { get; init; }

    public IReadOnlyList<string> ReferenceTo { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> PicklistValues { get; init; } = Array.Empty<string>();

    public bool IsId => string.Equals(Name, "Id", StringComparison.OrdinalIgnoreCase);

    public bool IsReference => string.Equals(Type, "reference", StringComparison.OrdinalIgnoreCase);
}
