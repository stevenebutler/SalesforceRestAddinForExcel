using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Core.DataPlane;

public sealed class WizardCriteriaClause
{
    public required FieldDescriptor Field { get; init; }

    public required string Operator { get; init; }

    public string? Value { get; init; }
}
