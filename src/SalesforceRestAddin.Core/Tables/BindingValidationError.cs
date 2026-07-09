namespace SalesforceRestAddin.Core.Tables;

public sealed class BindingValidationError
{
    public required CellRef Cell { get; init; }

    public required string Message { get; init; }
}

public sealed class BindingValidationResult
{
    public ForceTableBinding? Binding { get; init; }

    public IReadOnlyList<BindingValidationError> Errors { get; init; } = Array.Empty<BindingValidationError>();

    public bool Succeeded => Binding is not null && Errors.Count == 0;
}
