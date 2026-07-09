namespace SalesforceRestAddin.Core.Tables;

public sealed class DataOperationResult
{
    public SheetProjection? Projection { get; init; }

    public ClearBodyRegion? ClearBody { get; init; }

    public IReadOnlyList<RowOutcome> RowOutcomes { get; init; } = Array.Empty<RowOutcome>();

    public int RecordsProcessed { get; init; }

    public bool WasCancelled { get; init; }

    public string? ErrorSummary { get; init; }

    public bool Succeeded => string.IsNullOrEmpty(ErrorSummary) && !WasCancelled;
}
