namespace SalesforceRestAddin.Core.Tables;

public sealed class SheetProjection
{
    public required object?[,] Values { get; init; }

    /// <summary>1-based worksheet row of <see cref="Values"/>[0,0].</summary>
    public int StartRow { get; init; }

    /// <summary>1-based worksheet column of <see cref="Values"/>[0,0].</summary>
    public int StartColumn { get; init; }

    public IReadOnlyList<ColumnFormat>? ColumnFormats { get; init; }

    public IReadOnlyList<RowOutcome>? RowOutcomes { get; init; }
}
