namespace SalesforceRestAddin.Core.Tables;

/// <summary>Excel instruction to clear body rows before a query write.</summary>
public sealed class ClearBodyRegion
{
    /// <summary>1-based first body row (typically 3).</summary>
    public int StartRow { get; init; }

    /// <summary>1-based first column to clear (typically 1).</summary>
    public int StartColumn { get; init; } = 1;

    /// <summary>1-based last column to clear.</summary>
    public int EndColumn { get; init; }

    /// <summary>1-based last row to clear (may extend beyond current body).</summary>
    public int EndRow { get; init; }
}
