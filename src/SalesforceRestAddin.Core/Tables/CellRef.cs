namespace SalesforceRestAddin.Core.Tables;

/// <summary>1-based worksheet row/column for user-facing validation messages.</summary>
public readonly record struct CellRef(int Row, int Column)
{
    public override string ToString() => $"R{Row}C{Column}";
}
