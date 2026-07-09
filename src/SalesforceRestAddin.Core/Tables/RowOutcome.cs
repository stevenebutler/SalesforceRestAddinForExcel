namespace SalesforceRestAddin.Core.Tables;

public sealed class RowOutcome
{
    /// <summary>0-based body row index (row 3 = 0).</summary>
    public int BodyRowIndex { get; init; }

    public bool Succeeded { get; init; }

    public IReadOnlyList<string> ErrorMessages { get; init; } = Array.Empty<string>();

    /// <summary>New Salesforce Id after successful insert.</summary>
    public string? IdWriteback { get; init; }

    /// <summary>Display override for Id cell (e.g. "deleted").</summary>
    public string? IdDisplayOverride { get; init; }
}
