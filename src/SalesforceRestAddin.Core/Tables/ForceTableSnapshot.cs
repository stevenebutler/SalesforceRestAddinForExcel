namespace SalesforceRestAddin.Core.Tables;

/// <summary>
/// Bulk-read ForceConnector table layout from Excel — no Excel types.
/// Row 1: criteria; row 2: headers; row 3+: body.
/// </summary>
public sealed class ForceTableSnapshot
{
    public required string ObjectApiName { get; init; }

    /// <summary>Row 1 values from column B onward (criteria triplets).</summary>
    public required object?[] CriteriaRow { get; init; }

    /// <summary>
    /// Salesforce Ids expanded by the Excel host from range/name values in hidden legacy
    /// reference criteria. The key is the zero-based criteria-row value-cell index.
    /// </summary>
    public IReadOnlyDictionary<int, IReadOnlyList<string>> CriteriaReferenceIds { get; init; } =
        new Dictionary<int, IReadOnlyList<string>>();

    /// <summary>Row 2 header labels (column A through last data column).</summary>
    public required object?[] HeaderLabels { get; init; }

    /// <summary>API names from row-2 cell comments, parallel to <see cref="HeaderLabels"/>.</summary>
    public required string?[] HeaderApiNames { get; init; }

    /// <summary>Body rows (row 3+), 0-based first dimension = body row index.</summary>
    public required object?[,] Body { get; init; }

    public int IdColumnIndex { get; init; } = -1;

    public IReadOnlyList<int>? HiddenRowIndices { get; init; }

    public IReadOnlyList<int>? HiddenColumnIndices { get; init; }

    public int ColumnCount => HeaderLabels.Length;

    public int BodyRowCount => Body.GetLength(0);

    /// <summary>Excel row of the table anchor (row 1 / object cell).</summary>
    public int StartRow { get; init; } = 1;

    /// <summary>Excel column of the table anchor (column A of the table).</summary>
    public int StartColumn { get; init; } = 1;
}
