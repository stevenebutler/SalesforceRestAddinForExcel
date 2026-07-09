namespace SalesforceRestAddin.Core.Tables;

/// <summary>Selected body rows and column span within a ForceConnector table.</summary>
public sealed class ForceTableSelection
{
    /// <summary>0-based body row indices (row 3 = index 0).</summary>
    public required IReadOnlyList<int> BodyRowIndices { get; init; }

    /// <summary>0-based inclusive column range within the table (bounding box).</summary>
    public required int StartColumnIndex { get; init; }

    public required int EndColumnIndex { get; init; }

    public bool IsMultiArea { get; init; }

    /// <summary>
    /// Body row → selected table-relative column indices (per-row update fields).
    /// When set, <see cref="ColumnCount"/> is the distinct column count across all rows.
    /// </summary>
    public IReadOnlyDictionary<int, IReadOnlyList<int>>? ColumnsByBodyRow { get; init; }

    /// <summary>
    /// Column count for selection limits: distinct selected columns when
    /// <see cref="ColumnsByBodyRow"/> is set; otherwise the bounding-box width.
    /// </summary>
    public int ColumnCount
    {
        get
        {
            if (ColumnsByBodyRow is { Count: > 0 } map)
            {
                return map.Values.SelectMany(cols => cols).Distinct().Count();
            }

            return EndColumnIndex - StartColumnIndex + 1;
        }
    }

    public int RowCount => BodyRowIndices.Count;

    /// <summary>
    /// Columns selected for a body row. Uses <see cref="ColumnsByBodyRow"/> when present;
    /// otherwise the inclusive <see cref="StartColumnIndex"/>…<see cref="EndColumnIndex"/> span.
    /// </summary>
    public IReadOnlyList<int> ColumnsForBodyRow(int bodyRowIndex)
    {
        if (ColumnsByBodyRow is not null
            && ColumnsByBodyRow.TryGetValue(bodyRowIndex, out var columns))
        {
            return columns;
        }

        if (EndColumnIndex < StartColumnIndex)
        {
            return Array.Empty<int>();
        }

        var count = EndColumnIndex - StartColumnIndex + 1;
        var span = new int[count];
        for (var i = 0; i < count; i++)
        {
            span[i] = StartColumnIndex + i;
        }

        return span;
    }
}
