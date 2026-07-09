namespace SalesforceRestAddin.Core.Tables;

/// <summary>
/// Resolves ForceConnector table column bounds when multiple tables share a sheet,
/// separated by blank header columns.
/// </summary>
public static class TableBoundaryResolver
{
    /// <summary>
    /// From the active sheet column (1-based), scan the header row left to the start of the
    /// contiguous non-blank header block, then right to its end.
    /// </summary>
    /// <param name="headerRowBySheetColumn">
    /// Sparse or dense map of sheet column (1-based) → header cell value.
    /// Missing / null / whitespace = blank.
    /// </param>
    /// <param name="activeSheetColumn">1-based Excel column of the selection/active cell.</param>
    public static TableColumnBounds Resolve(
        IReadOnlyDictionary<int, object?> headerRowBySheetColumn,
        int activeSheetColumn)
    {
        if (headerRowBySheetColumn is null)
        {
            throw new ArgumentNullException(nameof(headerRowBySheetColumn));
        }

        if (activeSheetColumn < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(activeSheetColumn));
        }

        // If the active column's header is blank, walk left to the nearest non-blank header
        // in this block (user may have selected a body cell under a blank criteria gap — rare).
        var probe = activeSheetColumn;
        if (IsBlank(Get(headerRowBySheetColumn, probe)))
        {
            while (probe > 1 && IsBlank(Get(headerRowBySheetColumn, probe)))
            {
                probe--;
            }

            if (IsBlank(Get(headerRowBySheetColumn, probe)))
            {
                return TableColumnBounds.Invalid(
                    "Could not locate a ForceConnector table header at the selection. " +
                    "Select a cell inside a table (headers on row 2 of the table).");
            }
        }

        var start = probe;
        while (start > 1 && !IsBlank(Get(headerRowBySheetColumn, start - 1)))
        {
            start--;
        }

        var end = start;
        while (!IsBlank(Get(headerRowBySheetColumn, end + 1)))
        {
            end++;
        }

        return new TableColumnBounds
        {
            Succeeded = true,
            StartColumn = start,
            EndColumn = end,
            ColumnCount = end - start + 1,
        };
    }

    /// <summary>
    /// Convenience overload: dense 0-based array aligned to sheet columns starting at
    /// <paramref name="firstSheetColumn"/> (usually 1).
    /// </summary>
    public static TableColumnBounds Resolve(object?[] headerCellsFromFirstSheetColumn, int activeSheetColumn, int firstSheetColumn = 1)
    {
        if (headerCellsFromFirstSheetColumn is null)
        {
            throw new ArgumentNullException(nameof(headerCellsFromFirstSheetColumn));
        }

        var map = new Dictionary<int, object?>();
        for (var i = 0; i < headerCellsFromFirstSheetColumn.Length; i++)
        {
            map[firstSheetColumn + i] = headerCellsFromFirstSheetColumn[i];
        }

        return Resolve(map, activeSheetColumn);
    }

    public static bool IsBlank(object? value) =>
        value is null || string.IsNullOrWhiteSpace(value.ToString());

    private static object? Get(IReadOnlyDictionary<int, object?> map, int sheetColumn) =>
        map.TryGetValue(sheetColumn, out var value) ? value : null;
}

public sealed class TableColumnBounds
{
    public bool Succeeded { get; init; }

    public int StartColumn { get; init; }

    public int EndColumn { get; init; }

    public int ColumnCount { get; init; }

    public string? ErrorMessage { get; init; }

    public static TableColumnBounds Invalid(string message) =>
        new() { Succeeded = false, ErrorMessage = message };
}
