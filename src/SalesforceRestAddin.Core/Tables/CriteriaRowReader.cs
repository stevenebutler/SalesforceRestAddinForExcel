namespace SalesforceRestAddin.Core.Tables;

internal sealed class CriteriaRowReadResult
{
    public object?[] CriteriaRow { get; init; } = Array.Empty<object?>();

    public BindingValidationError? Error { get; init; }

    public bool Succeeded => Error is null;
}

internal static class CriteriaRowReader
{
    private const int TripletsPerChunk = 4;
    private const int CellsPerTriplet = 3;
    private const int ChunkWidth = TripletsPerChunk * CellsPerTriplet + 1;

    public static CriteriaRowReadResult Read(
        int criteriaRow,
        int startColumn,
        Func<int, int, object?[]> readColumns,
        FieldCatalog catalog)
    {
        if (readColumns is null)
        {
            throw new ArgumentNullException(nameof(readColumns));
        }

        if (catalog is null)
        {
            throw new ArgumentNullException(nameof(catalog));
        }

        if (startColumn < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(startColumn));
        }

        if (criteriaRow < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(criteriaRow));
        }

        var values = new List<object?>();
        var currentStartColumn = startColumn;

        while (true)
        {
            var slice = NormalizeSlice(readColumns(currentStartColumn, ChunkWidth), ChunkWidth);
            if (slice.Length == 0 || IsBlank(slice[0]))
            {
                return new CriteriaRowReadResult { CriteriaRow = values.ToArray() };
            }

            for (var triplet = 0; triplet < TripletsPerChunk; triplet++)
            {
                var baseIndex = triplet * CellsPerTriplet;
                var fieldCell = slice[baseIndex];
                if (IsBlank(fieldCell))
                {
                    return new CriteriaRowReadResult { CriteriaRow = values.ToArray() };
                }

                var fieldLabel = fieldCell?.ToString()?.Trim();
                if (string.IsNullOrWhiteSpace(fieldLabel))
                {
                    return new CriteriaRowReadResult { CriteriaRow = values.ToArray() };
                }

                if (catalog.ResolveHeaderField(fieldLabel, null) is null)
                {
                    return new CriteriaRowReadResult
                    {
                        Error = new BindingValidationError
                        {
                            Cell = new CellRef(criteriaRow, currentStartColumn + baseIndex),
                            Message = $"Unknown criteria field '{fieldLabel}' at R{criteriaRow}C{currentStartColumn + baseIndex}.",
                        },
                    };
                }

                values.Add(slice[baseIndex]);
                values.Add(slice[baseIndex + 1]);
                values.Add(slice[baseIndex + 2]);
            }

            if (IsBlank(slice[ChunkWidth - 1]))
            {
                return new CriteriaRowReadResult { CriteriaRow = values.ToArray() };
            }

            currentStartColumn += TripletsPerChunk * CellsPerTriplet;
        }
    }

    private static object?[] NormalizeSlice(object?[] slice, int width)
    {
        if (slice.Length == width)
        {
            return slice;
        }

        var normalized = new object?[width];
        Array.Copy(slice, normalized, Math.Min(slice.Length, width));
        return normalized;
    }

    private static bool IsBlank(object? value) =>
        value is null || string.IsNullOrWhiteSpace(value.ToString());
}
