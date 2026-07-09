namespace SalesforceRestAddin.Core.Tables;

public static class ForceTableParser
{
    /// <summary>Validates object API name from the table anchor cell (comment or value).</summary>
    public static BindingValidationResult ValidateObjectName(string? objectApiName, int startRow = 1, int startColumn = 1)
    {
        if (string.IsNullOrWhiteSpace(objectApiName) || objectApiName.Contains(' '))
        {
            var address = $"{GetColumnLetter(startColumn)}{startRow}";
            return new BindingValidationResult
            {
                Errors =
                [
                    new BindingValidationError
                    {
                        Cell = new CellRef(startRow, startColumn),
                        Message =
                            $"Could not locate an object name in cell {address}. " +
                            "The entity name must appear above the first (Id) column of the table.",
                    },
                ],
            };
        }

        return new BindingValidationResult();
    }

    /// <summary>Finds Id column index (0-based) from header API names or labels.</summary>
    public static int FindIdColumnIndex(ForceTableSnapshot snapshot)
    {
        for (var col = 0; col < snapshot.HeaderLabels.Length; col++)
        {
            var apiName = snapshot.HeaderApiNames[col];
            if (string.Equals(apiName, "Id", StringComparison.OrdinalIgnoreCase))
            {
                return col;
            }

            var label = snapshot.HeaderLabels[col]?.ToString();
            if (string.Equals(label, "Account ID", StringComparison.OrdinalIgnoreCase)
                || string.Equals(label, "Id", StringComparison.OrdinalIgnoreCase))
            {
                // Label-only hint; binding confirms via describe.
            }
        }

        return snapshot.IdColumnIndex;
    }

    private static string GetColumnLetter(int columnNumber)
    {
        var dividend = columnNumber;
        var columnName = string.Empty;
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }
}
