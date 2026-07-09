using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Core.Tables;

public static class SelectionLimits
{
    public const int DefaultMaxRows = 3500;
    public const int DefaultMaxColumns = 20;

    public static string? ValidateSelection(
        ForceTableSelection selection,
        ConnectorOptions options,
        bool allowMultiArea = false)
    {
        if (selection is null)
        {
            throw new ArgumentNullException(nameof(selection));
        }

        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        if (selection.IsMultiArea && !allowMultiArea)
        {
            return "Multiple selection areas are not supported.";
        }

        if (options.NoQueryLimit)
        {
            return null;
        }

        if (selection.RowCount > DefaultMaxRows)
        {
            return $"Selection exceeds the maximum of {DefaultMaxRows} rows.";
        }

        if (selection.ColumnCount > DefaultMaxColumns)
        {
            return $"Selection exceeds the maximum of {DefaultMaxColumns} columns.";
        }

        return null;
    }

    public static string? ValidateRowCount(int rowCount, ConnectorOptions options, bool allowNoQueryLimit = true)
    {
        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        if (allowNoQueryLimit && options.NoQueryLimit)
        {
            return null;
        }

        if (rowCount > DefaultMaxRows)
        {
            return $"Selection exceeds the maximum of {DefaultMaxRows} rows.";
        }

        return null;
    }
}
