using System;
using System.Drawing;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using Microsoft.Office.Interop.Excel;

namespace SalesforceRestAddin.Tables;

public static class SheetProjectionWriter
{
    /// <summary>Legacy cap: long text must not expand a row beyond ~3× the sheet standard height.</summary>
    public const double MaxRowHeightMultiplier = 3d;

    public static void Apply(
        Worksheet worksheet,
        SheetProjection projection,
        ColumnSizingMode columnSizingMode,
        RowSizingMode rowSizingMode)
    {
        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        if (projection is null)
        {
            throw new ArgumentNullException(nameof(projection));
        }

        var rows = projection.Values.GetLength(0);
        var cols = projection.Values.GetLength(1);
        if (rows == 0 || cols == 0)
        {
            return;
        }

        var target = WriteValues(worksheet, projection);
        ApplyColumnFormats(target, projection.ColumnFormats);
        ApplyColumnSizing(worksheet, target, projection, rows, cols, columnSizingMode, preserveExistingColumnWidths: false);
        ApplyRowSizing(worksheet, target, rowSizingMode);
        ApplyRowOutcomes(worksheet, projection);
    }

    /// <summary>Writes and sizes one page of a streamed query.</summary>
    public static void ApplyPage(
        Worksheet worksheet,
        SheetProjection projection,
        bool applyColumnFormats,
        bool autoFitColumns,
        RowSizingMode rowSizingMode,
        bool preserveExistingColumnWidths)
    {
        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        if (projection is null)
        {
            throw new ArgumentNullException(nameof(projection));
        }

        if (projection.Values.GetLength(0) == 0 || projection.Values.GetLength(1) == 0)
        {
            return;
        }

        var target = WriteValues(worksheet, projection);
        if (applyColumnFormats)
        {
            ApplyColumnFormats(target, projection.ColumnFormats);
        }

        if (autoFitColumns)
        {
            ApplyColumnSizing(
                worksheet,
                target,
                projection,
                projection.Values.GetLength(0),
                projection.Values.GetLength(1),
                ColumnSizingMode.AllDownloadedData,
                preserveExistingColumnWidths);
        }

        ApplyRowSizing(worksheet, target, rowSizingMode);
        ApplyRowOutcomes(worksheet, projection);
    }

    /// <summary>
    /// Applies Salesforce-derived formats to the full known result body before streamed
    /// page values are written. The row count deliberately comes from Salesforce's
    /// cursor total, rather than the (potentially larger) stale-body clear range.
    /// </summary>
    internal static void ApplyColumnFormatsToBody(
        Worksheet worksheet,
        SheetProjection projection,
        int rowCount)
    {
        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        if (projection is null)
        {
            throw new ArgumentNullException(nameof(projection));
        }

        var columnCount = projection.Values.GetLength(1);
        if (rowCount <= 0 || columnCount == 0)
        {
            return;
        }

        var target = CreateRange(
            worksheet,
            projection.StartRow,
            projection.StartColumn,
            rowCount,
            columnCount);
        SessionFlowTrace.Log(
            $"QueryTable: applying body column formats rows={rowCount} columns={columnCount}");
        ApplyColumnFormats(target, projection.ColumnFormats);
    }

    /// <summary>AutoFits the field-header row without considering body data.</summary>
    internal static void ApplyHeaderColumnSizing(Worksheet worksheet, SheetProjection projection)
    {
        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        if (projection is null)
        {
            throw new ArgumentNullException(nameof(projection));
        }

        var columnCount = projection.Values.GetLength(1);
        if (columnCount == 0)
        {
            return;
        }

        var headerRow = projection.StartRow > 1 ? projection.StartRow - 1 : projection.StartRow;
        CreateRange(worksheet, headerRow, projection.StartColumn, rowCount: 1, columnCount: columnCount).Columns.AutoFit();
    }

    /// <summary>Disables wrapping and restores standard height for a rectangular result body.</summary>
    internal static void ForceRowsToSingleLine(
        Worksheet worksheet,
        SheetProjection projection,
        int rowCount)
    {
        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        if (projection is null)
        {
            throw new ArgumentNullException(nameof(projection));
        }

        var columnCount = projection.Values.GetLength(1);
        if (rowCount <= 0 || columnCount == 0)
        {
            return;
        }

        var target = CreateRange(worksheet, projection.StartRow, projection.StartColumn, rowCount, columnCount);
        target.WrapText = false;
        target.RowHeight = worksheet.StandardHeight;
    }

    /// <summary>Applies layout once across all pages of a streamed query.</summary>
    public static void ApplyLayout(
        Worksheet worksheet,
        int startRow,
        int startColumn,
        int rowCount,
        int columnCount,
        AutomaticSizingMode automaticSizingMode)
    {
        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        if (rowCount <= 0 || columnCount <= 0 || automaticSizingMode == AutomaticSizingMode.None)
        {
            return;
        }

        var start = (Range)worksheet.Cells[startRow, startColumn];
        var end = (Range)worksheet.Cells[startRow + rowCount - 1, startColumn + columnCount - 1];
        var target = worksheet.Range[start, end];
        target.Columns.AutoFit();
        if (automaticSizingMode == AutomaticSizingMode.Both)
        {
            target.Rows.AutoFit();
            CapRowHeights(target, worksheet.StandardHeight * MaxRowHeightMultiplier);
        }
    }

    public static void ClearBody(Worksheet worksheet, ClearBodyRegion region)
    {
        var start = (Range)worksheet.Cells[region.StartRow, region.StartColumn];
        var end = (Range)worksheet.Cells[region.EndRow, region.EndColumn];
        var range = worksheet.Range[start, end];
        range.Clear();
    }

    private static Range WriteValues(Worksheet worksheet, SheetProjection projection)
    {
        var rows = projection.Values.GetLength(0);
        var cols = projection.Values.GetLength(1);
        var target = CreateRange(worksheet, projection.StartRow, projection.StartColumn, rows, cols);
        target.Value2 = projection.Values;
        return target;
    }

    private static Range CreateRange(
        Worksheet worksheet,
        int startRow,
        int startColumn,
        int rowCount,
        int columnCount)
    {
        var start = (Range)worksheet.Cells[startRow, startColumn];
        var end = (Range)worksheet.Cells[startRow + rowCount - 1, startColumn + columnCount - 1];
        return worksheet.Range[start, end];
    }

    private static void ApplyColumnFormats(Range target, IReadOnlyList<ColumnFormat>? formats)
    {
        if (formats is null || formats.Count == 0)
        {
            return;
        }

        // One NumberFormat assignment per column (not per cell) — NFR write path.
        var columnCount = Math.Min(formats.Count, target.Columns.Count);
        for (var c = 0; c < columnCount; c++)
        {
            var excelFormat = formats[c].ExcelFormat
                ?? DefaultExcelFormat(formats[c].Kind);
            if (excelFormat is null)
            {
                continue;
            }

            var column = (Range)target.Columns[c + 1];
            column.NumberFormat = excelFormat;
        }
    }

    private static string? DefaultExcelFormat(ColumnFormatKind kind) =>
        kind switch
        {
            ColumnFormatKind.Date => "yyyy-MM-dd",
            ColumnFormatKind.DateTime => "yyyy-MM-dd HH:mm:ss",
            ColumnFormatKind.Text => "@",
            _ => null,
        };

    private static void ApplyColumnSizing(
        Worksheet worksheet,
        Range target,
        SheetProjection projection,
        int rows,
        int cols,
        ColumnSizingMode columnSizingMode,
        bool preserveExistingColumnWidths = false)
    {
        if (columnSizingMode == ColumnSizingMode.HeadersOnly)
        {
            ApplyHeaderColumnSizing(worksheet, projection);
            return;
        }

        // Include the header row above the body when present so widths fit labels too.
        var widthStartRow = projection.StartRow > 1 ? projection.StartRow - 1 : projection.StartRow;
        var widthStart = (Range)worksheet.Cells[widthStartRow, projection.StartColumn];
        var widthEnd = (Range)worksheet.Cells[projection.StartRow + rows - 1, projection.StartColumn + cols - 1];
        var widthRange = worksheet.Range[widthStart, widthEnd];
        var priorWidths = preserveExistingColumnWidths ? CaptureColumnWidths(widthRange, cols) : null;
        widthRange.Columns.AutoFit();
        RestoreNarrowedColumns(widthRange, priorWidths);
    }

    private static void ApplyRowSizing(Worksheet worksheet, Range target, RowSizingMode rowSizingMode)
    {
        switch (rowSizingMode)
        {
            case RowSizingMode.FitEachPage:
                target.Rows.AutoFit();
                CapRowHeights(target, worksheet.StandardHeight * MaxRowHeightMultiplier);
                break;
            case RowSizingMode.ForceSingleLine:
                target.WrapText = false;
                target.RowHeight = worksheet.StandardHeight;
                break;
            case RowSizingMode.None:
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(rowSizingMode));
        }
    }

    private static double[] CaptureColumnWidths(Range range, int columnCount)
    {
        var widths = new double[columnCount];
        for (var c = 0; c < columnCount; c++)
        {
            var column = (Range)range.Columns[c + 1];
            widths[c] = Convert.ToDouble(column.ColumnWidth);
        }

        return widths;
    }

    private static void RestoreNarrowedColumns(Range range, double[]? priorWidths)
    {
        if (priorWidths is null)
        {
            return;
        }

        for (var c = 0; c < priorWidths.Length; c++)
        {
            var column = (Range)range.Columns[c + 1];
            if (Convert.ToDouble(column.ColumnWidth) < priorWidths[c])
            {
                column.ColumnWidth = priorWidths[c];
            }
        }
    }

    private static void CapRowHeights(Range target, double maxRowHeight)
    {
        // O(rows) COM — Excel has no bulk "max height" API; still far cheaper than per-cell.
        var rowCount = target.Rows.Count;
        for (var i = 1; i <= rowCount; i++)
        {
            var row = (Range)target.Rows[i];
            if (Convert.ToDouble(row.RowHeight) > maxRowHeight)
            {
                row.RowHeight = maxRowHeight;
            }
        }
    }

    private static void ApplyRowOutcomes(Worksheet worksheet, SheetProjection projection)
    {
        if (projection.RowOutcomes is null)
        {
            return;
        }

        foreach (var outcome in projection.RowOutcomes)
        {
            if (outcome.Succeeded)
            {
                continue;
            }

            var row = projection.StartRow + outcome.BodyRowIndex;
            var cols = projection.Values.GetLength(1);
            var highlightStart = (Range)worksheet.Cells[row, projection.StartColumn];
            var highlightEnd = (Range)worksheet.Cells[row, projection.StartColumn + cols - 1];
            var highlight = worksheet.Range[highlightStart, highlightEnd];
            highlight.Interior.Color = ColorTranslator.ToOle(Color.Orange);
        }
    }
}
