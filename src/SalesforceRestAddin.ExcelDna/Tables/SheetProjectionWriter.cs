using System;
using System.Drawing;
using SalesforceRestAddin.Core.Tables;
using Microsoft.Office.Interop.Excel;

namespace SalesforceRestAddin.Tables;

public static class SheetProjectionWriter
{
    /// <summary>Legacy cap: long text must not expand a row beyond ~3× the sheet standard height.</summary>
    public const double MaxRowHeightMultiplier = 3d;

    public static void Apply(Worksheet worksheet, SheetProjection projection)
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

        var start = (Range)worksheet.Cells[projection.StartRow, projection.StartColumn];
        var end = (Range)worksheet.Cells[projection.StartRow + rows - 1, projection.StartColumn + cols - 1];
        var target = worksheet.Range[start, end];
        target.Value2 = projection.Values;
        ApplyColumnFormats(target, projection.ColumnFormats);
        ApplyColumnAndRowLayout(worksheet, target, projection, rows, cols);
        ApplyRowOutcomes(worksheet, projection);
    }

    public static void ClearBody(Worksheet worksheet, ClearBodyRegion region)
    {
        var start = (Range)worksheet.Cells[region.StartRow, region.StartColumn];
        var end = (Range)worksheet.Cells[region.EndRow, region.EndColumn];
        var range = worksheet.Range[start, end];
        range.Clear();
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

    /// <summary>
    /// Column AutoFit (header + body when a header row sits above the projection) and
    /// row AutoFit with a 3× standard-height cap — range/column scoped, not per-cell.
    /// </summary>
    private static void ApplyColumnAndRowLayout(
        Worksheet worksheet,
        Range target,
        SheetProjection projection,
        int rows,
        int cols)
    {
        // Include the header row above the body when present so widths fit labels too.
        var widthStartRow = projection.StartRow > 1 ? projection.StartRow - 1 : projection.StartRow;
        var widthStart = (Range)worksheet.Cells[widthStartRow, projection.StartColumn];
        var widthEnd = (Range)worksheet.Cells[projection.StartRow + rows - 1, projection.StartColumn + cols - 1];
        var widthRange = worksheet.Range[widthStart, widthEnd];
        widthRange.Columns.AutoFit();

        target.Rows.AutoFit();
        CapRowHeights(target, worksheet.StandardHeight * MaxRowHeightMultiplier);
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
