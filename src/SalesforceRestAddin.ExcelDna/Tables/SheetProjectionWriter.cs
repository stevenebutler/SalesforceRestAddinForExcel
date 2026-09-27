using System;
using System.Collections.Generic;
using System.Drawing;
using ExcelDna.Integration;
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
        RowSizingMode rowSizingMode,
        bool useNativeExcelWrites = false,
        IntPtr? nativeSheetId = null)
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

        var target = WriteValues(worksheet, projection, useNativeExcelWrites, nativeSheetId);
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
        bool preserveExistingColumnWidths,
        bool useNativeExcelWrites = false,
        IntPtr? nativeSheetId = null)
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

        var target = WriteValues(worksheet, projection, useNativeExcelWrites, nativeSheetId);
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

    private static Range WriteValues(
        Worksheet worksheet,
        SheetProjection projection,
        bool useNativeExcelWrites,
        IntPtr? nativeSheetId)
    {
        var rows = projection.Values.GetLength(0);
        var cols = projection.Values.GetLength(1);
        var target = CreateRange(worksheet, projection.StartRow, projection.StartColumn, rows, cols);
        LogFirstColumnIdentityStats(projection.Values);
        var useRowScopedWrites = AutoFilterWriteScope.HasActiveFilters(worksheet);
        if (useRowScopedWrites)
        {
            SessionFlowTrace.Log(
                $"Excel write: active AutoFilter detected; using row-scoped values rows={rows} columns={cols}");
        }

        if (useNativeExcelWrites)
        {
            WriteNativeValues(
                worksheet,
                projection,
                rows,
                cols,
                nativeSheetId,
                useRowScopedWrites);
        }
        else if (useRowScopedWrites)
        {
            WriteComValuesByRow(worksheet, projection, rows, cols);
        }
        else
        {
            target.Value2 = projection.Values;
        }

        return target;
    }

    private static void WriteNativeValues(
        Worksheet worksheet,
        SheetProjection projection,
        int rowCount,
        int columnCount,
        IntPtr? nativeSheetId,
        bool useRowScopedWrites)
    {
        var rowFirst = projection.StartRow - 1;
        var rowLast = rowFirst + rowCount - 1;
        var columnFirst = projection.StartColumn - 1;
        var columnLast = columnFirst + columnCount - 1;
        var sheetName = GetQualifiedSheetName(worksheet);

        SessionFlowTrace.Log(
            $"Native Excel write sheet={sheetName} " +
            $"rows={projection.StartRow}-{projection.StartRow + rowCount - 1} " +
            $"columns={projection.StartColumn}-{projection.StartColumn + columnCount - 1}");

        var failureMessage =
            $"Excel rejected the native worksheet write to {sheetName} " +
            $"at row {projection.StartRow}, column {projection.StartColumn}, " +
            $"size {rowCount}x{columnCount}.";
        if (useRowScopedWrites)
        {
            for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
            {
                var rowValues = CopyRow(projection.Values, rowOffset, columnCount);
                WriteNativeBlock(
                    rowFirst + rowOffset,
                    rowFirst + rowOffset,
                    columnFirst,
                    columnLast,
                    nativeSheetId ?? CaptureNativeSheetId(worksheet),
                    rowValues,
                    failureMessage);
            }

            return;
        }

        WriteNativeBlock(
            rowFirst,
            rowLast,
            columnFirst,
            columnLast,
            nativeSheetId ?? CaptureNativeSheetId(worksheet),
            projection.Values,
            failureMessage);
    }

    private static void WriteNativeBlock(
        int rowFirst,
        int rowLast,
        int columnFirst,
        int columnLast,
        IntPtr sheetId,
        object?[,] values,
        string failureMessage)
    {
        XlCall.XlReturn xlReturn;
        object result;
        try
        {
            var reference = new ExcelReference(
                rowFirst,
                rowLast,
                columnFirst,
                columnLast,
                sheetId);
            xlReturn = XlCall.TryExcel(
                XlCall.xlSet,
                out result,
                reference,
                values);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(failureMessage, ex);
        }

        if (xlReturn != XlCall.XlReturn.XlReturnSuccess || result is not bool succeeded || !succeeded)
        {
            throw new InvalidOperationException(
                $"{failureMessage} Excel C API result: {xlReturn}; " +
                $"xlSet result: {result?.ToString() ?? "<null>"}.");
        }
    }

    private static void WriteComValuesByRow(
        Worksheet worksheet,
        SheetProjection projection,
        int rowCount,
        int columnCount)
    {
        for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
        {
            var row = projection.StartRow + rowOffset;
            var target = CreateRange(worksheet, row, projection.StartColumn, 1, columnCount);
            target.Value2 = CopyRow(projection.Values, rowOffset, columnCount);
        }
    }

    private static object?[,] CopyRow(object?[,] values, int rowOffset, int columnCount)
    {
        var rowValues = new object?[1, columnCount];
        for (var column = 0; column < columnCount; column++)
        {
            rowValues[0, column] = values[rowOffset, column];
        }

        return rowValues;
    }

    /// <summary>
    /// Resolves a stable native sheet id while Excel is directly executing the initiating
    /// macro/COM call. Streamed pages reuse it inside the nested STA progress dispatcher.
    /// </summary>
    internal static IntPtr CaptureNativeSheetId(Worksheet worksheet)
    {
        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        // The caller captures this while the target worksheet is active and Excel is
        // directly executing the initiating macro/COM call. Avoid another xlSheetId
        // lookup later from the nested progress dispatcher.
        var xlReturn = XlCall.TryExcel(XlCall.xlSheetId, out var result);
        if (xlReturn == XlCall.XlReturn.XlReturnSuccess
            && result is ExcelReference reference
            && reference.SheetId != IntPtr.Zero)
        {
            SessionFlowTrace.Log(
                $"Native Excel sheet id captured for {GetQualifiedSheetName(worksheet)}");
            return reference.SheetId;
        }

        throw new InvalidOperationException(
            $"Excel rejected native sheet identification for {GetQualifiedSheetName(worksheet)}. " +
            $"Excel C API result: {xlReturn}; result type: {result?.GetType().FullName ?? "<null>"}.");
    }

    private static string GetQualifiedSheetName(Worksheet worksheet)
    {
        var workbook = (Workbook)worksheet.Parent;
        return $"'[{EscapeSheetReferencePart(workbook.Name)}]{EscapeSheetReferencePart(worksheet.Name)}'";
    }

    private static string EscapeSheetReferencePart(string value) => value.Replace("'", "''");

    private static void LogFirstColumnIdentityStats(object?[,] values)
    {
        var rows = values.GetLength(0);
        var distinct = new HashSet<string>(StringComparer.Ordinal);
        var nonEmpty = 0;
        for (var row = 0; row < rows; row++)
        {
            var text = values[row, 0]?.ToString();
            if (string.IsNullOrEmpty(text))
            {
                continue;
            }

            nonEmpty++;
            distinct.Add(text!);
        }

        SessionFlowTrace.Log(
            $"Excel write source firstColumnNonEmpty={nonEmpty} " +
            $"firstColumnDistinct={distinct.Count}");
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
