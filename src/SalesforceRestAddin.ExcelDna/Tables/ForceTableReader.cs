using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using Microsoft.Office.Interop.Excel;

namespace SalesforceRestAddin.Tables;

public static class ForceTableReader
{
    /// <summary>
    /// Captures the ForceConnector table containing <paramref name="activeCell"/>.
    /// Tables are delimited by blank columns in the header row; the object name sits
    /// in the cell above the first field column of that header block.
    /// </summary>
    public static ForceTableSnapshot Capture(Application application, Worksheet worksheet, Range activeCell)
    {
        if (application is null)
        {
            throw new ArgumentNullException(nameof(application));
        }

        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        if (activeCell is null)
        {
            throw new ArgumentNullException(nameof(activeCell));
        }

        var anchor = (Range)activeCell.Cells[1, 1];
        var activeColumn = anchor.Column;
        var activeRow = anchor.Row;

        // Prefer CurrentRegion to locate the header row of this contiguous block, then
        // trim left/right by blank header columns so side-by-side tables stay isolated.
        var region = anchor.CurrentRegion;
        var regionStartRow = region.Row;
        var regionStartColumn = region.Column;
        var regionColumnCount = region.Columns.Count;
        var regionRowCount = region.Rows.Count;

        // Header is row 2 of a ForceConnector table (object on row 1). If the region is
        // only the body (user selected deep in data), CurrentRegion still includes headers
        // when contiguous; otherwise fall back to activeRow - 1 / activeRow - 2.
        var headerSheetRow = regionRowCount >= 2 ? regionStartRow + 1 : Math.Max(1, activeRow - 1);
        var objectSheetRow = headerSheetRow - 1;
        if (objectSheetRow < 1)
        {
            objectSheetRow = 1;
            headerSheetRow = 2;
        }

        var headerValueStopwatch = Stopwatch.StartNew();
        var headerMap = ReadHeaderRowMap(worksheet, headerSheetRow, regionStartColumn, regionColumnCount);
        headerValueStopwatch.Stop();
        // Also probe a few columns left of the region in case CurrentRegion started mid-table
        // after a blank body cell — rare, but left-scan needs those headers.
        ExtendHeaderMapLeft(worksheet, headerSheetRow, headerMap, regionStartColumn, activeColumn);

        var bounds = TableBoundaryResolver.Resolve(headerMap, activeColumn);
        if (!bounds.Succeeded)
        {
            throw new InvalidOperationException(bounds.ErrorMessage ?? "Could not locate a ForceConnector table.");
        }

        var startColumn = bounds.StartColumn;
        var columnCount = bounds.ColumnCount;
        var startRow = objectSheetRow;

        var objectCell = (Range)worksheet.Cells[startRow, startColumn];
        var objectApiName = ReadObjectApiName(objectCell);
        if (string.IsNullOrWhiteSpace(objectApiName) || objectApiName.Contains(' '))
        {
            var address = objectCell.Address[false, false];
            throw new InvalidOperationException(
                $"Could not locate an object name in cell {address}. " +
                "The entity name must appear above the first field column of the table.");
        }

        var table = worksheet.Range[
            worksheet.Cells[startRow, startColumn],
            worksheet.Cells[Math.Max(startRow + 1, regionStartRow + regionRowCount - 1), startColumn + columnCount - 1]];

        var headerLabels = ReadRow(table, 2, columnCount);
        var headerCommentStopwatch = Stopwatch.StartNew();
        var headerApiNames = ReadHeaderApiNames(table, columnCount);
        headerCommentStopwatch.Stop();
        SessionFlowTrace.Log(
            $"ForceTableReader: header row read values={headerValueStopwatch.ElapsedMilliseconds}ms " +
            $"comments={headerCommentStopwatch.ElapsedMilliseconds}ms");
        var bodyRowCount = Math.Max(0, table.Rows.Count - 2);
        var body = ReadBody(table, bodyRowCount, columnCount);
        var criteriaRow = ReadCriteriaRow(table, columnCount);

        return new ForceTableSnapshot
        {
            ObjectApiName = objectApiName,
            CriteriaRow = criteriaRow,
            HeaderLabels = headerLabels,
            HeaderApiNames = headerApiNames,
            Body = body,
            StartRow = startRow,
            StartColumn = startColumn,
            HiddenRowIndices = ReadHiddenRows(worksheet, startRow, bodyRowCount),
            HiddenColumnIndices = ReadHiddenColumns(worksheet, startColumn, columnCount),
        };
    }

    private static Dictionary<int, object?> ReadHeaderRowMap(
        Worksheet worksheet,
        int headerSheetRow,
        int regionStartColumn,
        int regionColumnCount)
    {
        var map = new Dictionary<int, object?>();
        if (regionColumnCount <= 0)
        {
            return map;
        }

        var start = (Range)worksheet.Cells[headerSheetRow, regionStartColumn];
        var end = (Range)worksheet.Cells[headerSheetRow, regionStartColumn + regionColumnCount - 1];
        var rowRange = worksheet.Range[start, end];
        var raw = rowRange.Value2;

        if (raw is object[,] block)
        {
            for (var c = 0; c < regionColumnCount; c++)
            {
                map[regionStartColumn + c] = block[1, c + 1];
            }
        }
        else
        {
            // Single cell
            map[regionStartColumn] = raw;
        }

        return map;
    }

    private static void ExtendHeaderMapLeft(
        Worksheet worksheet,
        int headerSheetRow,
        Dictionary<int, object?> map,
        int regionStartColumn,
        int activeColumn)
    {
        // Walk left from the region start while headers are non-blank so a mid-table
        // CurrentRegion still finds the true Id/object column.
        for (var col = regionStartColumn - 1; col >= 1; col--)
        {
            var value = ((Range)worksheet.Cells[headerSheetRow, col]).Value2;
            if (TableBoundaryResolver.IsBlank(value))
            {
                break;
            }

            map[col] = value;
        }

        // Ensure the active column is present even if outside the initial region read.
        if (!map.ContainsKey(activeColumn))
        {
            map[activeColumn] = ((Range)worksheet.Cells[headerSheetRow, activeColumn]).Value2;
        }
    }

    private static string ReadObjectApiName(Range startCell)
    {
        var comment = startCell.Comment?.Text() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(comment))
        {
            var firstLine = comment.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .FirstOrDefault() ?? string.Empty;
            if (firstLine.StartsWith("API Name:", StringComparison.OrdinalIgnoreCase))
            {
                return firstLine.Substring("API Name:".Length).Trim();
            }

            if (!comment.Contains(' '))
            {
                return comment.Trim();
            }
        }

        return startCell.Value2?.ToString()?.Trim() ?? string.Empty;
    }

    private static object?[] ReadRow(Range table, int rowIndex, int columnCount)
    {
        var values = new object?[columnCount];
        for (var col = 1; col <= columnCount; col++)
        {
            values[col - 1] = ((Range)table.Cells[rowIndex, col]).Value2;
        }

        return values;
    }

    private static string?[] ReadHeaderApiNames(Range table, int columnCount)
    {
        var apiNames = new string?[columnCount];
        for (var col = 1; col <= columnCount; col++)
        {
            var cell = (Range)table.Cells[2, col];
            var comment = cell.Comment?.Text() ?? string.Empty;
            var firstLine = comment.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;
            if (firstLine.StartsWith("API Name:", StringComparison.OrdinalIgnoreCase))
            {
                apiNames[col - 1] = firstLine.Substring("API Name:".Length).Trim();
            }
        }

        return apiNames;
    }

    private static object?[,] ReadBody(Range table, int bodyRowCount, int columnCount)
    {
        if (bodyRowCount == 0)
        {
            return new object?[0, columnCount];
        }

        var start = (Range)table.Cells[3, 1];
        var end = (Range)table.Cells[2 + bodyRowCount, columnCount];
        var block = table.Worksheet.Range[start, end];
        var raw = block.Value2;
        if (raw is object[,] matrix)
        {
            return NormalizeBlock(matrix, bodyRowCount, columnCount);
        }

        var single = new object?[bodyRowCount, columnCount];
        single[0, 0] = raw;
        return single;
    }

    private static object?[] ReadCriteriaRow(Range table, int columnCount)
    {
        var criteria = new object?[Math.Max(0, columnCount - 1)];
        for (var col = 2; col <= columnCount; col++)
        {
            criteria[col - 2] = ((Range)table.Cells[1, col]).Value2;
        }

        return criteria;
    }

    private static object?[,] NormalizeBlock(object[,] raw, int rows, int columns)
    {
        var result = new object?[rows, columns];
        if (raw.GetLength(0) == 1 && raw.GetLength(1) == 1)
        {
            result[0, 0] = raw[1, 1];
            return result;
        }

        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                result[r, c] = raw[r + 1, c + 1];
            }
        }

        return result;
    }

    private static IReadOnlyList<int>? ReadHiddenRows(Worksheet worksheet, int startRow, int bodyRowCount)
    {
        var hidden = new List<int>();
        for (var r = 0; r < bodyRowCount; r++)
        {
            var row = (Range)worksheet.Rows[startRow + 2 + r];
            if ((bool)row.Hidden)
            {
                hidden.Add(r);
            }
        }

        return hidden.Count == 0 ? null : hidden;
    }

    private static IReadOnlyList<int>? ReadHiddenColumns(Worksheet worksheet, int startColumn, int columnCount)
    {
        var hidden = new List<int>();
        for (var c = 0; c < columnCount; c++)
        {
            var column = (Range)worksheet.Columns[startColumn + c];
            if ((bool)column.Hidden)
            {
                hidden.Add(c);
            }
        }

        return hidden.Count == 0 ? null : hidden;
    }
}
