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
            $"ForceTableReader: header row read cells={columnCount} " +
            $"values={headerValueStopwatch.ElapsedMilliseconds}ms " +
            $"notes={headerCommentStopwatch.ElapsedMilliseconds}ms");
        var bodyRowCount = Math.Max(0, table.Rows.Count - 2);
        var body = ReadBody(table, bodyRowCount, columnCount);
        var criteriaRow = ReadCriteriaRow(table, columnCount);
        var criteriaReferenceIds = ReadCriteriaReferenceIds(worksheet, criteriaRow);
        var hiddenRows = ReadHiddenRows(worksheet, startRow, bodyRowCount);
        var hiddenColumns = ReadHiddenColumns(worksheet, startColumn, columnCount);
        SessionFlowTrace.Log(
            $"ForceTableReader: visibility bodyRows={bodyRowCount} columns={columnCount} " +
            $"hiddenBodyRows={FormatHiddenRows(hiddenRows, startRow)} " +
            $"hiddenColumns={FormatHiddenColumns(hiddenColumns, startColumn)}");

        return new ForceTableSnapshot
        {
            ObjectApiName = objectApiName,
            CriteriaRow = criteriaRow,
            CriteriaReferenceIds = criteriaReferenceIds,
            HeaderLabels = headerLabels,
            HeaderApiNames = headerApiNames,
            Body = body,
            StartRow = startRow,
            StartColumn = startColumn,
            HiddenRowIndices = hiddenRows,
            HiddenColumnIndices = hiddenColumns,
        };
    }

    internal static ForceTableSnapshot CaptureShell(Application application, Worksheet worksheet, Range activeCell)
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

        var region = anchor.CurrentRegion;
        var regionStartRow = region.Row;
        var regionStartColumn = region.Column;
        var regionColumnCount = region.Columns.Count;
        var regionRowCount = region.Rows.Count;

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
            $"ForceTableReader: header row read cells={columnCount} " +
            $"values={headerValueStopwatch.ElapsedMilliseconds}ms " +
            $"notes={headerCommentStopwatch.ElapsedMilliseconds}ms");
        var bodyRowCount = Math.Max(0, table.Rows.Count - 2);
        var body = ReadBody(table, bodyRowCount, columnCount);
        var hiddenRows = ReadHiddenRows(worksheet, startRow, bodyRowCount);
        var hiddenColumns = ReadHiddenColumns(worksheet, startColumn, columnCount);
        SessionFlowTrace.Log(
            $"ForceTableReader: visibility bodyRows={bodyRowCount} columns={columnCount} " +
            $"hiddenBodyRows={FormatHiddenRows(hiddenRows, startRow)} " +
            $"hiddenColumns={FormatHiddenColumns(hiddenColumns, startColumn)}");

        return new ForceTableSnapshot
        {
            ObjectApiName = objectApiName,
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = headerLabels,
            HeaderApiNames = headerApiNames,
            Body = body,
            StartRow = startRow,
            StartColumn = startColumn,
            HiddenRowIndices = hiddenRows,
            HiddenColumnIndices = hiddenColumns,
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

    internal static object?[] ReadCriteriaRowSlice(Worksheet worksheet, int row, int startColumn, int width)
    {
        var endColumn = startColumn + Math.Max(0, width - 1);
        var start = (Range)worksheet.Cells[row, startColumn];
        var end = (Range)worksheet.Cells[row, endColumn];
        var rowRange = worksheet.Range[start, end];
        var raw = rowRange.Value2;

        if (raw is object[,] block)
        {
            var cells = new object?[width];
            for (var c = 0; c < width; c++)
            {
                cells[c] = block[1, c + 1];
            }

            return cells;
        }

        var normalized = new object?[width];
        normalized[0] = raw;
        return normalized;
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

    internal static IReadOnlyDictionary<int, IReadOnlyList<string>> ReadCriteriaReferenceIds(
        Worksheet worksheet,
        object?[] criteria)
    {
        var result = new Dictionary<int, IReadOnlyList<string>>();
        for (var i = 0; i + 2 < criteria.Length; i += 3)
        {
            var operation = criteria[i + 1]?.ToString()?.Trim();
            if (!string.Equals(operation, "in", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var rangeName = criteria[i + 2]?.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(rangeName))
            {
                continue;
            }

            try
            {
                var source = worksheet.Range[rangeName];
                result[i + 2] = ReadSalesforceIds(source);
            }
            catch
            {
                // Core reports a criteria-cell error when this is not a valid range/name.
            }
        }

        return result;
    }

    private static IReadOnlyList<string> ReadSalesforceIds(Range source)
    {
        var ids = new List<string>();
        var raw = source.Value2;
        if (raw is object[,] values)
        {
            for (var row = 1; row <= values.GetLength(0); row++)
            {
                for (var column = 1; column <= values.GetLength(1); column++)
                {
                    AddSalesforceId(ids, values[row, column]);
                }
            }
        }
        else
        {
            AddSalesforceId(ids, raw);
        }

        return ids;
    }

    private static void AddSalesforceId(List<string> ids, object? value)
    {
        var id = value?.ToString()?.Trim();
        if (id is null || !IsSalesforceId(id) || ids.Contains(id, StringComparer.OrdinalIgnoreCase))
        {
            return;
        }

        ids.Add(id);
    }

    private static bool IsSalesforceId(string value)
    {
        if (value.Length is not 15 and not 18)
        {
            return false;
        }

        return value.All(char.IsLetterOrDigit);
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

    private static string FormatHiddenRows(IReadOnlyList<int>? hiddenRows, int startRow) =>
        FormatHiddenIndices(hiddenRows, index => $"body={index}/sheet={startRow + 2 + index}");

    private static string FormatHiddenColumns(IReadOnlyList<int>? hiddenColumns, int startColumn) =>
        FormatHiddenIndices(hiddenColumns, index => $"table={index}/sheet={startColumn + index}");

    private static string FormatHiddenIndices(IReadOnlyList<int>? indices, Func<int, string> format)
    {
        if (indices is null || indices.Count == 0)
        {
            return "none";
        }

        const int maximumLoggedIndices = 20;
        var displayed = indices.Take(maximumLoggedIndices).Select(format);
        var suffix = indices.Count > maximumLoggedIndices ? ", …" : string.Empty;
        return $"count={indices.Count} [{string.Join(", ", displayed)}{suffix}]";
    }
}
