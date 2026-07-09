using System.Linq;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Core.Tables;

public static class ForceTableBinder
{
    public static BindingValidationResult Bind(ForceTableSnapshot snapshot, SObjectDescribe describe)
    {
        if (snapshot is null)
        {
            throw new ArgumentNullException(nameof(snapshot));
        }

        if (describe is null)
        {
            throw new ArgumentNullException(nameof(describe));
        }

        var objectNameResult = ForceTableParser.ValidateObjectName(
            snapshot.ObjectApiName,
            snapshot.StartRow,
            snapshot.StartColumn);
        if (!objectNameResult.Succeeded && objectNameResult.Errors.Count > 0)
        {
            return objectNameResult;
        }

        var catalog = new FieldCatalog(describe);
        var columns = new List<BoundColumn>();
        var errors = new List<BindingValidationError>();
        var idColumnIndex = -1;

        for (var col = 0; col < snapshot.HeaderLabels.Length; col++)
        {
            var label = snapshot.HeaderLabels[col]?.ToString();
            if (string.IsNullOrWhiteSpace(label))
            {
                break;
            }

            var field = catalog.ResolveHeaderField(label, snapshot.HeaderApiNames[col]);
            if (field is null)
            {
                errors.Add(new BindingValidationError
                {
                    Cell = new CellRef(2, col + 1),
                    Message = $"Unknown field '{label}' in column {GetColumnLetter(col + 1)}.",
                });
                break;
            }

            columns.Add(new BoundColumn { ColumnIndex = col, Field = field });
            if (field.IsId)
            {
                idColumnIndex = col;
            }
        }

        if (idColumnIndex < 0)
        {
            errors.Add(new BindingValidationError
            {
                Cell = new CellRef(2, 1),
                Message = "No Id column found in the header row.",
            });
        }
        else if (idColumnIndex != 0)
        {
            errors.Add(new BindingValidationError
            {
                Cell = new CellRef(2, idColumnIndex + 1),
                Message = "Record Id must be the first field column (same column as the object name on row 1).",
            });
        }

        if (errors.Count > 0)
        {
            return new BindingValidationResult { Errors = errors };
        }

        var boundSnapshot = new ForceTableSnapshot
        {
            ObjectApiName = snapshot.ObjectApiName,
            CriteriaRow = TrimCriteria(snapshot.CriteriaRow, columns.Count),
            HeaderLabels = TrimHeaders(snapshot.HeaderLabels, columns.Count),
            HeaderApiNames = TrimApiNames(snapshot.HeaderApiNames, columns.Count),
            Body = TrimBody(snapshot.Body, columns.Count),
            IdColumnIndex = idColumnIndex,
            StartRow = snapshot.StartRow,
            StartColumn = snapshot.StartColumn,
            HiddenRowIndices = snapshot.HiddenRowIndices,
            HiddenColumnIndices = FilterHiddenColumns(snapshot.HiddenColumnIndices, columns.Count),
        };

        return new BindingValidationResult
        {
            Binding = new ForceTableBinding
            {
                Snapshot = boundSnapshot,
                Describe = describe,
                Catalog = catalog,
                Columns = columns,
            },
        };
    }

    private static object?[] TrimHeaders(object?[] headers, int columnCount)
    {
        if (headers.Length == columnCount)
        {
            return headers;
        }

        var trimmed = new object?[columnCount];
        Array.Copy(headers, trimmed, columnCount);
        return trimmed;
    }

    private static string?[] TrimApiNames(string?[] apiNames, int columnCount)
    {
        if (apiNames.Length == columnCount)
        {
            return apiNames;
        }

        var trimmed = new string?[columnCount];
        Array.Copy(apiNames, trimmed, Math.Min(apiNames.Length, columnCount));
        return trimmed;
    }

    private static object?[] TrimCriteria(object?[] criteria, int columnCount)
    {
        // Criteria are columns 2..N of the table (length ColumnCount - 1).
        var expected = Math.Max(0, columnCount - 1);
        if (criteria.Length == expected)
        {
            return criteria;
        }

        var trimmed = new object?[expected];
        Array.Copy(criteria, trimmed, Math.Min(criteria.Length, expected));
        return trimmed;
    }

    private static object?[,] TrimBody(object?[,] body, int columnCount)
    {
        var rows = body.GetLength(0);
        var cols = body.GetLength(1);
        if (cols == columnCount)
        {
            return body;
        }

        var trimmed = new object?[rows, columnCount];
        var copyCols = Math.Min(cols, columnCount);
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < copyCols; c++)
            {
                trimmed[r, c] = body[r, c];
            }
        }

        return trimmed;
    }

    private static IReadOnlyList<int>? FilterHiddenColumns(IReadOnlyList<int>? hidden, int columnCount)
    {
        if (hidden is null || hidden.Count == 0)
        {
            return hidden;
        }

        var filtered = hidden.Where(c => c >= 0 && c < columnCount).ToList();
        return filtered.Count == 0 ? null : filtered;
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
