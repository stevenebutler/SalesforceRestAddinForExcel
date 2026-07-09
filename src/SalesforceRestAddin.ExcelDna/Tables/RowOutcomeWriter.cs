using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using SalesforceRestAddin.Core.Tables;
using Microsoft.Office.Interop.Excel;

namespace SalesforceRestAddin.Tables;

public static class RowOutcomeWriter
{
    public static void Apply(
        Worksheet worksheet,
        ForceTableBinding binding,
        IReadOnlyList<RowOutcome> outcomes)
    {
        if (worksheet is null || outcomes.Count == 0)
        {
            return;
        }

        var startRow = binding.Snapshot.StartRow;
        var startColumn = binding.Snapshot.StartColumn;
        var idColumn = startColumn + binding.IdColumnIndex;

        var idUpdates = outcomes
            .Where(o => o.Succeeded && (o.IdWriteback is not null || o.IdDisplayOverride is not null))
            .OrderBy(o => o.BodyRowIndex)
            .ToList();

        if (idUpdates.Count > 0)
        {
            ApplyIdColumnUpdates(worksheet, startRow, idColumn, idUpdates);
        }

        var endColumn = startColumn + Math.Max(1, binding.Snapshot.ColumnCount) - 1;
        foreach (var outcome in outcomes.Where(o => !o.Succeeded))
        {
            var excelRow = startRow + 2 + outcome.BodyRowIndex;
            // Colour only this table's columns — not the entire worksheet row (multi-table sheets).
            var highlightStart = (Range)worksheet.Cells[excelRow, startColumn];
            var highlightEnd = (Range)worksheet.Cells[excelRow, endColumn];
            var highlight = worksheet.Range[highlightStart, highlightEnd];
            highlight.Interior.Color = ColorTranslator.ToOle(Color.Orange);

            var commentCell = (Range)worksheet.Cells[excelRow, idColumn];
            ReplaceComment(commentCell, BuildErrorComment(outcome));
        }
    }

    private static void ApplyIdColumnUpdates(
        Worksheet worksheet,
        int tableStartRow,
        int idColumn,
        IReadOnlyList<RowOutcome> updates)
    {
        var values = new object[updates.Count, 1];
        for (var i = 0; i < updates.Count; i++)
        {
            values[i, 0] = updates[i].IdWriteback ?? updates[i].IdDisplayOverride ?? string.Empty;
        }

        var firstExcelRow = tableStartRow + 2 + updates[0].BodyRowIndex;
        var lastExcelRow = tableStartRow + 2 + updates[updates.Count - 1].BodyRowIndex;
        var start = (Range)worksheet.Cells[firstExcelRow, idColumn];
        var end = (Range)worksheet.Cells[lastExcelRow, idColumn];
        var target = worksheet.Range[start, end];

        if (updates.Count == 1 || !IsContiguous(updates))
        {
            for (var i = 0; i < updates.Count; i++)
            {
                var row = tableStartRow + 2 + updates[i].BodyRowIndex;
                ((Range)worksheet.Cells[row, idColumn]).Value2 = values[i, 0];
            }

            return;
        }

        target.Value2 = values;
    }

    private static bool IsContiguous(IReadOnlyList<RowOutcome> updates)
    {
        for (var i = 1; i < updates.Count; i++)
        {
            if (updates[i].BodyRowIndex != updates[i - 1].BodyRowIndex + 1)
            {
                return false;
            }
        }

        return true;
    }

    private static string BuildErrorComment(RowOutcome outcome)
    {
        var title = "Operation Failed";
        if (outcome.ErrorMessages.Count == 0)
        {
            return title;
        }

        return title + Environment.NewLine + string.Join(Environment.NewLine, outcome.ErrorMessages);
    }

    private static void ReplaceComment(Range cell, string text)
    {
        if (cell.Comment is not null)
        {
            cell.Comment.Delete();
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        cell.AddComment(text);
        if (cell.Comment is not null)
        {
            cell.Comment.Shape.TextFrame.AutoSize = true;
        }
    }
}
