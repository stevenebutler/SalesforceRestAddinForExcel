using System;
using System.Collections.Generic;
using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Tables;
using Microsoft.Office.Interop.Excel;

namespace SalesforceRestAddin.Tables;

/// <summary>
/// Writes the ForceConnector table header block produced by the Table Query Wizard (rows 1–2).
/// </summary>
public static class WizardTableLayoutWriter
{
    public static void Apply(
        Worksheet worksheet,
        int startRow,
        int startColumn,
        SObjectDescribe describe,
        IReadOnlyList<FieldDescriptor> fields,
        IReadOnlyList<WizardCriteriaClause>? criteria = null)
    {
        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        if (describe is null)
        {
            throw new ArgumentNullException(nameof(describe));
        }

        if (fields is null || fields.Count == 0)
        {
            throw new ArgumentException("At least one field is required.", nameof(fields));
        }

        var objectCell = (Range)worksheet.Cells[startRow, startColumn];
        objectCell.Value2 = describe.Label;
        ReplaceComment(objectCell, describe.Name);

        if (criteria is { Count: > 0 })
        {
            WriteCriteriaRow(worksheet, startRow, startColumn, criteria);
        }

        var labels = new object[1, fields.Count];
        for (var i = 0; i < fields.Count; i++)
        {
            labels[0, i] = fields[i].Label;
        }

        var headerRow = startRow + 1;
        var headerStart = (Range)worksheet.Cells[headerRow, startColumn];
        var headerEnd = (Range)worksheet.Cells[headerRow, startColumn + fields.Count - 1];
        worksheet.Range[headerStart, headerEnd].Value2 = labels;

        for (var i = 0; i < fields.Count; i++)
        {
            var cell = (Range)worksheet.Cells[headerRow, startColumn + i];
            ReplaceComment(cell, BuildFieldComment(fields[i]));
        }

        // Blank separator column so multi-table discovery does not merge with neighbouring content.
        var separatorColumn = startColumn + fields.Count;
        var separatorTop = (Range)worksheet.Cells[startRow, separatorColumn];
        var separatorBottom = (Range)worksheet.Cells[headerRow, separatorColumn];
        worksheet.Range[separatorTop, separatorBottom].Clear();
    }

    private static void WriteCriteriaRow(
        Worksheet worksheet,
        int startRow,
        int startColumn,
        IReadOnlyList<WizardCriteriaClause> criteria)
    {
        var column = startColumn + 1;
        for (var i = 0; i < criteria.Count; i++)
        {
            var clause = criteria[i];
            var fieldCell = (Range)worksheet.Cells[startRow, column++];
            fieldCell.Value2 = clause.Field.Label;
            ReplaceComment(fieldCell, clause.Field.Name);

            ((Range)worksheet.Cells[startRow, column++]).Value2 = clause.Operator;
            ((Range)worksheet.Cells[startRow, column++]).Value2 = clause.Value ?? string.Empty;
            if (i < criteria.Count - 1)
            {
                ((Range)worksheet.Cells[startRow, column++]).Value2 = "and";
            }
        }
    }

    private static string BuildFieldComment(FieldDescriptor field)
    {
        var lines = new List<string> { $"API Name: {field.Name}" };
        if (!field.Updateable)
        {
            lines.Add("Read Only Field");
        }

        if (field.IsId)
        {
            lines.Add("Primary Object Identifier");
        }
        else if (!field.Createable)
        {
            lines.Add("Required on Insert");
        }

        lines.Add($"Type: {field.Type}");
        foreach (var value in field.PicklistValues)
        {
            lines.Add(value);
        }

        return string.Join(Environment.NewLine, lines);
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
