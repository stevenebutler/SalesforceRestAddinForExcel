using System;
using System.Collections.Generic;
using System.Drawing;
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
        var headerRange = worksheet.Range[headerStart, headerEnd];
        headerRange.Value2 = labels;
        headerRange.Font.Bold = true;
        ApplyHeaderBucketFills(worksheet, headerRow, startColumn, fields);

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

    /// <summary>
    /// Shades contiguous header runs by wizard field bucket (FR-TQW-6).
    /// </summary>
    private static void ApplyHeaderBucketFills(
        Worksheet worksheet,
        int headerRow,
        int startColumn,
        IReadOnlyList<FieldDescriptor> fields)
    {
        var runStart = 0;
        var runBucket = WizardTableLayoutBuilder.GetFieldBucket(fields[0]);
        for (var i = 1; i <= fields.Count; i++)
        {
            var bucket = i < fields.Count
                ? WizardTableLayoutBuilder.GetFieldBucket(fields[i])
                : -1;
            if (bucket == runBucket)
            {
                continue;
            }

            var (r, g, b) = WizardTableLayoutBuilder.HeaderFillRgb(runBucket);
            var runStartCell = (Range)worksheet.Cells[headerRow, startColumn + runStart];
            var runEndCell = (Range)worksheet.Cells[headerRow, startColumn + i - 1];
            worksheet.Range[runStartCell, runEndCell].Interior.Color =
                ColorTranslator.ToOle(Color.FromArgb(r, g, b));

            runStart = i;
            runBucket = bucket;
        }
    }

    private static void WriteCriteriaRow(
        Worksheet worksheet,
        int startRow,
        int startColumn,
        IReadOnlyList<WizardCriteriaClause> criteria)
    {
        var values = WizardTableLayoutBuilder.BuildCriteriaRowValues(criteria);
        var valueCount = values.GetLength(1);
        if (valueCount == 0)
        {
            return;
        }

        var headerStart = (Range)worksheet.Cells[startRow, startColumn + 1];
        var headerEnd = (Range)worksheet.Cells[startRow, startColumn + valueCount];
        worksheet.Range[headerStart, headerEnd].Value2 = values;

        for (var i = 0; i < criteria.Count; i++)
        {
            var clause = criteria[i];
            var fieldCell = (Range)worksheet.Cells[startRow, startColumn + 1 + (i * 3)];
            ReplaceComment(fieldCell, clause.Field.Name);
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
