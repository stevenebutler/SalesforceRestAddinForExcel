using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Core.DataPlane;

public static class WizardTableLayoutBuilder
{
    public const int BucketCount = 8;

    /// <summary>
    /// FR-TQW-6: Id → Name → Required (standard/custom) → Standard → Custom → Read-only (standard/custom).
    /// Within a bucket, preserve describe field order.
    /// </summary>
    public static IReadOnlyList<FieldDescriptor> OrderFieldsForWizard(SObjectDescribe describe)
    {
        return describe.Fields
            .Select((field, index) => (Field: field, Index: index))
            .OrderBy(x => GetFieldBucket(x.Field))
            .ThenBy(x => x.Index)
            .Select(x => x.Field)
            .ToList();
    }

    public static string DefaultWhereClause() => "Id != null";

    /// <summary>
    /// Row 2 field headers aligned with row 1 object cell. Wizard layout places Id first.
    /// </summary>
    public static object?[] BuildHeaderRowLabels(IReadOnlyList<FieldDescriptor> fields) =>
        fields.Select(static f => (object?)f.Label).ToArray();

    /// <summary>Row 1 criteria cells written contiguously as field/operator/value triplets.</summary>
    public static object?[,] BuildCriteriaRowValues(IReadOnlyList<WizardCriteriaClause> criteria)
    {
        var values = new object?[1, criteria.Count * 3];
        for (var i = 0; i < criteria.Count; i++)
        {
            var clause = criteria[i];
            var column = i * 3;
            values[0, column] = clause.Field.Label;
            values[0, column + 1] = clause.Operator;
            values[0, column + 2] = clause.Value ?? string.Empty;
        }

        return values;
    }

    /// <summary>Required on create: non-nillable and createable (bugs.md #8 / FR-TQW-6).</summary>
    public static bool IsRequiredOnCreate(FieldDescriptor field) =>
        !field.Nillable && field.Createable;

    public static bool IsCustomField(FieldDescriptor field) =>
        field.Custom || field.Name.EndsWith("__c", StringComparison.Ordinal);

    /// <summary>Name field: describe <c>nameField</c>, else API name <c>Name</c>.</summary>
    public static bool IsNameField(FieldDescriptor field) =>
        field.NameField || string.Equals(field.Name, "Name", StringComparison.OrdinalIgnoreCase);

    /// <summary>Wizard field-list display: <c>{Label} ({ApiName})</c> — no flag suffixes.</summary>
    public static string FormatFieldListDisplay(FieldDescriptor field) =>
        $"{field.Label} ({field.Name})";

    /// <summary>
    /// Wizard header classification (first match wins):
    /// 0 Id, 1 Name, 2 Required (standard), 3 Required (custom),
    /// 4 Standard, 5 Custom, 6 Read-only (standard), 7 Read-only (custom).
    /// </summary>
    public static int GetFieldBucket(FieldDescriptor field)
    {
        if (field.IsId)
        {
            return 0;
        }

        if (IsNameField(field))
        {
            return 1;
        }

        if (IsRequiredOnCreate(field))
        {
            return IsCustomField(field) ? 3 : 2;
        }

        if (field.Updateable)
        {
            return IsCustomField(field) ? 5 : 4;
        }

        return IsCustomField(field) ? 7 : 6;
    }

    /// <summary>Legend / kind label for a sort bucket (shared with fills and field-list rows).</summary>
    public static string BucketLegendLabel(int bucket) =>
        bucket switch
        {
            0 => "Id",
            1 => "Name",
            2 => "Required (standard)",
            3 => "Required (custom)",
            4 => "Standard",
            5 => "Custom",
            6 => "Read-only (standard)",
            7 => "Read-only (custom)",
            _ => string.Empty,
        };

    public static string BucketDisplayLabel(int bucket) =>
        $"{bucket + 1} {BucketLegendLabel(bucket)}";

    /// <summary>
    /// Soft pastel fills for wizard header columns, field-list rows, and legend (readable with bold black text).
    /// </summary>
    public static (byte R, byte G, byte B) HeaderFillRgb(int bucket) =>
        bucket switch
        {
            0 => (189, 215, 238), // Id — light blue
            1 => (255, 230, 153), // Name — soft gold
            2 => (248, 203, 173), // Required (standard) — peach
            3 => (230, 160, 120), // Required (custom) — darker peach
            4 => (198, 239, 206), // Standard — mint
            5 => (226, 213, 241), // Custom — lavender
            6 => (217, 217, 217), // Read-only (standard) — gray
            7 => (160, 160, 160), // Read-only (custom) — darker gray
            _ => (255, 255, 255),
        };
}
