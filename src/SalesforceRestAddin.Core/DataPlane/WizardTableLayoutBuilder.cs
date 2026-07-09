using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Core.DataPlane;

public static class WizardTableLayoutBuilder
{
    /// <summary>
    /// FR-TQW-6: Id → required → Name → standard → custom → read-only.
    /// </summary>
    public static IReadOnlyList<FieldDescriptor> OrderFieldsForWizard(SObjectDescribe describe)
    {
        return describe.Fields
            .OrderBy(WizardFieldBucket)
            .ThenBy(f => f.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static string DefaultWhereClause() => "Id != null";

    /// <summary>
    /// Row 2 field headers aligned with row 1 object cell — Id shares the anchor column (legacy layout).
    /// </summary>
    public static object?[] BuildHeaderRowLabels(IReadOnlyList<FieldDescriptor> fields) =>
        fields.Select(static f => (object?)f.Label).ToArray();

    /// <summary>Required on create: non-nillable and createable (bugs.md #8 / FR-TQW-6).</summary>
    public static bool IsRequiredOnCreate(FieldDescriptor field) =>
        !field.Nillable && field.Createable;

    public static bool IsCustomField(FieldDescriptor field) =>
        field.Custom || field.Name.EndsWith("__c", StringComparison.Ordinal);

    private static int WizardFieldBucket(FieldDescriptor field)
    {
        if (field.IsId)
        {
            return 0;
        }

        if (IsRequiredOnCreate(field))
        {
            return 1;
        }

        if (string.Equals(field.Name, "Name", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }

        if (!field.Updateable)
        {
            return 5;
        }

        if (IsCustomField(field))
        {
            return 4;
        }

        return 3;
    }
}
