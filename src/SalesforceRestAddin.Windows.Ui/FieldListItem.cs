using System.Collections.Generic;
using System.Text;
using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>
/// Wizard field-picker row: label + API name + flags; custom rows use a darker background.
/// </summary>
internal sealed class FieldListItem
{
    public FieldListItem(FieldDescriptor field)
    {
        Field = field;
        DisplayText = FormatDisplay(field);
        ToolTipText = FormatToolTip(field);
        IsCustom = WizardTableLayoutBuilder.IsCustomField(field);
    }

    public FieldDescriptor Field { get; }

    public string DisplayText { get; }

    public string ToolTipText { get; }

    public bool IsCustom { get; }

    public override string ToString() => DisplayText;

    internal static string FormatDisplay(FieldDescriptor field)
    {
        var flags = new List<string>();
        if (WizardTableLayoutBuilder.IsRequiredOnCreate(field))
        {
            flags.Add("req");
        }

        if (!field.Updateable)
        {
            flags.Add("ro");
        }

        if (field.IsReference)
        {
            flags.Add("lk");
        }

        var text = new StringBuilder();
        text.Append(field.Label);
        text.Append(" (");
        text.Append(field.Name);
        text.Append(')');
        foreach (var flag in flags)
        {
            text.Append(" [");
            text.Append(flag);
            text.Append(']');
        }

        return text.ToString();
    }

    internal static string FormatToolTip(FieldDescriptor field)
    {
        var tip = new StringBuilder(field.Type);
        if (field.Length is > 0)
        {
            tip.Append(", length ");
            tip.Append(field.Length.Value);
        }

        if (field.Precision is > 0)
        {
            tip.Append(", precision ");
            tip.Append(field.Precision.Value);
            if (field.Scale is >= 0)
            {
                tip.Append(", scale ");
                tip.Append(field.Scale.Value);
            }
        }

        return tip.ToString();
    }
}
