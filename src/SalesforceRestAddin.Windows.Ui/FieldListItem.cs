using System.Text;
using System.Windows.Media;
using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>
/// Wizard field-picker row: label + API name; row fill from the shared eight-bucket palette.
/// </summary>
internal sealed class FieldListItem
{
    public FieldListItem(FieldDescriptor field)
    {
        Field = field;
        DisplayText = WizardTableLayoutBuilder.FormatFieldListDisplay(field);
        ToolTipText = FormatToolTip(field);
        var (r, g, b) = WizardTableLayoutBuilder.HeaderFillRgb(WizardTableLayoutBuilder.GetFieldBucket(field));
        RowBackground = new SolidColorBrush(Color.FromRgb(r, g, b));
        RowBackground.Freeze();
    }

    public FieldDescriptor Field { get; }

    public string DisplayText { get; }

    public string ToolTipText { get; }

    public Brush RowBackground { get; }

    public override string ToString() => DisplayText;

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
