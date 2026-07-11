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
    public FieldListItem(FieldDescriptor field, int describeIndex)
    {
        Field = field;
        DescribeIndex = describeIndex;
        Bucket = WizardTableLayoutBuilder.GetFieldBucket(field);
        DisplayText = WizardTableLayoutBuilder.FormatFieldListDisplay(field);
        CategoryText = WizardTableLayoutBuilder.BucketDisplayLabel(Bucket);
        LabelText = field.Label;
        ApiName = field.Name;
        ToolTipText = FormatToolTip(field);
        var (r, g, b) = WizardTableLayoutBuilder.HeaderFillRgb(Bucket);
        CategoryBackground = new SolidColorBrush(Color.FromRgb(r, g, b));
        CategoryBackground.Freeze();
    }

    public FieldDescriptor Field { get; }

    public int DescribeIndex { get; }

    public int Bucket { get; }

    public string DisplayText { get; }

    public string CategoryText { get; }

    public string LabelText { get; }

    public string ApiName { get; }

    public string ToolTipText { get; }

    public Brush CategoryBackground { get; }

    public Brush CategoryForeground => Brushes.Black;

    public Brush RowBackground => CategoryBackground;

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
