using System.Windows.Media;
using SalesforceRestAddin.Core.DataPlane;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>Display-only legend row sharing wizard bucket labels and fills.</summary>
internal sealed class LegendListItem
{
    public LegendListItem(int bucket)
    {
        DisplayText = WizardTableLayoutBuilder.BucketLegendLabel(bucket);
        var (r, g, b) = WizardTableLayoutBuilder.HeaderFillRgb(bucket);
        RowBackground = new SolidColorBrush(Color.FromRgb(r, g, b));
        RowBackground.Freeze();
    }

    public string DisplayText { get; }

    public Brush RowBackground { get; }

    public override string ToString() => DisplayText;
}
