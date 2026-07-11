using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace SalesforceRestAddin.Windows.Ui;

internal static class UiListBox
{
    public static ListBox Create(SelectionMode selectionMode = SelectionMode.Single)
    {
        return new ListBox
        {
            SelectionMode = selectionMode,
            BorderBrush = Brushes.LightGray,
            BorderThickness = new Thickness(1),
            ItemContainerStyle = CreatePaddedItemStyle(),
        };
    }

    public static void MakeScrollable(ListBox list)
    {
        ScrollViewer.SetVerticalScrollBarVisibility(list, ScrollBarVisibility.Auto);
        list.VerticalAlignment = VerticalAlignment.Stretch;
    }

    public static Style CreatePaddedItemStyle()
    {
        var style = new Style(typeof(ListBoxItem));
        style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(8, 4, 8, 4)));
        return style;
    }

    public static Style CreateStretchingPaddedItemStyle()
    {
        var style = CreatePaddedItemStyle();
        style.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch));
        style.Setters.Add(new Setter(Control.VerticalContentAlignmentProperty, VerticalAlignment.Stretch));
        return style;
    }

    /// <summary>
    /// Field picker: tooltip = Salesforce type; row background from bucket fill.
    /// </summary>
    public static Style CreateFieldPickerItemStyle()
    {
        var style = CreatePaddedItemStyle();
        style.Setters.Add(new Setter(
            FrameworkElement.ToolTipProperty,
            new Binding(nameof(FieldListItem.ToolTipText))));
        style.Setters.Add(new Setter(
            Control.BackgroundProperty,
            new Binding(nameof(FieldListItem.RowBackground))));
        return style;
    }

    /// <summary>
    /// Legend rows: same bucket fills; display-only (not selectable / not hit-testable).
    /// </summary>
    public static Style CreateLegendItemStyle()
    {
        var style = CreatePaddedItemStyle();
        style.Setters.Add(new Setter(
            Control.BackgroundProperty,
            new Binding(nameof(LegendListItem.RowBackground))));
        style.Setters.Add(new Setter(UIElement.IsHitTestVisibleProperty, false));
        style.Setters.Add(new Setter(UIElement.FocusableProperty, false));
        return style;
    }
}
