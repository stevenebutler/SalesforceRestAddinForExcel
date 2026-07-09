using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace SalesforceRestAddin.Windows.Ui;

internal static class UiListBox
{
    private static readonly SolidColorBrush CustomFieldRowBrush =
        CreateFrozenBrush(0xD8, 0xD8, 0xD8);

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

    /// <summary>
    /// Field picker: tooltip = Salesforce type; custom fields get a darker row background.
    /// </summary>
    public static Style CreateFieldPickerItemStyle()
    {
        var style = CreatePaddedItemStyle();
        style.Setters.Add(new Setter(
            FrameworkElement.ToolTipProperty,
            new Binding(nameof(FieldListItem.ToolTipText))));

        var customTrigger = new DataTrigger
        {
            Binding = new Binding(nameof(FieldListItem.IsCustom)),
            Value = true,
        };
        customTrigger.Setters.Add(new Setter(Control.BackgroundProperty, CustomFieldRowBrush));
        style.Triggers.Add(customTrigger);
        return style;
    }

    private static SolidColorBrush CreateFrozenBrush(byte r, byte g, byte b)
    {
        var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
        brush.Freeze();
        return brush;
    }
}
