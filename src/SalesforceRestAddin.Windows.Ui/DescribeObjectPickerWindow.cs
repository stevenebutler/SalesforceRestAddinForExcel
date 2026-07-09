using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class DescribeObjectPickerWindow : Window
{
    private readonly ListBox _list;

    public DescribeObjectPickerWindow(IReadOnlyList<SObjectSummary> objects)
    {
        Title = "Describe Sforce Object";
        Width = 520;
        Height = 480;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        var filter = new TextBox { Margin = new Thickness(16, 16, 16, 8) };
        _list = UiListBox.Create(SelectionMode.Extended);
        _list.Margin = new Thickness(16, 0, 16, 8);

        var items = objects
            .Where(o => o.Queryable)
            .OrderBy(o => o.Label, StringComparer.OrdinalIgnoreCase)
            .ToList();
        foreach (var item in items)
        {
            _list.Items.Add(new SObjectListItem(item));
        }

        filter.TextChanged += (_, _) =>
        {
            var text = filter.Text.Trim();
            _list.Items.Clear();
            foreach (var item in items.Where(i =>
                         string.IsNullOrEmpty(text)
                         || i.Label.Contains(text, StringComparison.OrdinalIgnoreCase)
                         || i.Name.Contains(text, StringComparison.OrdinalIgnoreCase)))
            {
                _list.Items.Add(new SObjectListItem(item));
            }
        };

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(16),
        };
        var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0) };
        var cancel = new Button { Content = "Cancel", Width = 80 };
        ok.Click += (_, _) => { DialogResult = true; Close(); };
        cancel.Click += (_, _) => { DialogResult = false; Close(); };
        buttons.Children.Add(ok);
        buttons.Children.Add(cancel);

        var root = new DockPanel();
        DockPanel.SetDock(buttons, Dock.Bottom);
        DockPanel.SetDock(filter, Dock.Top);
        root.Children.Add(buttons);
        root.Children.Add(filter);
        root.Children.Add(_list);
        Content = root;
    }

    public IReadOnlyList<string> GetSelectedApiNames() =>
        _list.SelectedItems.Cast<SObjectListItem>().Select(i => i.Summary.Name).ToList();
}
