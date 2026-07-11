using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class DescribeObjectPickerWindow : Window
{
    private readonly ObjectPickerControl _picker;

    public DescribeObjectPickerWindow(IReadOnlyList<SObjectSummary> objects)
    {
        Title = "Describe Sforce Object";
        Width = 720;
        Height = 480;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        _picker = new ObjectPickerControl(objects, SelectionMode.Extended);
        _picker.Margin = new Thickness(16, 16, 16, 0);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(16),
        };
        var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
        var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
        ok.Click += (_, _) => { DialogResult = true; Close(); };
        cancel.Click += (_, _) => { DialogResult = false; Close(); };
        buttons.Children.Add(ok);
        buttons.Children.Add(cancel);

        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        Grid.SetRow(_picker, 0);
        Grid.SetRow(buttons, 1);
        root.Children.Add(_picker);
        root.Children.Add(buttons);
        Content = root;
    }

    public IReadOnlyList<string> GetSelectedApiNames() =>
        _picker.SelectedApiNames;
}
