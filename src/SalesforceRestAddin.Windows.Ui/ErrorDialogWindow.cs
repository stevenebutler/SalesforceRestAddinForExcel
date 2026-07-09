using System;
using System.Windows;
using System.Windows.Controls;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>
/// Scrollable error details for ribbon/COM failures — not a one-line MessageBox.
/// </summary>
public sealed class ErrorDialogWindow : Window
{
    public ErrorDialogWindow(string title, string details)
    {
        Title = title;
        Width = 560;
        Height = 320;
        MinWidth = 420;
        MinHeight = 200;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.CanResizeWithGrip;

        var root = new Grid { Margin = new Thickness(16) };
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var detailsBox = new TextBox
        {
            Text = details,
            IsReadOnly = true,
            TextWrapping = TextWrapping.Wrap,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            FontFamily = new System.Windows.Media.FontFamily("Consolas"),
            Margin = new Thickness(0, 0, 0, 12),
        };
        Grid.SetRow(detailsBox, 0);
        root.Children.Add(detailsBox);

        var ok = new Button
        {
            Content = "OK",
            Width = 80,
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        ok.Click += (_, _) =>
        {
            DialogResult = true;
            Close();
        };
        Grid.SetRow(ok, 1);
        root.Children.Add(ok);

        Content = root;
    }

    public static void Show(string title, string message) =>
        WpfUiThread.Run(() => ShowWindow(title, message));

    public static void Show(string title, string? operation, Exception exception) =>
        WpfUiThread.Run(() => ShowWindow(title, ExceptionDetailFormatter.Format(operation, exception)));

    private static void ShowWindow(string title, string details)
    {
        var window = new ErrorDialogWindow(title, details);
        window.ShowDialog();
    }
}
