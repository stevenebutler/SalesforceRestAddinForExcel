using System.Windows;
using System.Windows.Controls;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class AboutWindow : Window
{
    public AboutWindow(string buildInfo, bool isLoggedIn)
    {
        Title = "About Salesforce REST Add-in for Excel";
        Width = 420;
        SizeToContent = SizeToContent.Height;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.NoResize;

        var root = new StackPanel { Margin = new Thickness(16) };
        root.Children.Add(new TextBlock
        {
            Text = "Salesforce REST Add-in for Excel",
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(0, 0, 0, 8),
        });
        root.Children.Add(new TextBlock { Text = $"Build: {buildInfo}", Margin = new Thickness(0, 0, 0, 4) });
        root.Children.Add(new TextBlock
        {
            Text = isLoggedIn ? "Session: signed in" : "Session: not signed in",
            Margin = new Thickness(0, 0, 0, 12),
        });

        var ok = new Button { Content = "OK", Width = 80, HorizontalAlignment = HorizontalAlignment.Right };
        ok.Click += (_, _) => { DialogResult = true; Close(); };
        root.Children.Add(ok);
        Content = root;
    }
}
