using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Navigation;
using SalesforceRestAddin.Core;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class AboutWindow : Window
{
    public AboutWindow(string buildInfo, bool isLoggedIn)
    {
        Title = "About " + ProductBranding.ProductName;
        Width = 460;
        SizeToContent = SizeToContent.Height;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.NoResize;

        var root = new StackPanel { Margin = new Thickness(16) };
        root.Children.Add(new TextBlock
        {
            Text = ProductBranding.ProductName,
            FontWeight = FontWeights.Bold,
            FontSize = 16,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
        });
        root.Children.Add(new TextBlock
        {
            Text = ProductBranding.Tagline,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 12),
        });
        root.Children.Add(new TextBlock { Text = $"Build: {buildInfo}", Margin = new Thickness(0, 0, 0, 4) });
        root.Children.Add(new TextBlock
        {
            Text = isLoggedIn ? "Session: signed in" : "Session: not signed in",
            Margin = new Thickness(0, 0, 0, 12),
        });
        root.Children.Add(new TextBlock
        {
            Text = ProductBranding.Author,
            Margin = new Thickness(0, 0, 0, 2),
        });
        root.Children.Add(new TextBlock
        {
            Text = ProductBranding.Copyright,
            Margin = new Thickness(0, 0, 0, 12),
        });

        var repoBlock = new TextBlock { Margin = new Thickness(0, 0, 0, 16), TextWrapping = TextWrapping.Wrap };
        repoBlock.Inlines.Add("GitHub: ");
        var link = new Hyperlink(new Run(ProductBranding.RepositoryDisplayName))
        {
            NavigateUri = new Uri(ProductBranding.RepositoryUrl),
        };
        link.RequestNavigate += OnRepositoryLinkClick;
        repoBlock.Inlines.Add(link);
        root.Children.Add(repoBlock);

        var ok = new Button { Content = "OK", Width = 80, HorizontalAlignment = HorizontalAlignment.Right };
        ok.Click += (_, _) => { DialogResult = true; Close(); };
        root.Children.Add(ok);
        Content = root;
    }

    private static void OnRepositoryLinkClick(object sender, RequestNavigateEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = e.Uri.AbsoluteUri,
                UseShellExecute = true,
            });
        }
        catch
        {
            // Browser launch can fail on locked-down hosts; ignore.
        }

        e.Handled = true;
    }
}
