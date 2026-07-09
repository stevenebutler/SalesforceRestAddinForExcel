using System.Windows;
using System.Windows.Controls;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class OptionsWindow : Window
{
    private readonly CheckBox _useReference;
    private readonly CheckBox _noWarning;
    private readonly CheckBox _confirmLargeQuery;
    private readonly CheckBox _noQueryLimit;
    private readonly CheckBox _autoAssignRule;
    private readonly CheckBox _includeHiddenCells;
    private readonly TextBox _batchSize;
    private readonly System.Action? _clearCurrentInstanceCache;
    private readonly string? _instanceUrl;

    public OptionsWindow(
        ConnectorOptions current,
        System.Action? clearCurrentInstanceCache = null,
        string? instanceUrl = null)
    {
        _clearCurrentInstanceCache = clearCurrentInstanceCache;
        _instanceUrl = instanceUrl;

        Title = "Salesforce REST Add-in Options";
        Width = 520;
        SizeToContent = SizeToContent.Height;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.NoResize;

        var root = new StackPanel { Margin = new Thickness(16) };
        _useReference = MakeCheck("Use Reference Name/Id", current.UseReference);
        _noWarning = MakeCheck("Do not show warning dialogs before commencing operations.", current.NoWarning);
        _confirmLargeQuery = MakeCheck("Confirm Large Query Download", current.ConfirmLargeQuery);
        _noQueryLimit = MakeCheck("No Query Limit", current.NoQueryLimit);
        _autoAssignRule = MakeCheck("Enable Auto Assign Rule", current.AutoAssignRule);
        _includeHiddenCells = MakeCheck("Include Hidden Columns/Rows", current.IncludeHiddenCells);
        root.Children.Add(_useReference);
        root.Children.Add(_noWarning);
        root.Children.Add(_confirmLargeQuery);
        root.Children.Add(_noQueryLimit);
        root.Children.Add(_autoAssignRule);
        root.Children.Add(_includeHiddenCells);

        root.Children.Add(new TextBlock { Text = "Composite batch size", Margin = new Thickness(0, 12, 0, 4) });
        _batchSize = new TextBox { Text = current.CompositeBatchSize.ToString() };
        root.Children.Add(_batchSize);

        var clearCache = new Button
        {
            Content = "Clear metadata cache for this org",
            Width = 220,
            HorizontalAlignment = HorizontalAlignment.Left,
            Margin = new Thickness(0, 16, 0, 0),
            IsEnabled = _clearCurrentInstanceCache is not null && !string.IsNullOrWhiteSpace(_instanceUrl),
        };
        clearCache.Click += (_, _) => ClearMetadataCache();
        root.Children.Add(clearCache);
        if (string.IsNullOrWhiteSpace(_instanceUrl))
        {
            root.Children.Add(new TextBlock
            {
                Text = "Sign in to clear this org’s metadata cache.",
                Margin = new Thickness(0, 4, 0, 0),
                TextWrapping = TextWrapping.Wrap,
            });
        }

        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 16, 0, 0) };
        var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0) };
        var cancel = new Button { Content = "Cancel", Width = 80 };
        ok.Click += (_, _) => { DialogResult = true; Close(); };
        cancel.Click += (_, _) => { DialogResult = false; Close(); };
        buttons.Children.Add(ok);
        buttons.Children.Add(cancel);
        root.Children.Add(buttons);
        Content = root;
    }

    public ConnectorOptions GetOptions()
    {
        var batch = int.TryParse(_batchSize.Text, out var size) && size > 0 ? size : 200;
        return new ConnectorOptions
        {
            UseReference = _useReference.IsChecked == true,
            NoWarning = _noWarning.IsChecked == true,
            ConfirmLargeQuery = _confirmLargeQuery.IsChecked == true,
            NoQueryLimit = _noQueryLimit.IsChecked == true,
            AutoAssignRule = _autoAssignRule.IsChecked == true,
            IncludeHiddenCells = _includeHiddenCells.IsChecked == true,
            CompositeBatchSize = batch,
        };
    }

    private void ClearMetadataCache()
    {
        if (_clearCurrentInstanceCache is null || string.IsNullOrWhiteSpace(_instanceUrl))
        {
            MessageBox.Show(
                this,
                "Sign in to clear this org’s metadata cache.",
                Title,
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var confirm = MessageBox.Show(
            this,
            "Clear cached object lists and describe metadata for the current Salesforce org only? Other orgs and sandboxes are left unchanged. The next operation for this org will re-download metadata.",
            Title,
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        _clearCurrentInstanceCache();
        MessageBox.Show(
            this,
            "Metadata cache cleared for this org.",
            Title,
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private static CheckBox MakeCheck(string label, bool value) =>
        new()
        {
            Content = new TextBlock
            {
                Text = label,
                TextWrapping = TextWrapping.Wrap,
                Width = 470,
            },
            IsChecked = value,
            Margin = new Thickness(0, 0, 0, 8),
        };
}
