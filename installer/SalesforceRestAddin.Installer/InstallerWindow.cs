using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace SalesforceRestAddin.Installer;

internal sealed class InstallerWindow : Window
{
    private readonly InstallerService _service = new();
    private readonly TextBlock _statusText;
    private readonly TextBlock _sourceValue;
    private readonly TextBlock _bitnessValue;
    private readonly TextBlock _assetValue;
    private readonly TextBlock _installFolderValue;
    private readonly TextBlock _installedBuildValue;
    private readonly TextBlock _availableVersionValue;
    private readonly TextBlock _forceConnectorValue;
    private readonly ProgressBar _progressBar;
    private readonly Button _installButton;
    private readonly Button _removeButton;
    private readonly Button _refreshButton;
    private readonly Button _uninstallForceConnectorButton;
    private readonly StringBuilder _log = new();
    private InstallerState _state = new();

    public InstallerWindow(IEnumerable<string> args)
    {
        Title = "Salesforce REST Add-in Installer";
        Width = 780;
        MinWidth = 740;
        Height = 430;
        MinHeight = 430;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.CanResizeWithGrip;
        Icon = LoadExecutableIcon();

        var root = new Grid { Margin = new Thickness(18) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var header = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Thickness(0, 0, 0, 4),
        };
        header.Children.Add(new Image
        {
            Source = new BitmapImage(new Uri("pack://application:,,,/Images/install.png", UriKind.Absolute)),
            Width = 42,
            Height = 42,
            Stretch = Stretch.Uniform,
            Margin = new Thickness(0, 0, 10, 0),
        });
        var title = new TextBlock
        {
            Text = InstallerBranding.ProductName,
            FontSize = 22,
            FontWeight = FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center,
        };
        header.Children.Add(title);
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        var subtitle = new TextBlock
        {
            Text = "Install or update the Salesforce REST Add-in for Excel for this Windows user.",
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 14),
        };
        Grid.SetRow(subtitle, 1);
        root.Children.Add(subtitle);

        _statusText = new TextBlock
        {
            Text = "Starting...",
            FontWeight = FontWeights.SemiBold,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 10),
        };
        Grid.SetRow(_statusText, 2);
        root.Children.Add(_statusText);

        var detailsBorder = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(189, 215, 238)),
            BorderThickness = new Thickness(1),
            Background = new SolidColorBrush(Color.FromRgb(247, 251, 255)),
            Padding = new Thickness(12),
            Margin = new Thickness(0, 0, 0, 14),
        };
        var details = new Grid();
        details.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(130) });
        details.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        for (var index = 0; index < 7; index++)
        {
            details.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }

        _sourceValue = AddDetail(details, 0, "Install source:");
        _bitnessValue = AddDetail(details, 1, "Excel bitness:");
        _assetValue = AddDetail(details, 2, "Target XLL:");
        _installFolderValue = AddDetail(details, 3, "Install folder:");
        _installedBuildValue = AddDetail(details, 4, "Installed version:");
        _availableVersionValue = AddDetail(details, 5, "Available version:");
        _forceConnectorValue = AddDetail(details, 6, "ForceConnector:");
        detailsBorder.Child = details;
        Grid.SetRow(detailsBorder, 3);
        root.Children.Add(detailsBorder);

        _progressBar = new ProgressBar
        {
            Height = 14,
            Minimum = 0,
            Maximum = 100,
            Visibility = Visibility.Collapsed,
            Margin = new Thickness(0, 0, 0, 10),
        };
        Grid.SetRow(_progressBar, 4);
        root.Children.Add(_progressBar);

        var actions = new Grid();
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actions.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        _installButton = CreateButton("Install / Update", 118);
        _installButton.Click += async (_, _) => await InstallAsync();
        Grid.SetColumn(_installButton, 0);
        actions.Children.Add(_installButton);

        _removeButton = CreateButton("Remove", 78);
        _removeButton.Click += async (_, _) => await RemoveAsync();
        Grid.SetColumn(_removeButton, 1);
        actions.Children.Add(_removeButton);

        _uninstallForceConnectorButton = CreateButton("Uninstall ForceConnector", 164);
        _uninstallForceConnectorButton.Click += async (_, _) => await UninstallForceConnectorAsync();
        Grid.SetColumn(_uninstallForceConnectorButton, 2);
        actions.Children.Add(_uninstallForceConnectorButton);

        var logsButton = CreateButton("Logs", 70);
        logsButton.Click += (_, _) => ShowLogs();
        Grid.SetColumn(logsButton, 4);
        actions.Children.Add(logsButton);

        _refreshButton = CreateButton("Check latest", 100);
        _refreshButton.Click += async (_, _) => await RefreshAsync();
        Grid.SetColumn(_refreshButton, 5);
        actions.Children.Add(_refreshButton);

        var closeButton = CreateButton("Close", 70);
        closeButton.IsCancel = true;
        closeButton.Click += (_, _) => Close();
        Grid.SetColumn(closeButton, 6);
        actions.Children.Add(closeButton);

        Grid.SetRow(actions, 5);
        root.Children.Add(actions);
        Content = root;

        Loaded += async (_, _) => await RefreshAsync();
    }

    private static TextBlock AddDetail(Grid grid, int row, string label)
    {
        var labelBlock = new TextBlock
        {
            Text = label,
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 0, 12, 7),
        };
        Grid.SetRow(labelBlock, row);
        Grid.SetColumn(labelBlock, 0);
        grid.Children.Add(labelBlock);

        var valueBlock = new TextBlock
        {
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 7),
        };
        Grid.SetRow(valueBlock, row);
        Grid.SetColumn(valueBlock, 1);
        grid.Children.Add(valueBlock);
        return valueBlock;
    }

    private static Button CreateButton(string content, double minWidth) => new()
    {
        Content = content,
        MinWidth = minWidth,
        Margin = new Thickness(0, 0, 8, 0),
        Padding = new Thickness(7, 3, 7, 3),
    };

    private static ImageSource LoadExecutableIcon()
    {
        using var icon = System.Drawing.Icon.ExtractAssociatedIcon(typeof(InstallerWindow).Assembly.Location)
            ?? throw new InvalidOperationException("The installer executable does not contain an application icon.");
        var image = Imaging.CreateBitmapSourceFromHIcon(
            icon.Handle,
            Int32Rect.Empty,
            BitmapSizeOptions.FromEmptyOptions());
        image.Freeze();
        return image;
    }

    private async Task RefreshAsync()
    {
        SetBusy(true, "Checking current state...");
        try
        {
            _state = await _service.LoadStateAsync(AppendLogFromWorker);
            RenderState(_state);
        }
        catch (Exception ex)
        {
            SetStatus("Failed to check installer state.");
            AppendLog(ExceptionFormatter.Format("Installer refresh", ex));
        }
        finally
        {
            SetBusy(false, null);
        }
    }

    private async Task InstallAsync()
    {
        SetBusy(true, "Installing...");
        try
        {
            _state = await _service.LoadStateAsync(AppendLogFromWorker);
            if (_state.ForceConnectorUninstallEntries.Count > 0)
            {
                var entry = ChooseForceConnectorEntry(_state.ForceConnectorUninstallEntries);
                if (entry is null)
                {
                    return;
                }

                var confirmation = MessageBox.Show(
                    this,
                    "ForceConnector may conflict with the Salesforce REST Add-in for Excel.\n\n" +
                    "Install / Update will run the registered ForceConnector Windows uninstaller first. " +
                    "After it completes successfully, this installation will continue.",
                    "Uninstall ForceConnector first",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Warning);
                if (confirmation != MessageBoxResult.OK)
                {
                    return;
                }

                SetStatus("Uninstalling ForceConnector...");
                var uninstallResult = await _service.RunForceConnectorUninstallerAsync(entry, AppendLogFromWorker);
                if (!uninstallResult.Succeeded)
                {
                    throw new InvalidOperationException($"The ForceConnector uninstaller exited with code {uninstallResult.ExitCode}. Installation was not continued.");
                }

                _state = await _service.LoadStateAsync(AppendLogFromWorker);
                if (_state.ForceConnectorUninstallEntries.Count > 0)
                {
                    throw new InvalidOperationException("ForceConnector is still registered after its uninstaller completed. Installation was not continued.");
                }
            }

            var result = await _service.InstallAsync(
                progress => Dispatcher.Invoke(() => SetStatus(progress)),
                AppendLogFromWorker);
            _state = await _service.LoadStateAsync(AppendLogFromWorker);
            RenderState(_state);
            MessageBox.Show(
                this,
                $"Installed {result.AssetName} to:\n{result.XllPath}\n\nSource: {result.Source}\n\nRestart Excel fully to load the add-in.",
                "Install complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            var detail = ExceptionFormatter.Format("Install", ex);
            AppendLog(detail);
            MessageBox.Show(this, detail, "Install failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false, null);
        }
    }

    private async Task RemoveAsync()
    {
        SetBusy(true, "Removing...");
        try
        {
            var result = await _service.RemoveAsync(AppendLogFromWorker);
            _state = await _service.LoadStateAsync(AppendLogFromWorker);
            RenderState(_state);
            MessageBox.Show(this, result.Message, "Remove complete", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            var detail = ExceptionFormatter.Format("Remove", ex);
            AppendLog(detail);
            MessageBox.Show(this, detail, "Remove failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false, null);
        }
    }

    private async Task UninstallForceConnectorAsync()
    {
        var entry = ChooseForceConnectorEntry(_state.ForceConnectorUninstallEntries);
        if (entry is null)
        {
            return;
        }

        var confirmation = MessageBox.Show(
            this,
            $"Start the registered Windows uninstaller for:\n\n{entry.Description}\n\nWindows may request elevation. Complete the uninstaller, then restart Excel.",
            "Uninstall ForceConnector",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Warning);
        if (confirmation != MessageBoxResult.OK)
        {
            return;
        }

        SetBusy(true, "Starting ForceConnector uninstaller...");
        try
        {
            SetStatus("Uninstalling ForceConnector...");
            var result = await _service.RunForceConnectorUninstallerAsync(entry, AppendLogFromWorker);
            if (!result.Succeeded)
            {
                throw new InvalidOperationException($"The ForceConnector uninstaller exited with code {result.ExitCode}.");
            }
            MessageBox.Show(
                this,
                "ForceConnector was uninstalled. Restart Excel before using the Salesforce REST Add-in.",
                "ForceConnector uninstalled",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            var detail = ExceptionFormatter.Format("Uninstall ForceConnector", ex);
            AppendLog(detail);
            MessageBox.Show(this, detail, "Unable to start uninstaller", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false, null);
        }
    }

    private ForceConnectorUninstallEntry? ChooseForceConnectorEntry(IReadOnlyList<ForceConnectorUninstallEntry> entries)
    {
        if (entries.Count == 0)
        {
            return null;
        }

        if (entries.Count == 1)
        {
            return entries[0];
        }

        var dialog = new Window
        {
            Title = "Select ForceConnector installation",
            Owner = this,
            Width = 560,
            Height = 280,
            MinWidth = 460,
            MinHeight = 220,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
        };
        var root = new Grid { Margin = new Thickness(14) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        var prompt = new TextBlock { Text = "Choose the ForceConnector installation to uninstall:", Margin = new Thickness(0, 0, 0, 8) };
        Grid.SetRow(prompt, 0);
        root.Children.Add(prompt);
        var list = new ListBox { ItemsSource = entries, DisplayMemberPath = nameof(ForceConnectorUninstallEntry.Description), SelectedIndex = 0 };
        Grid.SetRow(list, 1);
        root.Children.Add(list);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
        var cancel = CreateButton("Cancel", 80);
        cancel.IsCancel = true;
        cancel.Click += (_, _) => dialog.DialogResult = false;
        var select = CreateButton("Select", 80);
        select.Click += (_, _) => dialog.DialogResult = true;
        buttons.Children.Add(cancel);
        buttons.Children.Add(select);
        Grid.SetRow(buttons, 2);
        root.Children.Add(buttons);
        dialog.Content = root;
        return dialog.ShowDialog() == true ? list.SelectedItem as ForceConnectorUninstallEntry : null;
    }

    private void SetBusy(bool busy, string? status)
    {
        _installButton.IsEnabled = !busy;
        _removeButton.IsEnabled = !busy;
        _refreshButton.IsEnabled = !busy;
        _uninstallForceConnectorButton.IsEnabled = !busy && _state.ForceConnectorUninstallEntries.Count > 0;
        _progressBar.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        _progressBar.IsIndeterminate = busy;
        if (status is not null)
        {
            SetStatus(status);
        }
    }

    private void SetStatus(string status) => _statusText.Text = status;

    private void RenderState(InstallerState state)
    {
        var forceConnectorDetected = state.ForceConnectorUninstallEntries.Count > 0;
        _statusText.Text = forceConnectorDetected
            ? "ForceConnector may conflict with the Salesforce REST Add-in for Excel; uninstall it to avoid conflicts."
            : state.Status;
        _statusText.Foreground = forceConnectorDetected ? Brushes.DarkOrange : SystemColors.ControlTextBrush;
        _sourceValue.Text = state.BundledAssetPath is not null
            ? $"Bundled local package (offline-ready): {state.BundledAssetPath}"
            : state.GitHubCheckError is null
                ? "GitHub latest release"
                : "GitHub unavailable — no bundled local XLL found";
        _bitnessValue.Text = state.ExcelBitness;
        _assetValue.Text = state.AssetName;
        _installFolderValue.Text = state.InstallDirectory;
        _installedBuildValue.Text = state.InstalledVersion ?? (state.IsInstalled ? "Unknown (older XLL)" : "Not installed");
        _availableVersionValue.Text = state.AvailableVersion ?? "Unavailable";
        _forceConnectorValue.Text = state.ForceConnectorUninstallEntries.Count switch
        {
            0 => "Not detected",
            1 => $"Detected — {state.ForceConnectorUninstallEntries[0].Description}. Install / Update will uninstall it first.",
            _ => $"Detected — {state.ForceConnectorUninstallEntries.Count} Windows installations found. Install / Update will ask which one to uninstall first.",
        };
        _uninstallForceConnectorButton.Visibility = state.ForceConnectorUninstallEntries.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        _uninstallForceConnectorButton.IsEnabled = state.ForceConnectorUninstallEntries.Count > 0;

        AppendLog($"State: {state.Status}");
        if (state.BundledAssetPath is not null)
        {
            AppendLog("Source: bundled local package; GitHub was not checked.");
        }
        else if (state.GitHubCheckError is not null)
        {
            AppendLog("GitHub check unavailable: " + state.GitHubCheckError);
        }
    }

    private void AppendLogFromWorker(string message) => Dispatcher.Invoke(() => AppendLog(message));

    private void AppendLog(string message)
    {
        _log.Append('[').Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("] ").AppendLine(message);
    }

    private void ShowLogs()
    {
        var dialog = new Window
        {
            Title = "Installer logs",
            Owner = this,
            Width = 760,
            Height = 500,
            MinWidth = 560,
            MinHeight = 320,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
        };
        var root = new Grid { Margin = new Thickness(14) };
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        var text = new TextBox
        {
            Text = _log.ToString(),
            IsReadOnly = true,
            AcceptsReturn = true,
            TextWrapping = TextWrapping.NoWrap,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            FontFamily = new FontFamily("Consolas"),
            FontSize = 12,
            Padding = new Thickness(8),
        };
        Grid.SetRow(text, 0);
        root.Children.Add(text);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
        var copy = CreateButton("Copy", 80);
        copy.Click += (_, _) =>
        {
            text.SelectAll();
            Clipboard.SetText(text.Text);
        };
        var close = CreateButton("Close", 80);
        close.IsCancel = true;
        close.Click += (_, _) => dialog.Close();
        buttons.Children.Add(copy);
        buttons.Children.Add(close);
        Grid.SetRow(buttons, 1);
        root.Children.Add(buttons);
        dialog.Content = root;
        dialog.ShowDialog();
    }
}

internal static class InstallerBranding
{
    public const string ProductName = "Salesforce REST Add-in for Excel";
    public const string Repository = "stevenebutler/SalesforceRestAddinForExcel";
    public const string ReleasesUrl = "https://github.com/stevenebutler/SalesforceRestAddinForExcel/releases/latest";
}
