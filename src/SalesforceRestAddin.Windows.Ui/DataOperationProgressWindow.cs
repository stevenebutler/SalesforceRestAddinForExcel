using System;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class DataOperationProgressWindow : Window
{
    private readonly TextBlock _statusText;
    private readonly ProgressBar _progressBar;
    private readonly CancellationTokenSource _cts = new();

    public DataOperationProgressWindow(string title, bool allowCancel = true)
    {
        Title = title;
        Width = 420;
        SizeToContent = SizeToContent.Height;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.NoResize;

        var root = new StackPanel { Margin = new Thickness(16) };
        _statusText = new TextBlock
        {
            Text = "Starting...",
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
        };
        _progressBar = new ProgressBar
        {
            Height = 20,
            Minimum = 0,
            Maximum = 100,
            IsIndeterminate = true,
        };
        root.Children.Add(_statusText);
        root.Children.Add(_progressBar);

        if (allowCancel)
        {
            var cancel = new Button
            {
                Content = "Cancel",
                Width = 80,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0),
                IsCancel = true,
            };
            cancel.Click += (_, _) =>
            {
                _statusText.Text = "Cancelling...";
                _progressBar.IsIndeterminate = true;
                _cts.Cancel();
                // Keep the window open until the host finishes observing cancellation;
                // closing here would race with progress updates and hide feedback.
            };
            root.Children.Add(cancel);
        }

        Closing += (_, _) =>
        {
            if (!_cts.IsCancellationRequested)
            {
                _cts.Cancel();
            }
        };
        Closed += (_, _) =>
        {
            try
            {
                _cts.Dispose();
            }
            catch
            {
                // Already disposed.
            }
        };
        Content = root;
    }

    public CancellationToken CancellationToken => _cts.Token;

    public void SetStatus(string message) => _statusText.Text = message;

    /// <summary>
    /// Updates status and switches to a determinate bar when <paramref name="total"/> is known and positive.
    /// </summary>
    public void SetProgress(string message, int completed, int? total)
    {
        _statusText.Text = message;
        if (total is int t && t > 0)
        {
            _progressBar.IsIndeterminate = false;
            _progressBar.Minimum = 0;
            _progressBar.Maximum = t;
            _progressBar.Value = Math.Max(0, Math.Min(completed, t));
        }
        else
        {
            _progressBar.IsIndeterminate = true;
        }
    }
}
