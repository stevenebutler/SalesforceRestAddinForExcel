using System;
using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Windows.Ui;
using Microsoft.Office.Interop.Excel;

namespace SalesforceRestAddin;

/// <summary>
/// Sole wait/pump for async work started on the Excel STA: modal WPF <c>ShowDialog</c>
/// (nested message loop), progress UI, Cancel, and StatusBar. Do not add a second pump
/// (<c>Dispatcher.PushFrame</c>, bare <c>GetResult()</c> on the STA, etc.).
/// </summary>
internal static class ExcelStaAsyncHost
{
    /// <summary>
    /// Progress + Cancel (query/update/describe and other long data-plane ops).
    /// </summary>
    public static T Run<T>(
        string title,
        Application excel,
        Func<CancellationToken, Action<OperationProgress>, Task<T>> work) =>
        RunCore(title, excel, allowCancel: true, work);

    /// <summary>
    /// Status-only wait (login, ListObjects, Describe binding). Cancel is hidden.
    /// </summary>
    public static T Run<T>(
        string title,
        Application excel,
        Func<CancellationToken, Action<string>, Task<T>> work) =>
        RunCore(
            title,
            excel,
            allowCancel: false,
            (ct, report) => work(ct, status => report(OperationProgress.Status(status))));

    /// <summary>Status-only wait when the work returns a non-generic <see cref="Task"/>.</summary>
    public static void Run(
        string title,
        Application excel,
        Func<CancellationToken, Action<string>, Task> work) =>
        Run(
            title,
            excel,
            async (ct, report) =>
            {
                await work(ct, report).ConfigureAwait(false);
                return 0;
            });

    private static T RunCore<T>(
        string title,
        Application excel,
        bool allowCancel,
        Func<CancellationToken, Action<OperationProgress>, Task<T>> work)
    {
        if (excel is null)
        {
            throw new ArgumentNullException(nameof(excel));
        }

        if (work is null)
        {
            throw new ArgumentNullException(nameof(work));
        }

        DataOperationProgressWindow? window = null;
        WpfUiThread.Run(() => window = new DataOperationProgressWindow(title, allowCancel));

        void Report(OperationProgress progress)
        {
            // WPF updates must stay on the dispatcher (ShowDialog is pumping it).
            WpfUiThread.Run(() =>
            {
                if (window is null)
                {
                    return;
                }

                if (progress.Completed is int completed)
                {
                    window.SetProgress(progress.Message, completed, progress.Total);
                }
                else
                {
                    window.SetStatus(progress.Message);
                }
            });

            // Never touch Excel COM from a thread-pool continuation (zombie EXCEL.EXE).
            var status = progress.Message;
            var truncated = status.Length > 128 ? status.Substring(0, 128) : status;
            ExcelComThread.BeginInvoke(() =>
            {
                try
                {
                    excel.StatusBar = truncated;
                }
                catch
                {
                    // Excel may reject status bar updates during shutdown.
                }
            });
        }

        try
        {
            var task = work(window!.CancellationToken, Report);
            WaitUntilClosed(window, task);
            // Task is complete after ShowDialog returns (or was already complete).
            return task.GetAwaiter().GetResult();
        }
        finally
        {
            WpfUiThread.Run(() =>
            {
                if (window is not null && window.IsVisible)
                {
                    window.Close();
                }
            });
            ExcelComThread.Invoke(() =>
            {
                try
                {
                    excel.StatusBar = false;
                }
                catch
                {
                    // Ignore status bar reset failures.
                }
            });
        }
    }

    /// <summary>
    /// Shows the progress window modally so WPF runs a nested message loop (paint + Cancel).
    /// Closing is marshalled onto the dispatcher when <paramref name="task"/> completes —
    /// never mutate dispatcher frame state from a thread-pool thread.
    /// </summary>
    private static void WaitUntilClosed<T>(DataOperationProgressWindow window, Task<T> task)
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher is null)
        {
            throw new InvalidOperationException("WPF dispatcher is not available for ExcelStaAsyncHost.");
        }

        if (task.IsCompleted)
        {
            return;
        }

        var closePosted = 0;

        void TryClose()
        {
            if (Interlocked.Exchange(ref closePosted, 1) != 0)
            {
                return;
            }

            try
            {
                if (window.IsVisible)
                {
                    window.Close();
                }
            }
            catch
            {
                // Window may already be closing with Excel.
            }
        }

        task.ContinueWith(
            _ =>
            {
                try
                {
                    if (dispatcher.CheckAccess())
                    {
                        TryClose();
                    }
                    else
                    {
                        dispatcher.BeginInvoke(new System.Action(TryClose));
                    }
                }
                catch (Exception ex)
                {
                    SessionFlowTrace.LogException("ExcelStaAsyncHost: failed to close progress window", ex);
                }
            },
            CancellationToken.None,
            TaskContinuationOptions.None,
            TaskScheduler.Default);

        // Task may finish before ShowDialog; do not call ShowDialog on a closed window.
        if (task.IsCompleted)
        {
            TryClose();
            return;
        }

        // Modal ShowDialog pumps input on this STA thread (Cancel works). Modeless Show +
        // GetResult freezes the dispatcher; PushFrame is forbidden (undocumented / re-entrant).
        if (dispatcher.CheckAccess())
        {
            window.ShowDialog();
        }
        else
        {
            dispatcher.Invoke(() => window.ShowDialog());
        }
    }
}
